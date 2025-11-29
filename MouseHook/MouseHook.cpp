#include <windows.h>
#include <shlobj.h>
#include <shldisp.h>
#include <shlwapi.h>
#include <exdisp.h>  // For IShellWindows
#include <comdef.h>   // For _bstr_t, _variant_t
#include <atlbase.h> // 使用 CComPtr 简化 COM 内存管理
#include <oleauto.h>
#include <set>
#include <string>
#include <vector>
#include <iostream>
#include <algorithm>

#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Shlwapi.lib")

HHOOK g_mouseHook;
bool g_isDragging = false;
POINT g_dragStartPos = { 0, 0 };
const int DRAG_THRESHOLD = 3; // 拖动阈值（像素）

void ExtractFileInfoFromDropClipboard();
void CheckOtherDataFormats(IDataObject* pDataObject);
void ExtractFileTypeInfo(const std::wstring& filePath);


//----------------------------------------------------------------------------------------

class FileDetector {
private:
	static const std::set<std::wstring> targetExtensions;
	static const int SWC_DESKTOP = 8;
	static const int SWFO_NEEDDISPATCH = 1;

	struct ShellWindowInfo {
		HWND hwnd;
		bool isDesktop;
	};

	static ShellWindowInfo FindShellParent(HWND hWnd) {
		ShellWindowInfo result = { nullptr, false };
		HWND current = hWnd;

		while (current != nullptr) {
			wchar_t className[256];
			GetClassName(current, className, sizeof(className) / sizeof(wchar_t));

			// Check if it's a regular explorer window
			if (wcscmp(className, L"CabinetWClass") == 0) {
				result.hwnd = current;
				result.isDesktop = false;
				return result;
			}

			// Check if it's desktop
			if (wcscmp(className, L"Progman") == 0 || wcscmp(className, L"WorkerW") == 0) {
				result.hwnd = current;
				result.isDesktop = true;
				return result;
			}

			current = GetParent(current);
		}

		// Fallback: check for desktop handle
		HWND progman = FindWindow(L"Progman", nullptr);
		if (progman != nullptr) {
			// Simplified desktop detection
		}

		return result;
	}

	static bool CheckExplorerWindow(IShellWindows* shellWindows, HWND targetHwnd) {
		long count;
		if (FAILED(shellWindows->get_Count(&count))) {
			return false;
		}

		for (long i = 0; i < count; i++) {
			VARIANT index;
			index.vt = VT_I4;
			index.lVal = i;

			IDispatch* dispatch = nullptr;
			if (FAILED(shellWindows->Item(index, &dispatch)) || !dispatch) {
				continue;
			}

			IWebBrowser2* browser = nullptr;
			if (FAILED(dispatch->QueryInterface(IID_IWebBrowser2, (void**)&browser))) {
				dispatch->Release();
				continue;
			}

			HWND windowHwnd = nullptr;
			if (SUCCEEDED(browser->get_HWND((LONG_PTR*)&windowHwnd))) {
				if (windowHwnd == targetHwnd) {
					bool result = HasValidSelection(browser);
					browser->Release();
					dispatch->Release();
					return result;
				}
			}

			browser->Release();
			dispatch->Release();
		}

		return false;
	}

	static bool CheckDesktopSelection(IShellWindows* shellWindows) {
		IShellBrowser* shellBrowser = nullptr;
		VARIANT empty = { VT_EMPTY };
		long hwnd = 0;

		// Get desktop window
		IDispatch* desktopDispatch = nullptr;
		if (FAILED(shellWindows->FindWindowSW(&empty, &empty, SWC_DESKTOP, &hwnd,
			SWFO_NEEDDISPATCH, &desktopDispatch)) ||
			!desktopDispatch) {
			return false;
		}

		IWebBrowser2* desktopBrowser = nullptr;
		if (FAILED(desktopDispatch->QueryInterface(IID_IWebBrowser2, (void**)&desktopBrowser))) {
			desktopDispatch->Release();
			return false;
		}

		bool result = HasValidSelection(desktopBrowser);
		desktopBrowser->Release();
		desktopDispatch->Release();
		return result;
	}

	static bool HasValidSelection(IWebBrowser2* browser) {
		IDispatch* documentDispatch = nullptr;
		if (FAILED(browser->get_Document(&documentDispatch)) || !documentDispatch) {
			return false;
		}

		IShellFolderViewDual* folderView = nullptr;
		if (FAILED(documentDispatch->QueryInterface(IID_IShellFolderViewDual, (void**)&folderView))) {
			documentDispatch->Release();
			return false;
		}

		FolderItems* selectionDispatch = nullptr;
		if (FAILED(folderView->SelectedItems(&selectionDispatch)) || !selectionDispatch) {
			folderView->Release();
			documentDispatch->Release();
			return false;
		}

		FolderItems* selectedItems = nullptr;
		if (FAILED(selectionDispatch->QueryInterface(IID_FolderItems, (void**)&selectedItems))) {
			selectionDispatch->Release();
			folderView->Release();
			documentDispatch->Release();
			return false;
		}

		long itemCount = 0;
		if (FAILED(selectedItems->get_Count(&itemCount)) || itemCount == 0) {
			selectedItems->Release();
			selectionDispatch->Release();
			folderView->Release();
			documentDispatch->Release();
			return false;
		}

		bool foundValidFile = false;
		for (long i = 0; i < itemCount; i++) {
			VARIANT itemIndex;
			itemIndex.vt = VT_I4;
			itemIndex.lVal = i;

			FolderItem* item = nullptr;
			if (FAILED(selectedItems->Item(itemIndex, &item)) || !item) {
				continue;
			}

			BSTR pathBstr = nullptr;
			if (SUCCEEDED(item->get_Path(&pathBstr)) && pathBstr) {
				std::wstring path(pathBstr);
				SysFreeString(pathBstr);

				// Check if it's a directory
				DWORD attrs = GetFileAttributes(path.c_str());
				if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY)) {
					item->Release();
					continue;
				}

				// Check extension
				std::wstring ext = PathFindExtension(path.c_str());
				if (targetExtensions.find(ext) != targetExtensions.end()) {
					foundValidFile = true;
					item->Release();
					break;
				}
			}
			item->Release();
		}

		selectedItems->Release();
		selectionDispatch->Release();
		folderView->Release();
		documentDispatch->Release();
		return foundValidFile;
	}

public:
	static bool IsDraggingSupportedFile() {
		try {
			// Initialize COM
			if (FAILED(CoInitialize(nullptr))) {
				return false;
			}

			// Get cursor position
			POINT mousePos;
			if (!GetCursorPos(&mousePos)) {
				CoUninitialize();
				return false;
			}

			// Get window under cursor
			HWND targetHwnd = WindowFromPoint(mousePos);
			if (!targetHwnd) {
				CoUninitialize();
				return false;
			}

			// Find shell parent
			ShellWindowInfo shellInfo = FindShellParent(targetHwnd);
			if (!shellInfo.hwnd) {
				CoUninitialize();
				return false;
			}

			// Create Shell.Application
			IShellWindows* shellWindows = nullptr;
			if (FAILED(CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL,
				IID_IShellWindows, (void**)&shellWindows)) ||
				!shellWindows) {
				CoUninitialize();
				return false;
			}

			bool result = false;
			if (shellInfo.isDesktop) {
				result = CheckDesktopSelection(shellWindows);
			}
			else {
				result = CheckExplorerWindow(shellWindows, shellInfo.hwnd);
			}

			shellWindows->Release();
			CoUninitialize();
			return result;

		}
		catch (...) {
			CoUninitialize();
			return false;
		}
	}
};

// Initialize static member
const std::set<std::wstring> FileDetector::targetExtensions = {
	L".txt", L".csv", L".log", L".xml", L".json", L".cs", L".html",
	L".md", L".xaml", L".py", L".java", L".c", L".cpp"
};

class FileDetector1
{
private:
	static bool IsTargetExtension(const std::wstring& path)
	{
		// 定义目标后缀名
		static const std::set<std::wstring> targetExtensions = {
			L".txt", L".csv", L".log", L".xml", L".json", L".cs",
			L".html", L".md", L".xaml", L".py", L".java", L".c", L".cpp"
		};

		LPWSTR ext = PathFindExtensionW(path.c_str());
		if (ext == NULL) return false;

		std::wstring extStr(ext);
		// 转换为小写进行比较
		std::transform(extStr.begin(), extStr.end(), extStr.begin(), ::towlower);

		return targetExtensions.count(extStr) > 0;
	}

	// 检查具体的 Window/View 是否包含选中的合法文件
	static bool HasValidSelection(IDispatch* pDispWindow)
	{
		if (!pDispWindow) return false;

		CComPtr<IWebBrowser2> pBrowser;
		HRESULT hr = pDispWindow->QueryInterface(IID_IWebBrowser2, (void**)&pBrowser);
		if (FAILED(hr)) return false;

		// 获取 Document
		CComPtr<IDispatch> pDispDoc;
		hr = pBrowser->get_Document(&pDispDoc);
		if (FAILED(hr) || !pDispDoc) return false;

		// 获取 Folder View
		CComPtr<IShellFolderViewDual> pFolderView;
		hr = pDispDoc->QueryInterface(IID_IShellFolderViewDual, (void**)&pFolderView);
		if (FAILED(hr)) return false;

		// 获取 SelectedItems
		CComPtr<FolderItems> pSelectedItems;
		hr = pFolderView->SelectedItems(&pSelectedItems);
		if (FAILED(hr) || !pSelectedItems) return false;

		long count = 0;
		pSelectedItems->get_Count(&count);
		if (count == 0) return false;

		// 遍历选中项
		for (long i = 0; i < count; i++)
		{
			CComVariant varIndex(i);
			CComPtr<FolderItem> pItem;
			hr = pSelectedItems->Item(varIndex, &pItem);

			if (SUCCEEDED(hr) && pItem)
			{
				VARIANT_BOOL isFolder = VARIANT_FALSE;
				pItem->get_IsFolder(&isFolder);

				// 如果是文件夹，跳过 (对应 C# Directory.Exists 逻辑)
				if (isFolder == VARIANT_TRUE) continue;

				BSTR bstrPath = NULL;
				pItem->get_Path(&bstrPath);
				if (bstrPath)
				{
					std::wstring path(bstrPath, SysStringLen(bstrPath));
					SysFreeString(bstrPath);

					if (IsTargetExtension(path))
					{
						return true;
					}
				}
			}
		}
		return false;
	}

	static HWND FindShellParent(HWND hWnd, bool& isDesktop)
	{
		isDesktop = false;
		HWND current = hWnd;

		while (current != NULL)
		{
			wchar_t className[256];
			GetClassNameW(current, className, 256);

			// 检查是否是普通资源管理器
			if (wcscmp(className, L"CabinetWClass") == 0)
			{
				return current;
			}

			// 检查是否是桌面
			if (wcscmp(className, L"Progman") == 0 || wcscmp(className, L"WorkerW") == 0)
			{
				isDesktop = true;
				return current;
			}

			current = GetParent(current);
		}

		// 兜底检查：如果上面的循环没找到，但 WorkerW 可能覆盖在 Progman 上
		HWND progman = FindWindowW(L"Progman", NULL);
		if (progman != NULL)
		{
			// 这里简化处理，如果找不到父级但也没报错，返回 NULL
		}

		return NULL;
	}

public:
	static bool IsDraggingSupportedFile()
	{
		// COM 初始化 (实际使用中建议在线程入口处初始化一次，不要在函数内频繁调用)
		//CoInitialize(NULL);
		bool result = false;

		try
		{
			// 1. 获取鼠标位置
			POINT mousePos;
			GetCursorPos(&mousePos);
			std::cout << "Mouse pos at (" << mousePos.x << ", " << mousePos.y << ")\n";
			// 2. 获取鼠标下的窗口句柄
			HWND targetHwnd = WindowFromPoint(mousePos);
			if (targetHwnd == NULL)
			{
				std::cout << "No window at mouse position\n" << std::endl;
				throw 0;
			}

			// 3. 向上回溯
			bool isDesktop = false;
			HWND shellHwnd = FindShellParent(targetHwnd, isDesktop);
			if (shellHwnd == NULL)
			{
				std::cout << "No shell parent found\n" << std::endl;
				throw 0;
			}

			// 4. 初始化 ShellWindows
			CComPtr<IShellWindows> pShellWindows;
			HRESULT hr = pShellWindows.CoCreateInstance(CLSID_ShellWindows);
			if (FAILED(hr))
			{
                std::cout << "Failed to create IShellWindows instance:" << hr << std::endl;
				throw 0;
			}
			// 5. 根据类型查找
			if (isDesktop)
			{
				// 对应 C#: shellWindows.FindWindowSW(..., SWC_DESKTOP, ..., SWFO_NEEDDISPATCH)
				// SWC_DESKTOP = 8, SWFO_NEEDDISPATCH = 1
				CComVariant vMissing;
				vMissing.vt = VT_ERROR;
				vMissing.scode = DISP_E_PARAMNOTFOUND;
				int hwndVal = 0;
				CComPtr<IDispatch> pDispDesktop;

				hr = pShellWindows->FindWindowSW(
					&vMissing, &vMissing,
					SWC_DESKTOP,
					(long*)&hwndVal,
					SWFO_NEEDDISPATCH,
					&pDispDesktop
				);

				if (SUCCEEDED(hr) && pDispDesktop)
				{
					result = HasValidSelection(pDispDesktop);
				}
			}
			else
			{
				// 遍历所有打开的 Explorer 窗口
				long winCount = 0;
				pShellWindows->get_Count(&winCount);

				for (long i = 0; i < winCount; i++)
				{
					CComVariant index(i);
					CComPtr<IDispatch> pDisp;
					hr = pShellWindows->Item(index, &pDisp);

					if (SUCCEEDED(hr) && pDisp)
					{
						CComPtr<IWebBrowser2> pBrowser;
						hr = pDisp->QueryInterface(IID_IWebBrowser2, (void**)&pBrowser);
						if (SUCCEEDED(hr))
						{
							SHANDLE_PTR hWindow = 0;
							pBrowser->get_HWND(&hWindow);

							if ((HWND)hWindow == shellHwnd)
							{
								result = HasValidSelection(pDisp);
								if (result) break;
							}
						}
					}
				}
			}
		}
		catch (...)
		{
			result = false;
		}

		//CoUninitialize();
		return result;
	}
};


//----------------------------------------------------------------------------------------


LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
	if (nCode == HC_ACTION) {
		MSLLHOOKSTRUCT* pMouse = (MSLLHOOKSTRUCT*)lParam;

		if (wParam == WM_LBUTTONDOWN) {
			/*std::cout << "Mouse Button Down at (" << pMouse->pt.x << ", " << pMouse->pt.y << ")\n";*/
			// 你可以在这里检测是否是拖拽开始的条件
			// 记录拖动起始位置
			g_dragStartPos = pMouse->pt;
			g_isDragging = false;
		}

		if (wParam == WM_MOUSEMOVE) {
			// 检测鼠标是否有拖拽操作
			if (GetAsyncKeyState(VK_LBUTTON) & 0x8000) {
				// 检查是否超过拖动阈值
				int deltaX = abs(pMouse->pt.x - g_dragStartPos.x);
				int deltaY = abs(pMouse->pt.y - g_dragStartPos.y);

				if (!g_isDragging && (deltaX > DRAG_THRESHOLD || deltaY > DRAG_THRESHOLD)) {
					g_isDragging = true;
					std::cout << "move begin..." << std::endl;

					// 尝试从拖放剪贴板获取文件信息
					//ExtractFileInfoFromDropClipboard();
					std::cout << "Mouse Button move at (" << pMouse->pt.x << ", " << pMouse->pt.y << ")\n";
					bool isDraggingSupportedFile = FileDetector1::IsDraggingSupportedFile();
					std::cout << "isDraggingSupportedFile: " << isDraggingSupportedFile << std::endl;
					
				}
			}
		}

		if (wParam == WM_LBUTTONUP)
		{
			if (g_isDragging) {
				std::cout << "move end" << std::endl;
				g_isDragging = false;
			}
		}
	}
	return CallNextHookEx(g_mouseHook, nCode, wParam, lParam);
}

// 从拖放剪贴板提取文件信息
void ExtractFileInfoFromDropClipboard() {
	//HRESULT hr = OleInitialize(nullptr);
	HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
	if (FAILED(hr)) {
		std::cerr << "OleInitialize failed: " << hr << std::endl;
		return;
	}
	IDataObject* pDataObject = nullptr;
	// 获取拖放剪贴板数据
	hr = OleGetClipboard(&pDataObject);
	if (SUCCEEDED(hr) && pDataObject) {
		FORMATETC fmtetc = {
			CF_HDROP,
			NULL,
			DVASPECT_CONTENT,
			-1,
			TYMED_HGLOBAL
		};
		FORMATETC fmtetcUnicodeText = { CF_UNICODETEXT, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		FORMATETC fmtetcText = { CF_TEXT, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };


		if (SUCCEEDED(pDataObject->QueryGetData(&fmtetc))) {
			STGMEDIUM stgmed;
			// 查询HDROP数据
			hr = pDataObject->GetData(&fmtetc, &stgmed);
			if (SUCCEEDED(hr)) {
				HDROP hDrop = static_cast<HDROP>(GlobalLock(stgmed.hGlobal));
				if (hDrop) {
					// 获取文件数量
					UINT fileCount = DragQueryFile(hDrop, 0xFFFFFFFF, nullptr, 0);
					std::cout << "drag file count: " << fileCount << std::endl;

					// 遍历所有文件
					for (UINT i = 0; i < fileCount; i++) {
						// 获取文件路径长度
						UINT pathLength = DragQueryFile(hDrop, i, nullptr, 0);
						if (pathLength > 0) {
							std::vector<TCHAR> buffer(pathLength + 1);
							DragQueryFile(hDrop, i, buffer.data(), pathLength + 1);

							std::wstring filePath(buffer.data());
							std::wcout << L"file " << (i + 1) << L": " << filePath << std::endl;

							// 获取文件属性信息
							ExtractFileTypeInfo(filePath);
						}
					}

					GlobalUnlock(stgmed.hGlobal);
				}
				ReleaseStgMedium(&stgmed);
			}
		}
		else
		{
			std::wcout << L"CF_HDROP format not supported" << std::endl;
		}

		if (SUCCEEDED(pDataObject->QueryGetData(&fmtetcUnicodeText)))
		{
			STGMEDIUM stgmedText;

			hr = pDataObject->GetData(&fmtetcUnicodeText, &stgmedText);
			if (SUCCEEDED(hr)) {
				wchar_t* pText = static_cast<wchar_t*>(GlobalLock(stgmedText.hGlobal));
				if (pText) {
					std::wstring text = pText;
					std::wcout << L"CF_UNICODETEXT: " << text << std::endl;
					// 检查文本是否为文件路径（简单的检查：是否包含扩展名，并且路径存在）
					// 这里我们假设文本就是文件路径
					if (GetFileAttributes(text.c_str()) != INVALID_FILE_ATTRIBUTES) {
						std::wcout << L"Text is a valid file path: " << text << std::endl;
						ExtractFileTypeInfo(text);
					}
					else {
						// 如果文本不是文件路径，我们将其视为普通文本
						std::wcout << L"Text is not a file path: " << text << std::endl;
					}
					GlobalUnlock(stgmedText.hGlobal);
				}
				ReleaseStgMedium(&stgmedText);
			}
		}
		else
		{
			std::wcout << L"CF_UNICODETEXT format not supported" << std::endl;
		}

		if (SUCCEEDED(pDataObject->QueryGetData(&fmtetcText)))
		{
			// 尝试CF_TEXT
			STGMEDIUM stgmedTextA;

			hr = pDataObject->GetData(&fmtetcText, &stgmedTextA);
			if (SUCCEEDED(hr)) {
				char* pText = static_cast<char*>(GlobalLock(stgmedTextA.hGlobal));
				if (pText) {
					std::string text = pText;
					std::cout << "CF_TEXT: " << text << std::endl;
					// 转换为宽字符串
					std::wstring wtext(text.begin(), text.end());
					if (GetFileAttributes(wtext.c_str()) != INVALID_FILE_ATTRIBUTES) {
						std::wcout << L"Text is a valid file path: " << wtext << std::endl;
						ExtractFileTypeInfo(wtext);
					}
					else {
						std::wcout << L"Text is not a file path: " << wtext << std::endl;
					}
					GlobalUnlock(stgmedTextA.hGlobal);
				}
				ReleaseStgMedium(&stgmedTextA);
			}
		}
		else
		{
			std::wcout << L"TEXT format not supported" << std::endl;
			// 尝试其他数据格式
			CheckOtherDataFormats(pDataObject);
		}

		pDataObject->Release();
	}
	else {
		std::wcerr << L"can not get clipboard data" << std::endl;
	}
	//OleUninitialize();
	CoUninitialize();
}

// 检查其他数据格式
void CheckOtherDataFormats(IDataObject* pDataObject) {
	// 查询支持的数据格式
	IEnumFORMATETC* pEnumFormat = nullptr;
	HRESULT hr = pDataObject->EnumFormatEtc(DATADIR_GET, &pEnumFormat);

	if (SUCCEEDED(hr) && pEnumFormat) {
		FORMATETC fmtetc;
		ULONG fetched = 0;

		std::cout << "data format in clipboard:" << std::endl;

		while (pEnumFormat->Next(1, &fmtetc, &fetched) == S_OK && fetched == 1) {
			TCHAR formatName[256];
			if (GetClipboardFormatName(fmtetc.cfFormat, formatName, 256) > 0) {
				std::wcout << L"format1: " << formatName << L" (ID: " << fmtetc.cfFormat << L")" << std::endl;
			}
			else {
				// 标准格式
				std::string stdFormatName;
				switch (fmtetc.cfFormat) {
				case CF_TEXT: stdFormatName = "CF_TEXT"; break;
				case CF_UNICODETEXT: stdFormatName = "CF_UNICODETEXT"; break;
				case CF_BITMAP: stdFormatName = "CF_BITMAP"; break;
				case CF_DIB: stdFormatName = "CF_DIB"; break;
				case CF_HDROP: stdFormatName = "CF_HDROP"; break;
				case CF_LOCALE: stdFormatName = "CF_LOCALE"; break;
				case CF_OEMTEXT: stdFormatName = "CF_OEMTEXT"; break;
				default: stdFormatName = "unknown"; break;
				}
				std::cout << "format2: " << stdFormatName << " (ID: " << fmtetc.cfFormat << ")" << std::endl;
			}
		}

		pEnumFormat->Release();
	}
}

// 提取文件类型信息
void ExtractFileTypeInfo(const std::wstring& filePath) {
	// 获取文件属性
	DWORD attr = GetFileAttributes(filePath.c_str());
	if (attr != INVALID_FILE_ATTRIBUTES) {
		if (attr & FILE_ATTRIBUTE_DIRECTORY) {
			std::wcout << L"  type: directory" << std::endl;
		}
		else {
			std::wcout << L"  type: file" << std::endl;

			// 获取文件扩展名
			size_t dotPos = filePath.find_last_of(L'.');
			if (dotPos != std::wstring::npos) {
				std::wstring extension = filePath.substr(dotPos);
				std::wcout << L"  extension: " << extension << std::endl;
			}

			// 获取文件大小
			HANDLE hFile = CreateFile(filePath.c_str(), GENERIC_READ, FILE_SHARE_READ,
				nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
			if (hFile != INVALID_HANDLE_VALUE) {
				LARGE_INTEGER fileSize;
				if (GetFileSizeEx(hFile, &fileSize)) {
					std::wcout << L"  size: " << fileSize.QuadPart << L" bytes" << std::endl;
				}
				CloseHandle(hFile);
			}
		}

		// 显示文件属性
		std::wcout << L"  properties: ";
		if (attr & FILE_ATTRIBUTE_READONLY) std::wcout << L"[read-only]";
		if (attr & FILE_ATTRIBUTE_HIDDEN) std::wcout << L"[hidden]";
		if (attr & FILE_ATTRIBUTE_SYSTEM) std::wcout << L"[system]";
		if (attr & FILE_ATTRIBUTE_ARCHIVE) std::wcout << L"[archive]";
		std::wcout << std::endl;
	}
}

void InstallMouseHook() {
	g_mouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseProc, NULL, 0);
	if (g_mouseHook == NULL) {
		std::cerr << "Failed to install mouse hook!" << std::endl;
	}
}

int main1() {
	HRESULT hr = CoInitialize(NULL);

	if (hr == RPC_E_CHANGED_MODE)
	{
		// 如果进入这里，说明你的程序在其他地方已经初始化为 MTA 了
		std::cerr << "错误：线程已经是 MTA 模式，Shell 接口需要 STA 模式。" << std::endl;
		// 如果必须在 MTA 线程运行，请参考方案三
	}
	InstallMouseHook();

	// 进入消息循环
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	// 卸载钩子
	UnhookWindowsHookEx(g_mouseHook);

	// 2. 结束清理
	CoUninitialize();
	return 0;
}
