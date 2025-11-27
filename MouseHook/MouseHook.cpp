#include <windows.h>
#include <shlobj.h>
#include <vector>
#include <string>
#include <iostream>
#include <chrono>
#include <thread>

HHOOK g_mouseHook;
bool g_isDragging = false;
POINT g_dragStartPos = { 0, 0 };
const int DRAG_THRESHOLD = 3; // 拖动阈值（像素）

void ExtractFileInfoFromDropClipboard();
void CheckOtherDataFormats(IDataObject* pDataObject);
void ExtractFileTypeInfo(const std::wstring& filePath);


LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
	if (nCode == HC_ACTION) {
		MSLLHOOKSTRUCT* pMouse = (MSLLHOOKSTRUCT*)lParam;

		if (wParam == WM_LBUTTONDOWN) {
			//std::cout << "Mouse Button Down at (" << pMouse->pt.x << ", " << pMouse->pt.y << ")\n";
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
					ExtractFileInfoFromDropClipboard();
					
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
	InstallMouseHook();

	// 进入消息循环
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	// 卸载钩子
	UnhookWindowsHookEx(g_mouseHook);
	return 0;
}
