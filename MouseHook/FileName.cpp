#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <exdisp.h>  // For IShellWindows
#include <shldisp.h> // For IShellFolderViewDual
#include <comdef.h>   // For _bstr_t, _variant_t
#include <set>
#include <string>
#include <vector>
#include <iostream>

class FileDetector3 {
private:
	static const std::set<std::wstring> targetExtensions;
	static const int SWC_DESKTOP = 8;
	static const int SWFO_NEEDDISPATCH = 1;

	struct ShellWindowInfo {
		HWND hwnd;
		bool isDesktop;
	};

	static std::wstring GetWindowClassName(HWND hWnd) {
		wchar_t className[256];
		GetClassName(hWnd, className, sizeof(className) / sizeof(wchar_t));
		return std::wstring(className);
	}

	static ShellWindowInfo FindShellParent(HWND hWnd) {
		ShellWindowInfo result = { nullptr, false };
		HWND current = hWnd;

		while (current != nullptr) {
			std::wstring className = GetWindowClassName(current);

			// Check if it's a regular explorer window
			if (className == L"CabinetWClass") {
				result.hwnd = current;
				result.isDesktop = false;
				return result;
			}

			// Check if it's desktop
			if (className == L"Progman" || className == L"WorkerW") {
				result.hwnd = current;
				result.isDesktop = true;
				return result;
			}

			current = GetParent(current);
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

		// 直接获取 FolderItems
		FolderItems* selectedItems = nullptr;
		HRESULT hr = folderView->SelectedItems(&selectedItems);
		if (FAILED(hr) || !selectedItems) {
			folderView->Release();
			documentDispatch->Release();
			return false;
		}

		long itemCount = 0;
		if (FAILED(selectedItems->get_Count(&itemCount)) || itemCount == 0) {
			selectedItems->Release();
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
		folderView->Release();
		documentDispatch->Release();
		return foundValidFile;
	}

public:
	static bool IsDraggingSupportedFile() {
		// Initialize COM
		if (FAILED(CoInitialize(nullptr))) {
			return false;
		}

		bool result = false;
		try {
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

			if (shellInfo.isDesktop) {
				result = CheckDesktopSelection(shellWindows);
			}
			else {
				result = CheckExplorerWindow(shellWindows, shellInfo.hwnd);
			}

			shellWindows->Release();

		}
		catch (...) {
			result = false;
		}

		CoUninitialize();
		return result;
	}
};

// Initialize static member
const std::set<std::wstring> FileDetector3::targetExtensions = {
	L".txt", L".csv", L".log", L".xml", L".json", L".cs", L".html",
	L".md", L".xaml", L".py", L".java", L".c", L".cpp"
};

// Main 用于测试
int main3()
{
	std::cout << "Monitoring mouse... Drag a file (e.g., .txt) to see detection." << std::endl;
	std::cout << "Press Ctrl+C to exit." << std::endl;

	while (true)
	{
		if (FileDetector3::IsDraggingSupportedFile())
		{
			std::cout << "[DETECTED] Dragging supported file!" << std::endl;
		}
		Sleep(500);
	}
	return 0;
}