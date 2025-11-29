#include <windows.h>
#include <shlobj.h>
#include <shldisp.h>
#include <exdisp.h>
#include <shlwapi.h>
#include <atlbase.h> // 使用 CComPtr 简化 COM 内存管理
#include <iostream>
#include <vector>
#include <string>
#include <algorithm>
#include <set>

#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Shlwapi.lib")

namespace SystemDrag
{
	class FileDetector
	{
	private:
		static bool IsTargetExtension(const std::wstring& path)
		{
			// 定义目标后缀名
			static const std::set<std::wstring> targetExtensions = {
				L".txt", L".csv", L".log", L".xml", L".json", L".cs",
				L".html", L".md", L".xaml", L".py", L".java", L".c", L".cpp", L".png", L".doc", L".docx", L".pdf", L".jpg", L".jpeg", L".bmp"
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

				// 2. 获取鼠标下的窗口句柄
				HWND targetHwnd = WindowFromPoint(mousePos);
				if (targetHwnd == NULL) throw 0;

				// 3. 向上回溯
				bool isDesktop = false;
				HWND shellHwnd = FindShellParent(targetHwnd, isDesktop);
				if (shellHwnd == NULL) throw 0;

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
}


// 自定义消息 ID，用于主线程接收钩子发来的请求
#define WM_PERFORM_DRAG_CHECK (WM_USER + 100)
// 全局主窗口句柄（用于发送消息）
static DWORD g_mainThreadId = 0;


// 全局状态变量
static bool g_isLButtonDown = false;
static bool g_isDragging = false;
static bool g_detectionCalled = false;
static POINT g_dragStartPos = { 0, 0 };

// 系统拖拽阈值
static int g_minDragX = 0;
static int g_minDragY = 0;

// 钩子句柄
static HHOOK g_mouseHook = NULL;

// =========================================================
// 2. 钩子回调函数 (核心检测逻辑)
// =========================================================

LRESULT CALLBACK MouseHookProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	// 确保处理是有效的 (nCode >= 0)
	if (nCode >= 0)
	{
		MSLLHOOKSTRUCT* pMouseStruct = (MSLLHOOKSTRUCT*)lParam;
		POINT currentPos = pMouseStruct->pt;

		switch (wParam)
		{
		case WM_LBUTTONDOWN:
		{
			// 鼠标左键按下
			g_isLButtonDown = true;
			g_isDragging = false;
			g_detectionCalled = false;
			g_dragStartPos = currentPos;
			std::cout << "\n[EVENT] LButton Down.\n";
			break;
		}

		case WM_MOUSEMOVE:
		{
			if (g_isLButtonDown)
			{
				if (!g_isDragging)
				{
					// 正在按下，检查是否超过拖拽阈值
					long dx = std::abs(currentPos.x - g_dragStartPos.x);
					long dy = std::abs(currentPos.y - g_dragStartPos.y);

					if (dx >= g_minDragX || dy >= g_minDragY)
					{
						// 达到拖拽阈值
						g_isDragging = true;
						std::cout << "[EVENT] Dragging Started.\n";
					}
				}

				if (g_isDragging && !g_detectionCalled)
				{
					// 正在拖拽，执行文件检测
					PostThreadMessage(g_mainThreadId, WM_PERFORM_DRAG_CHECK, 0, 0);
					g_detectionCalled = true;
				}
			}
			break;
		}

		case WM_LBUTTONUP:
		{
			// 鼠标左键释放
			if (g_isDragging)
			{
				std::cout << "[EVENT] Dragging Released.\n";
			}
			// 重置状态
			g_isLButtonDown = false;
			g_isDragging = false;
			g_detectionCalled = false;
			break;
		}
		}
	}

	// 务必调用下一个钩子，将事件传递下去
	return CallNextHookEx(g_mouseHook, nCode, wParam, lParam);
}



// Main 用于测试
int main()
{
	std::cout << "Monitoring mouse... Drag a file (e.g., .txt) to see detection." << std::endl;
	std::cout << "Press Ctrl+C to exit." << std::endl;

	// --- 1. COM Initialization (STA 是必需的) ---
	HRESULT hr = CoInitialize(NULL);
	if (hr != S_OK && hr != S_FALSE)
	{
		std::cerr << "COM Initialization failed (HRESULT: " << std::hex << hr << "). Shell operations require STA." << std::endl;
		return 1;
	}

	// --- 2. 存储主线程 ID (供钩子使用) ---
	g_mainThreadId = GetCurrentThreadId();

	// --- 3. 获取拖拽阈值并安装钩子 (与之前相同) ---
	g_minDragX = GetSystemMetrics(SM_CXDRAG);
	g_minDragY = GetSystemMetrics(SM_CYDRAG);
	std::cout << "Drag Detector Active (Hook installed). Threshold: " << g_minDragX << "px" << std::endl;

	// --- 安装低级鼠标钩子 ---
	// WH_MOUSE_LL: 低级鼠标事件
	// MouseHookProc: 回调函数
	// NULL: HMODULE，对于低级钩子，通常设为 NULL
	// 0: dwThreadId，设为 0 表示全局钩子
	g_mouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseHookProc, NULL, 0);

	if (g_mouseHook == NULL)
	{
		std::cerr << "Failed to install hook! Error: " << GetLastError() << std::endl;
		CoUninitialize();
		return 1;
	}

	// --- 4. 运行消息循环 (直接处理自定义消息) ---
	// 钩子消息会发送到这个线程的消息队列，然后被 DispatchMessage 调用 MouseHookProc 处理。
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0))
	{
		// 检查是否是我们自定义的消息
		if (msg.message == WM_PERFORM_DRAG_CHECK)
		{
			// 确保 COM 操作在主线程（STA线程）中执行
			if (SystemDrag::FileDetector::IsDraggingSupportedFile())
			{
				std::cout << "[检测成功] 正在拖拽支持的文件!\n";
			}
		}
		else
		{
			// 处理其他标准系统消息
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	// --- 5. 清理 ---
	UnhookWindowsHookEx(g_mouseHook);
	CoUninitialize();
	return 0;
}