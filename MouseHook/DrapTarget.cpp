#include <windows.h>
#include <shlobj.h>
#include <vector>
#include <string>
#include <iostream>

class DropTarget : public IDropTarget {
private:
	ULONG m_refCount;

public:
	DropTarget() : m_refCount(1) {}

	// IUnknown methods
	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override {
		if (riid == IID_IDropTarget || riid == IID_IUnknown) {
			*ppvObject = static_cast<IDropTarget*>(this);
			AddRef();
			return S_OK;
		}
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	STDMETHODIMP_(ULONG) AddRef() override {
		return InterlockedIncrement(&m_refCount);
	}

	STDMETHODIMP_(ULONG) Release() override {
		ULONG refCount = InterlockedDecrement(&m_refCount);
		if (refCount == 0) {
			delete this;
		}
		return refCount;
	}

	// IDropTarget methods
	STDMETHODIMP DragEnter(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) override {
		std::cout << "DragEnter detected!" << std::endl;

		// 检查数据格式
		FORMATETC fmtetc = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		if (SUCCEEDED(pDataObj->QueryGetData(&fmtetc))) {
			std::cout << "File drag detected" << std::endl;
			*pdwEffect = DROPEFFECT_COPY;

			// 立即提取文件信息
			ExtractFileInfoFromDataObject(pDataObj);
		}
		else {
			std::cout << "No file data available" << std::endl;
			*pdwEffect = DROPEFFECT_NONE;
		}

		return S_OK;
	}

	STDMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) override {
		*pdwEffect = DROPEFFECT_COPY;
		return S_OK;
	}

	STDMETHODIMP DragLeave() override {
		std::cout << "DragLeave" << std::endl;
		return S_OK;
	}

	STDMETHODIMP Drop(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) override {
		std::cout << "Drop completed!" << std::endl;
		*pdwEffect = DROPEFFECT_COPY;
		return S_OK;
	}

private:
	void ExtractFileInfoFromDataObject(IDataObject* pDataObj) {
		FORMATETC fmtetc = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		STGMEDIUM stgmed;

		HRESULT hr = pDataObj->GetData(&fmtetc, &stgmed);
		if (SUCCEEDED(hr)) {
			HDROP hDrop = static_cast<HDROP>(GlobalLock(stgmed.hGlobal));
			if (hDrop) {
				UINT fileCount = DragQueryFile(hDrop, 0xFFFFFFFF, nullptr, 0);
				std::cout << "Files being dragged: " << fileCount << std::endl;

				for (UINT i = 0; i < fileCount; i++) {
					UINT pathLength = DragQueryFile(hDrop, i, nullptr, 0);
					if (pathLength > 0) {
						std::vector<TCHAR> buffer(pathLength + 1);
						DragQueryFile(hDrop, i, buffer.data(), pathLength + 1);
						std::wstring filePath(buffer.data());
						std::wcout << L"File " << (i + 1) << L": " << filePath << std::endl;
					}
				}
				GlobalUnlock(stgmed.hGlobal);
			}
			ReleaseStgMedium(&stgmed);
		}
	}
};

// 注册drop target的窗口
class DropTargetWindow {
private:
	HWND m_hwnd;
	DropTarget* m_dropTarget;
	DWORD m_dropTargetCookie;

public:
	DropTargetWindow() : m_hwnd(nullptr), m_dropTarget(nullptr), m_dropTargetCookie(0) {}

	bool Initialize() {
		WNDCLASSEX wc = {};
		wc.cbSize = sizeof(WNDCLASSEX);
		wc.lpfnWndProc = WndProc;
		wc.hInstance = GetModuleHandle(nullptr);
		wc.lpszClassName = L"GlobalDropTarget";
		wc.hCursor = LoadCursor(nullptr, IDC_ARROW);

		RegisterClassEx(&wc);

		// 创建不可见窗口
		m_hwnd = CreateWindowEx(
			WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
			L"GlobalDropTarget",
			L"Global Drop Monitor",
			WS_POPUP,
			0, 0, 1, 1,
			nullptr, nullptr, GetModuleHandle(nullptr), this);

		if (!m_hwnd) return false;

		// 注册drop target
		m_dropTarget = new DropTarget();
		HRESULT hr = RegisterDragDrop(m_hwnd, m_dropTarget);
		if (FAILED(hr)) {
			std::cerr << "RegisterDragDrop failed: " << hr << std::endl;
			return false;
		}

		// 使用CoLockObjectExternal保持对象活跃
		CoLockObjectExternal(m_dropTarget, TRUE, FALSE);

		return true;
	}

	static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
		if (msg == WM_CREATE) {
			SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)((CREATESTRUCT*)lParam)->lpCreateParams);
		}

		DropTargetWindow* self = (DropTargetWindow*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
		if (self) {
			return self->HandleMessage(hwnd, msg, wParam, lParam);
		}

		return DefWindowProc(hwnd, msg, wParam, lParam);
	}

	LRESULT HandleMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
		switch (msg) {
		case WM_DESTROY:
			RevokeDragDrop(m_hwnd);
			CoLockObjectExternal(m_dropTarget, FALSE, TRUE);
			m_dropTarget->Release();
			PostQuitMessage(0);
			return 0;
		}
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}

	HWND GetWindow() const { return m_hwnd; }
};

int main2() {
	OleInitialize(nullptr);

	DropTargetWindow dropWindow;
	if (dropWindow.Initialize()) {
		std::cout << "Global drop target registered. Start dragging files to see information!" << std::endl;

		MSG msg;
		while (GetMessage(&msg, nullptr, 0, 0)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	else {
		std::cerr << "Failed to initialize drop target!" << std::endl;
	}

	OleUninitialize();
	return 0;
}