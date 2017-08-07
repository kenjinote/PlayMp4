#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#include <atlbase.h>
#include <atlhost.h>

CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()

TCHAR szClassName[] = TEXT("Window");

BOOL PlayVideo(HWND hBrowser, LPCTSTR lpszFilePath)
{
	BOOL bReturn = FALSE;
	CComPtr<IUnknown> punkIE;
	if (AtlAxGetControl(hBrowser, &punkIE) == S_OK)
	{
		CComPtr<IHTMLDocument2> pIWebBrowser2;
		punkIE->QueryInterface(IID_IHTMLDocument2, (VOID**)&pIWebBrowser2);
		if (pIWebBrowser2)
		{
			IHTMLElement *pElement;
			pIWebBrowser2->get_body(&pElement);
			if (pElement)
			{
				TCHAR szHtml[1024];
				{
					TCHAR szUriPath[MAX_PATH] = { 0 };
					DWORD dwSize = _countof(szUriPath);
					UrlCreateFromPath(lpszFilePath, szUriPath, &dwSize, 0);
					wsprintf(szHtml, TEXT("<video style=\"width:100%%;height:100%%;left:50%%;position:absolute;top:50%%;transform:translate(-50%%,-50%%);\" src=\"%s\" autoplay loop></video>"), szUriPath);
				}
				BSTR bstrText = SysAllocString(szHtml);
				pElement->put_innerHTML(bstrText);
				SysFreeString(bstrText);
				pElement->Release();
				bReturn = TRUE;
			}
			pIWebBrowser2.Release();
		}
		punkIE.Release();
	}
	return bReturn;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND	hBrowser;
	static BOOL bFullScreen;
	switch (msg)
	{
	case WM_CREATE:
		AtlAxWinInit();
		_Module.Init(ObjectMap, ((LPCREATESTRUCT)lParam)->hInstance);
		hBrowser = CreateWindow(
			TEXT(ATLAXWIN_CLASS),
			TEXT("mshtml:<html><head><meta http-equiv=\"X-UA-Compatible\" content=\"IE=11\"></head><body style=\"margin:0; padding:0\"></body></html>"),
			WS_CHILD | WS_VISIBLE | WS_DISABLED, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_SIZE:
		MoveWindow(hBrowser, 0, 0, LOWORD(lParam), HIWORD(lParam), 1);
		break;
	case WM_DROPFILES:
		if (DragQueryFile((HDROP)wParam, 0xFFFFFFFF, NULL, 0) == 1)
		{
			TCHAR szFilePath[MAX_PATH];
			DragQueryFile((HDROP)wParam, 0, szFilePath, _countof(szFilePath));
			PlayVideo(hBrowser, szFilePath);
		}
		DragFinish((HDROP)wParam);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == 1000)
		{
			HWND hTray = FindWindow(TEXT("Shell_TrayWnd"), NULL);
			if (hTray)
			{
				if (bFullScreen)
				{
					ShowWindow(hTray, SW_SHOW);
					APPBARDATA ABData;
					ABData.cbSize = sizeof(APPBARDATA);
					ABData.hWnd = hTray;
					ABData.lParam = FALSE;
					SHAppBarMessage(ABM_SETSTATE, &ABData);
					SetWindowLong(hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) | WS_BORDER | WS_DLGFRAME | WS_THICKFRAME);
					SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOMOVE | SWP_FRAMECHANGED);
					ShowWindow(hWnd, SW_RESTORE);
					bFullScreen = FALSE;
				}
				else
				{
					ShowWindow(hTray, SW_SHOW);
					APPBARDATA ABData;
					ABData.cbSize = sizeof(APPBARDATA);
					ABData.hWnd = hTray;
					ABData.lParam = ABS_AUTOHIDE;
					SHAppBarMessage(ABM_SETSTATE, &ABData);
					SetWindowLong(hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) & ~WS_BORDER & ~WS_DLGFRAME & ~WS_THICKFRAME);
					SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOMOVE | SWP_FRAMECHANGED);
					ShowWindow(hWnd, SW_MAXIMIZE);
					bFullScreen = TRUE;
				}
			}
		}
		break;
	case WM_DESTROY:
		DestroyWindow(hBrowser);
		AtlAxWinTerm();
		_Module.Term();
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		0,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		0,
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("ドラッグ&ドロップされたMP4ファイルを再生する"),
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	ACCEL Accel[] = { { FVIRTKEY,VK_F11,1000 } };
	HACCEL hAccel = CreateAcceleratorTable(Accel, 1);
	while (GetMessage(&msg, 0, 0, 0))
	{
		if (!TranslateAccelerator(hWnd, hAccel, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	DestroyAcceleratorTable(hAccel);
	return (int)msg.wParam;
}
