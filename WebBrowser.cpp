////////////////////////////////////////////////////////////////////////////
// WebBrowser.cpp -- a Web browser control for Win32
// This file is part of MZC3.  See file "ReadMe.txt" and "License.txt".
////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"

#ifdef MZC_NO_INLINES
    #undef MZC_INLINE
    #define MZC_INLINE
    #include "WebBrowser_inl.hpp"
    #undef MZC_INLINE
    #define MZC_INLINE inline
#endif

///////////////////////////////////////////////////////////////////////////////
// MWebBrowser

MWebBrowser::MWebBrowser() :
    m_pWebBrowser2(NULL), m_pAxCtrl(new MAxCtrl)
{
    ::OleInitialize(NULL);
}

/*virtual*/ MWebBrowser::~MWebBrowser()
{
    if (m_pWebBrowser2)
    {
        m_pWebBrowser2->Stop();
        m_pWebBrowser2->Release();
    }
    if (m_pAxCtrl)
    {
        m_pAxCtrl->Release();
    }
    ::OleUninitialize();
}

BOOL MWebBrowser::Create(
    HWND hwndParent, LPCTSTR pszURL/* = TEXT("about:blank")*/)
{
    if (!GetAxCtrl()->Create(CLSID_WebBrowser))
        return FALSE;

    GetAxCtrl()->SetParent(hwndParent);
    GetAxCtrl()->Show(TRUE);

    IUnknown* pUnk = GetAxCtrl()->GetUnknown();
    if (pUnk)
    {
        HRESULT hr = pUnk->QueryInterface(
            IID_IWebBrowser2, reinterpret_cast<void**>(&m_pWebBrowser2));
        pUnk->Release();

        if (SUCCEEDED(hr))
        {
            if (pszURL)
            {
                Navigate(pszURL);
            }
            return TRUE;
        }
    }
    return FALSE;
}

/*virtual*/ void MWebBrowser::Navigate(LPCTSTR pszURL)
{
    VARIANT vURL;
    vURL.vt = VT_BSTR;
    vURL.bstrVal = ::SysAllocString(MTextToWide(pszURL));

    VARIANT ve1, ve2, ve3, ve4;
    ve1.vt = VT_EMPTY;
    ve2.vt = VT_EMPTY;
    ve3.vt = VT_EMPTY;
    ve4.vt = VT_EMPTY;

    GetWebBrowser2()->Navigate2(&vURL, &ve1, &ve2, &ve3, &ve4);

    // Also frees memory allocated by SysAllocString
    VariantClear(&vURL);
}

///////////////////////////////////////////////////////////////////////////////

#ifdef UNITTEST

    MWebBrowser g_wb;

    LRESULT CALLBACK WindowProc(
        HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
        switch (uMsg)
        {
        case WM_CREATE:
            if (!g_wb.Create(hwnd, TEXT("http://localhost/")))
                return -1;
            //if (!g_wb.Create(hwnd, TEXT("http://google.co.jp/")))
            //    return -1;
            break;

        case WM_SIZE:
            {
                MRect rc;
                ::GetClientRect(hwnd, &rc);
                g_wb.MoveWindow(0, 0, rc.Width(), rc.Height());
            }
            break;

        case WM_DESTROY:
            g_wb.Destroy();
            PostQuitMessage(0);
            break;

        default:
            return ::DefWindowProc(hwnd, uMsg, wParam, lParam);
        }
        return 0;
    }

    int WINAPI WinMain(
        HINSTANCE hInstance,
        HINSTANCE hPrevInstance,
        LPSTR lpCmdLine,
        INT nCmdShow)
    {
        const TCHAR *pszName = TEXT("KHMZ's WebBrowser Test");

        WNDCLASS wc;
        wc.style = 0; 
        wc.lpfnWndProc = WindowProc; 
        wc.cbClsExtra = 0; 
        wc.cbWndExtra = 0; 
        wc.hInstance = hInstance; 
        wc.hIcon = ::LoadIcon(NULL, IDI_APPLICATION);
        wc.hCursor = ::LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground =
            reinterpret_cast<HBRUSH>(static_cast<INT_PTR>(COLOR_3DFACE + 1));
        wc.lpszMenuName = NULL; 
        wc.lpszClassName = pszName;
        if (!::RegisterClass(&wc))
        {
            assert(0);
            return 1;
        }

        HWND hwndMain = ::CreateWindow(pszName, pszName,
            WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, 0, CW_USEDEFAULT, 0,
            NULL, NULL, hInstance, NULL);
        if (hwndMain == NULL)
        {
            assert(0);
            return 2;
        }

        ::ShowWindow(hwndMain, nCmdShow);
        ::UpdateWindow(hwndMain);

        MSG msg;
        while (::GetMessage(&msg, NULL, 0, 0))
        {
            if (!g_wb.TranslateAccelerator(&msg))
            {
                ::TranslateMessage(&msg);
                ::DispatchMessage(&msg);
            }
        }

        return static_cast<int>(msg.wParam);
    }
#endif  // def UNITTEST

///////////////////////////////////////////////////////////////////////////////
