////////////////////////////////////////////////////////////////////////////
// WebBrowser.hpp -- a Web browser control for Win32
// This file is part of MZC3.  See file "ReadMe.txt" and "License.txt".
////////////////////////////////////////////////////////////////////////////

#ifndef __MZC3_WEBBROWSER__
#define __MZC3_WEBBROWSER__

#ifndef __MZC3_AXCTRL__
    #include "AxCtrl.hpp"
#endif

///////////////////////////////////////////////////////////////////////////////
// MWebBrowser

class MWebBrowser EXTENDS_MOBJECT
{
public:
    MWebBrowser();
    virtual ~MWebBrowser();

    BOOL Create(HWND hwndParent, LPCTSTR pszURL = TEXT("about:blank"));
    void Destroy();

    virtual void Navigate(LPCTSTR pszURL);

    HWND GetParent() const;
    void SetParent(HWND hwndParent);
    void MoveWindow(int x, int y, int cx, int cy);
    void MoveWindow(const MRect& rc);

    MAxCtrl*        GetAxCtrl() const;
    IWebBrowser2*   GetWebBrowser2() const;
    BOOL TranslateAccelerator(LPMSG pmsg);

protected:
    IWebBrowser2*   m_pWebBrowser2;
    MAxCtrl*        m_pAxCtrl;
};

///////////////////////////////////////////////////////////////////////////////

#ifndef MZC_NO_INLINES
    #undef MZC_INLINE
    #define MZC_INLINE inline
    #include "WebBrowser_inl.hpp"
#endif

#endif  // ndef __MZC3_WEBBROWSER__
