////////////////////////////////////////////////////////////////////////////
// WebBrowser_inl.hpp -- a Web browser control for Win32
// This file is part of MZC3.  See file "ReadMe.txt" and "License.txt".
////////////////////////////////////////////////////////////////////////////

MZC_INLINE void MWebBrowser::Destroy()
{
    GetAxCtrl()->Destroy();
}

MZC_INLINE MAxCtrl* MWebBrowser::GetAxCtrl() const
{
    return m_pAxCtrl;
}

MZC_INLINE IWebBrowser2* MWebBrowser::GetWebBrowser2() const
{
    return m_pWebBrowser2;
}

MZC_INLINE void MWebBrowser::MoveWindow(int x, int y, int cx, int cy)
{
    GetAxCtrl()->SetPos(x, y, cx, cy);
}

MZC_INLINE void MWebBrowser::MoveWindow(const MRect& rc)
{
    GetAxCtrl()->SetPos(rc);
}

MZC_INLINE HWND MWebBrowser::GetParent() const
{
    return GetAxCtrl()->GetParent();
}

MZC_INLINE void MWebBrowser::SetParent(HWND hwndParent)
{
    GetAxCtrl()->SetParent(hwndParent);
}

MZC_INLINE BOOL MWebBrowser::TranslateAccelerator(LPMSG pmsg)
{
    return GetAxCtrl()->TranslateAccelerator(pmsg);
}

////////////////////////////////////////////////////////////////////////////
