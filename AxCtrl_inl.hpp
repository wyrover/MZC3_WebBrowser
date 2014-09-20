////////////////////////////////////////////////////////////////////////////
// AxCtrl_inl.hpp -- MZC3 ActiveX control
// This file is part of MZC3.  See file "ReadMe.txt" and "License.txt".
////////////////////////////////////////////////////////////////////////////

MZC_INLINE MAxCtrl::MAxCtrl() :
    m_nRefs(1), m_hwndParent(NULL), m_pUnknown(NULL)
{
}

MZC_INLINE BOOL MAxCtrl::Create(BSTR bstrClsid)
{
    CLSID clsid;
    ::CLSIDFromString(bstrClsid, &clsid);
    return Create(clsid);
}

MZC_INLINE HWND MAxCtrl::GetParent() const
{
    return m_hwndParent;
}

MZC_INLINE void MAxCtrl::SetParent(HWND hwndParent)
{
    m_hwndParent = hwndParent;
}

MZC_INLINE IUnknown* MAxCtrl::GetUnknown()
{
    if (m_pUnknown == NULL)
        return NULL;
    m_pUnknown->AddRef();
    return m_pUnknown;
}

MZC_INLINE HWND MAxCtrl::GetStatusWindow() const
{
    return m_hwndStatus;
}

MZC_INLINE void MAxCtrl::SetStatusWindow(HWND hwndStatus)
{
    m_hwndStatus = hwndStatus;
}

MZC_INLINE void MAxCtrl::InPlaceActivate()
{
    DoVerb(OLEIVERB_INPLACEACTIVATE);
}

MZC_INLINE void MAxCtrl::UIActivate()
{
    DoVerb(OLEIVERB_UIACTIVATE);
}

MZC_INLINE void MAxCtrl::Show(BOOL fVisible/* = TRUE*/)
{
    DoVerb(fVisible ? OLEIVERB_SHOW : OLEIVERB_HIDE);
}

MZC_INLINE void MAxCtrl::SetPos(int x, int y, int width, int height)
{
    SetPos(MRect(x, y, x + width, y + height));
}

////////////////////////////////////////////////////////////////////////////
