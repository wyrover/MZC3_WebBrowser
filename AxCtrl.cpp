////////////////////////////////////////////////////////////////////////////
// AxCtrl.cpp -- MZC3 ActiveX control
// This file is part of MZC3.  See file "ReadMe.txt" and "License.txt".
////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"

#ifdef MZC_NO_INLINES
    #undef MZC_INLINE
    #define MZC_INLINE
    #include "AxCtrl_inl.hpp"
    #undef MZC_INLINE
    #define MZC_INLINE inline
#endif

///////////////////////////////////////////////////////////////////////////////
// MAxCtrl

/*virtual*/ MAxCtrl::~MAxCtrl()
{
}

void MAxCtrl::DoVerb(LONG iVerb)
{
    if (m_pUnknown == NULL)
        return;

    IOleObject* pioo;
    HRESULT hr = m_pUnknown->QueryInterface(
        IID_IOleObject, reinterpret_cast<void**>(&pioo));
    if (FAILED(hr))
        return;

    pioo->DoVerb(iVerb, NULL, this, 0, m_hwndParent, &m_rcCtrl);
    pioo->Release();
}

BOOL MAxCtrl::Create(CLSID clsid)
{
    ::CoCreateInstance(clsid, NULL,
        CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
        IID_IUnknown, reinterpret_cast<void**>(&m_pUnknown));

    if (m_pUnknown == NULL)
        return FALSE;

    IOleObject* pioo;
    HRESULT hr = m_pUnknown->QueryInterface(
        IID_IOleObject, reinterpret_cast<void**>(&pioo));
    if (FAILED(hr))
        return FALSE;

    pioo->SetClientSite(this);
    pioo->Release();

    IPersistStreamInit* ppsi;
    hr = m_pUnknown->QueryInterface(
        IID_IPersistStreamInit, reinterpret_cast<void**>(&ppsi));
    if (SUCCEEDED(hr))
    {
        ppsi->InitNew();
        ppsi->Release();
        return TRUE;
    }
    return FALSE;
}

void MAxCtrl::Destroy()
{
    if (m_pUnknown == NULL)
        return;

    IOleObject* pioo;
    HRESULT hr = m_pUnknown->QueryInterface(
        IID_IOleObject, reinterpret_cast<void**>(&pioo));
    if (SUCCEEDED(hr))
    {
        pioo->Close(OLECLOSE_NOSAVE);
        pioo->SetClientSite(NULL);
        pioo->Release();
    }

    IOleInPlaceObject* pipo;
    hr = m_pUnknown->QueryInterface(
        IID_IOleInPlaceObject, reinterpret_cast<void**>(&pipo));
    if (SUCCEEDED(hr))
    {
        pipo->UIDeactivate();
        pipo->InPlaceDeactivate();
        pipo->Release();
    }

    m_pUnknown->Release();
    m_pUnknown = NULL;
}

void MAxCtrl::SetPos(const MRect& rc)
{
    m_rcCtrl = rc;

    if (m_pUnknown == NULL)
        return;

    IOleInPlaceObject* pipo;
    HRESULT hr = m_pUnknown->QueryInterface(
        IID_IOleInPlaceObject, reinterpret_cast<void**>(&pipo));
    if (FAILED(hr))
        return;

    pipo->SetObjectRects(&m_rcCtrl, &m_rcCtrl);
    pipo->Release();
}

BOOL MAxCtrl::TranslateAccelerator(LPMSG pMsg)
{
    if (m_pUnknown == NULL)
        return FALSE;

    IOleInPlaceActiveObject* pao;
    HRESULT hr = m_pUnknown->QueryInterface(
        IID_IOleInPlaceActiveObject, reinterpret_cast<void**>(&pao));
    if (FAILED(hr))
        return FALSE;

    HRESULT result = pao->TranslateAccelerator(pMsg);
    pao->Release();
    return FAILED(result);
}

STDMETHODIMP MAxCtrl::QueryInterface(REFIID riid, void** ppvObject)
{
    if (ppvObject == NULL)
        return E_POINTER;

    if (IsEqualIID(riid, IID_IOleClientSite))
        *ppvObject = dynamic_cast<IOleClientSite*>(this);
    else if (IsEqualIID(riid, IID_IOleInPlaceSite))
        *ppvObject = dynamic_cast<IOleInPlaceSite*>(this);
    else if (IsEqualIID(riid, IID_IOleInPlaceFrame))
        *ppvObject = dynamic_cast<IOleInPlaceFrame*>(this);
    else if (IsEqualIID(riid, IID_IOleInPlaceUIWindow))
        *ppvObject = dynamic_cast<IOleInPlaceUIWindow*>(this);
    else if (IsEqualIID(riid, IID_IOleControlSite))
        *ppvObject = dynamic_cast<IOleControlSite*>(this);
    else if (IsEqualIID(riid, IID_IOleWindow))
        *ppvObject = this;
    else if (IsEqualIID(riid, IID_IDispatch))
        *ppvObject = dynamic_cast<IDispatch*>(this);
    else if (IsEqualIID(riid, IID_IUnknown))
        *ppvObject = this;
    else
    {
        *ppvObject = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

STDMETHODIMP_(ULONG) MAxCtrl::AddRef()
{
    return ++m_nRefs;
}

STDMETHODIMP_(ULONG) MAxCtrl::Release()
{
    if (--m_nRefs == 0)
    {
        delete this;
        return 0;
    }
    return m_nRefs;
}

STDMETHODIMP MAxCtrl::SaveObject()
{
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::GetMoniker(
    DWORD dwAssign, DWORD dwWhichMoniker, LPMONIKER* ppMk)
{
    UNREFERENCED_PARAMETER(dwAssign);
    UNREFERENCED_PARAMETER(dwWhichMoniker);
    UNREFERENCED_PARAMETER(ppMk);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::GetContainer(LPOLECONTAINER* ppContainer)
{
    UNREFERENCED_PARAMETER(ppContainer);
    return E_NOINTERFACE;
}

STDMETHODIMP MAxCtrl::ShowObject()
{
    return S_OK;
}

STDMETHODIMP MAxCtrl::OnShowWindow(BOOL fShow)
{
    UNREFERENCED_PARAMETER(fShow);
    return S_OK;
}

STDMETHODIMP MAxCtrl::RequestNewObjectLayout()
{
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::GetWindow(HWND* phwnd)
{
    if (!::IsWindow(m_hwndParent))
        return S_FALSE;

    *phwnd = m_hwndParent;
    return S_OK;
}

STDMETHODIMP MAxCtrl::ContextSensitiveHelp(BOOL fEnterMode)
{
    UNREFERENCED_PARAMETER(fEnterMode);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::CanInPlaceActivate()
{
    return S_OK;
}

STDMETHODIMP MAxCtrl::OnInPlaceActivate()
{
    return S_OK;
}

STDMETHODIMP MAxCtrl::OnUIActivate()
{
    return S_OK;
}

STDMETHODIMP MAxCtrl::GetWindowContext(
    IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppIIPUIWin,
    LPRECT prcPosRect, LPRECT prcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
    *ppFrame = dynamic_cast<IOleInPlaceFrame*>(this);
    *ppIIPUIWin = NULL;

    RECT rc;
    ::GetClientRect(m_hwndParent, &rc);
    prcPosRect->left = 0;
    prcPosRect->top = 0;
    prcPosRect->right  = rc.right;
    prcPosRect->bottom = rc.bottom;
    *prcClipRect = *prcPosRect;

    lpFrameInfo->cb             = sizeof(OLEINPLACEFRAMEINFO);
    lpFrameInfo->fMDIApp        = FALSE;
    lpFrameInfo->hwndFrame      = m_hwndParent;
    lpFrameInfo->haccel         = NULL;
    lpFrameInfo->cAccelEntries  = 0;

    (*ppFrame)->AddRef();
    return S_OK;
}

STDMETHODIMP MAxCtrl::Scroll(SIZE scrollExtent)
{
    UNREFERENCED_PARAMETER(scrollExtent);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::OnUIDeactivate(BOOL fUndoable)
{
    UNREFERENCED_PARAMETER(fUndoable);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::OnInPlaceDeactivate()
{
    return S_OK;
}

STDMETHODIMP MAxCtrl::DiscardUndoState()
{
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::DeactivateAndUndo()
{
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::OnPosRectChange(LPCRECT prcPosRect)
{
    UNREFERENCED_PARAMETER(prcPosRect);
    return S_OK;
}

STDMETHODIMP MAxCtrl::GetBorder(LPRECT prcBorder)
{
    UNREFERENCED_PARAMETER(prcBorder);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::RequestBorderSpace(LPCBORDERWIDTHS lpborderwidths)
{
    UNREFERENCED_PARAMETER(lpborderwidths);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::SetBorderSpace(LPCBORDERWIDTHS lpborderwidths)
{
    UNREFERENCED_PARAMETER(lpborderwidths);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::SetActiveObject(
    IOleInPlaceActiveObject* pActiveObject, LPCOLESTR lpszObjName)
{
    UNREFERENCED_PARAMETER(pActiveObject);
    UNREFERENCED_PARAMETER(lpszObjName);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::InsertMenus(
    HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths)
{
    UNREFERENCED_PARAMETER(hmenuShared);
    UNREFERENCED_PARAMETER(lpMenuWidths);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::SetMenu(
    HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject)
{
    UNREFERENCED_PARAMETER(hmenuShared);
    UNREFERENCED_PARAMETER(holemenu);
    UNREFERENCED_PARAMETER(hwndActiveObject);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::RemoveMenus(HMENU hmenuShared)
{
    UNREFERENCED_PARAMETER(hmenuShared);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::SetStatusText(LPCOLESTR pszStatusText)
{
    if (m_hwndStatus && ::IsWindow(m_hwndStatus))
    {
        char szClsName[64];
        ::GetClassNameA(m_hwndStatus, szClsName, 64);

        MWideToText strText(pszStatusText);
        if (::lstrcmpiA(szClsName, STATUSCLASSNAMEA) == 0)
        {
            ::SendMessage(m_hwndStatus, SB_SETTEXT, 0,
                reinterpret_cast<LPARAM>(strText.c_str()));
        }
        else
        {
            ::SetWindowText(m_hwndStatus, strText.c_str());
        }
    }
    return S_OK;
}

STDMETHODIMP MAxCtrl::EnableModeless(BOOL fEnable)
{
    UNREFERENCED_PARAMETER(fEnable);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::TranslateAccelerator(LPMSG pMsg, WORD wID)
{
    UNREFERENCED_PARAMETER(pMsg);
    UNREFERENCED_PARAMETER(wID);
    return S_OK;
}

STDMETHODIMP MAxCtrl::OnControlInfoChanged()
{
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::LockInPlaceActive(BOOL fLock)
{
    UNREFERENCED_PARAMETER(fLock);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::GetExtendedControl(IDispatch** ppDisp)
{
    if (ppDisp == NULL)
        return E_INVALIDARG;

    *ppDisp = dynamic_cast<IDispatch*>(this);
    (*ppDisp)->AddRef();

    return S_OK;
}

STDMETHODIMP MAxCtrl::TransformCoords(POINTL* pptlHimetric, POINTF* pptfContainer, DWORD dwFlags)
{
    UNREFERENCED_PARAMETER(pptlHimetric);
    UNREFERENCED_PARAMETER(pptfContainer);
    UNREFERENCED_PARAMETER(dwFlags);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::TranslateAccelerator(LPMSG pMsg, DWORD grfModifiers)
{
    UNREFERENCED_PARAMETER(pMsg);
    UNREFERENCED_PARAMETER(grfModifiers);
    return S_FALSE;
}

STDMETHODIMP MAxCtrl::OnFocus(BOOL fGotFocus)
{
    UNREFERENCED_PARAMETER(fGotFocus);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::ShowPropertyFrame()
{
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::GetIDsOfNames(
    REFIID riid, OLECHAR** rgszNames, UINT cNames,
    LCID lcid, DISPID* rgdispid)
{
    UNREFERENCED_PARAMETER(static_cast<IID>(riid)); // IID cast required for MinGW
    UNREFERENCED_PARAMETER(rgszNames);
    UNREFERENCED_PARAMETER(cNames);
    UNREFERENCED_PARAMETER(lcid);

    *rgdispid = DISPID_UNKNOWN;
    return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP MAxCtrl::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)
{
    UNREFERENCED_PARAMETER(itinfo);
    UNREFERENCED_PARAMETER(lcid);
    UNREFERENCED_PARAMETER(pptinfo);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::GetTypeInfoCount(UINT* pctinfo)
{
    UNREFERENCED_PARAMETER(pctinfo);
    return E_NOTIMPL;
}

STDMETHODIMP MAxCtrl::Invoke(
    DISPID dispid, REFIID riid, LCID lcid, WORD wFlags,
    DISPPARAMS* pdispparams, VARIANT* pvarResult,
    EXCEPINFO* pexecinfo, UINT* puArgErr)
{
    UNREFERENCED_PARAMETER(dispid);
    UNREFERENCED_PARAMETER(static_cast<IID>(riid)); // IID cast required for MinGW
    UNREFERENCED_PARAMETER(lcid);
    UNREFERENCED_PARAMETER(wFlags);
    UNREFERENCED_PARAMETER(pdispparams);
    UNREFERENCED_PARAMETER(pvarResult);
    UNREFERENCED_PARAMETER(pexecinfo);
    UNREFERENCED_PARAMETER(puArgErr);
    return DISP_E_MEMBERNOTFOUND;
}

IDispatch* MAxCtrl::GetDispatch()
{
    if (m_pUnknown == NULL)
        return NULL;

    IDispatch* pDispatch;
    HRESULT hr = m_pUnknown->QueryInterface(
        IID_IDispatch, reinterpret_cast<void**>(&pDispatch));
    if (SUCCEEDED(hr))
        return pDispatch;
    return NULL;
}

///////////////////////////////////////////////////////////////////////////////
