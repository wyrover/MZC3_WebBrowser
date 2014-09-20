////////////////////////////////////////////////////////////////////////////
// AxCtrl.hpp -- MZC3 ActiveX control
// This file is part of MZC3.  See file "ReadMe.txt" and "License.txt".
////////////////////////////////////////////////////////////////////////////

#ifndef __MZC3_AXCTRL__
#define __MZC3_AXCTRL__

#include <exdisp.h>
#include <ocidl.h>

///////////////////////////////////////////////////////////////////////////////
// MAxCtrl --- ActiveX window

class MAxCtrl :
    public IOleClientSite, public IOleInPlaceSite, public IOleInPlaceFrame,
    public IOleControlSite, public IDispatch
{
public:
    MAxCtrl();
    virtual ~MAxCtrl();

    BOOL Create(BSTR bstrClsid);
    BOOL Create(CLSID clsid);

    void UIActivate();
    void InPlaceActivate();
    void Show(BOOL fVisible = TRUE);
    void Destroy();

    HWND GetParent() const;
    void SetParent(HWND hwndParent);

    HWND GetStatusWindow() const;
    void SetStatusWindow(HWND hwndStatus);

    BOOL TranslateAccelerator(LPMSG pMsg);
    void SetPos(const MRect& rc);
    void SetPos(int x, int y, int width, int height);

    IDispatch*   GetDispatch();
    IUnknown*    GetUnknown();

    void DoVerb(LONG iVerb);

    // IUnknown Methods
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IOleClientSite Methods
    STDMETHODIMP SaveObject();
    STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, LPMONIKER* ppMk);
    STDMETHODIMP GetContainer(LPOLECONTAINER* ppContainer);
    STDMETHODIMP ShowObject();
    STDMETHODIMP OnShowWindow(BOOL fShow);
    STDMETHODIMP RequestNewObjectLayout();

    // IOleWindow Methods
    STDMETHODIMP GetWindow(HWND* phwnd);
    STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);

    // IOleInPlaceSite Methods
    STDMETHODIMP CanInPlaceActivate();
    STDMETHODIMP OnInPlaceActivate();
    STDMETHODIMP OnUIActivate();
    STDMETHODIMP GetWindowContext(
        IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc,
        LPRECT prcPosRect, LPRECT prcClipRect,
        LPOLEINPLACEFRAMEINFO lpFrameInfo);
    STDMETHODIMP Scroll(SIZE scrollExtent);
    STDMETHODIMP OnUIDeactivate(BOOL fUndoable);
    STDMETHODIMP OnInPlaceDeactivate();
    STDMETHODIMP DiscardUndoState();
    STDMETHODIMP DeactivateAndUndo();
    STDMETHODIMP OnPosRectChange(LPCRECT prcPosRect);

    // IOleInPlaceUIWindow Methods
    STDMETHODIMP GetBorder(LPRECT prcBorder);
    STDMETHODIMP RequestBorderSpace(LPCBORDERWIDTHS lpborderwidths);
    STDMETHODIMP SetBorderSpace(LPCBORDERWIDTHS lpborderwidths);
    STDMETHODIMP SetActiveObject(
        IOleInPlaceActiveObject* pActiveObject, LPCOLESTR lpszObjName);

    // IOleInPlaceFrame Methods
    STDMETHODIMP InsertMenus(
        HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths);
    STDMETHODIMP SetMenu(
        HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject);
    STDMETHODIMP RemoveMenus(HMENU hmenuShared);
    STDMETHODIMP SetStatusText(LPCOLESTR pszStatusText);
    STDMETHODIMP EnableModeless(BOOL fEnable);
    STDMETHODIMP TranslateAccelerator(LPMSG lpmsg, WORD wID);

    // IOleControlSite Methods
    STDMETHODIMP OnControlInfoChanged();
    STDMETHODIMP LockInPlaceActive(BOOL fLock);
    STDMETHODIMP GetExtendedControl(IDispatch** ppDisp);
    STDMETHODIMP TransformCoords(
        POINTL* pptlHimetric, POINTF* pptfContainer, DWORD dwFlags);
    STDMETHODIMP TranslateAccelerator(LPMSG pMsg, DWORD grfModifiers);
    STDMETHODIMP OnFocus(BOOL fGotFocus);
    STDMETHODIMP ShowPropertyFrame();

    // IDispatch Methods
    STDMETHODIMP GetIDsOfNames(REFIID riid, OLECHAR** rgszNames,
        UINT cNames, LCID lcid, DISPID* rgdispid);
    STDMETHODIMP GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo);
    STDMETHODIMP GetTypeInfoCount(UINT* pctinfo);
    STDMETHODIMP Invoke(
        DISPID dispid, REFIID riid, LCID lcid, WORD wFlags,
        DISPPARAMS* pdispparams, VARIANT* pvarResult,
        EXCEPINFO* pexecinfo, UINT* puArgErr);

protected:
     ULONG       m_nRefs;
     HWND        m_hwndParent;
     HWND        m_hwndStatus;
     IUnknown*   m_pUnknown;
     MRect       m_rcCtrl;
};

///////////////////////////////////////////////////////////////////////////////

#ifndef MZC_NO_INLINES
    #undef MZC_INLINE
    #define MZC_INLINE inline
    #include "AxCtrl_inl.hpp"
#endif

///////////////////////////////////////////////////////////////////////////////

#endif  // ndef __MZC3_AXCTRL__
