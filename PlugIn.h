#pragma once

#include "HEInterface.h"
#include "..\shared\PluginUtils.h"

class ATL_NO_VTABLE CLiveColors :
    public CPluginFactoryExt::TPlugin<CLiveColors>,
    public IStyleInformerCreator,
    public IMainMenuHandler,
    public CFrameEventsImplExt
{
public:
    CLiveColors() : m_eDisplayMode(eDisplayModeUnderline3x) {}
    virtual ~CLiveColors() {}

    BEGIN_COM_MAP(CLiveColors)
        COM_INTERFACE_ENTRY(IPlugin)
        COM_INTERFACE_ENTRY2(IDispatch, IPlugin)
        COM_INTERFACE_ENTRY(IStyleInformerCreator)
        COM_INTERFACE_ENTRY(IFrameEvents)
        COM_INTERFACE_ENTRY(IMainMenuHandler)
    END_COM_MAP()

    DECLARE_PROTECT_FINAL_CONSTRUCT()

public:
    enum eDisplayMode
    {
        eDisplayModeBackground,
        eDisplayModeUnderline1x,
        eDisplayModeUnderline2x,
        eDisplayModeUnderline3x,
        eDisplayModeUnderline4x
    };

protected:
    typedef CComPtr<IStyle> Style;
    typedef CAdapt<Style> StylePtr;
    typedef std::map<COLORREF, StylePtr> StylesMap;
    typedef CComPtr<ICommandHandler> CommandPtr;

    StylesMap       m_mStylesCache;
    eDisplayMode    m_eDisplayMode;

    CommandPtr      m_spDispModeBack;
    CommandPtr      m_spDispModePlain1x;
    CommandPtr      m_spDispModePlain2x;
    CommandPtr      m_spDispModePlain3x;
    CommandPtr      m_spDispModePlain4x;

    void InitCommands();

public:
    static LPCSTR GUID() { return "{FC52F539-8F7A-4145-8E33-01E8371D82E2}"; }

    CComPtr<IStyle> GetColorStyle(COLORREF clr);

    eDisplayMode DisplayMode() const { return m_eDisplayMode; }
    void DisplayMode(eDisplayMode val);

public:
    // IPlugin
    STDMETHOD(GetInfo)(BSTR* psDescription, BSTR* psAuthor, BSTR* psEmail, BSTR* psHomepage);

    // IStyleInformerCreator
    STDMETHOD(Create)(ISyntax* pSyntax, IStyleInformer** ppInformer);

    // IFrameEvents Methods
    STDMETHOD(onWorkspaceOpen)(VARIANT_BOOL bSaveState);

    // IMainMenuHandler Methods
    STDMETHOD(Init)(IMenuObject* pMenu, VARIANT_BOOL bUpdate);
    STDMETHOD(OnSubMenuUpdate)(IMenuObject* /*pMenu*/) { return S_FALSE; }
};