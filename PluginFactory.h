#pragma once
#include <atlstr.h>
#include <map>

#pragma comment(lib,"version")

/************************************************************************/
/* Class constructor                                                    */
/************************************************************************/
namespace plugin_utils {
    template<typename derived_t, typename target_t = derived_t>
    class class_initializer
    {
        struct helper
        {
            helper() { target_t::static_ctor(); }
            ~helper() { target_t::static_dtor(); }
        };

        static helper helper_;

        static void use_helper() { (void)helper_; }
        template<void(*)()> struct helper2 {};

        helper2<&class_initializer::use_helper> helper2_;

        virtual void use_helper2() { (void)helper_; }

        // default implementation does nothing, but allow you to define only destructor or constructor
        static void static_ctor() {} // static constructor
        static void static_dtor() {} // static destructor
    };

    template<typename derived_t, typename target_t>
    typename class_initializer<derived_t, target_t>::helper class_initializer<derived_t, target_t>::helper_;
}

template <const GUID* plibid = &LIBID_HippoEditLib, WORD wMajor = 1, WORD wMinor = 0>
class ATL_NO_VTABLE CPluginFactory :
    public CComObjectRootEx<CComMultiThreadModel>,
    public IDispatchImpl<IPluginFactory, &__uuidof(IPluginFactory), plibid, wMajor, wMinor>
{
protected:
    typedef HRESULT (*LPCreator)(IPlugin**);
    typedef std::map<CStringA, LPCreator>   CreatorsMap;
    static CreatorsMap& GetMap()
    {
        static CAutoPtr<CreatorsMap> s_mRegisteredPlugins;
        if (!s_mRegisteredPlugins) s_mRegisteredPlugins.Attach(new CreatorsMap());
        return *s_mRegisteredPlugins.m_p;
    }

public:

    typedef CPluginFactory<plibid, wMajor, wMinor> TMe;

    template <class T>
    struct ATL_NO_VTABLE TPlugin abstract : public CComObjectRootEx<CComMultiThreadModel>,
                                            public IDispatchImpl<IPlugin, &__uuidof(IPlugin), plibid, wMajor, wMinor>,
                                            public plugin_utils::class_initializer<T>
    {
        CComPtr<IApplication>   m_pApplication;

        const CComPtr<IApplication>& Application() const { return m_pApplication; }

        static HRESULT CreateInstance(IPlugin** ppv)
        {
            CComObjectCached<T>* spPlugin = NULL;
            CComObjectCached<T>::CreateInstance(&spPlugin);
            return spPlugin->QueryInterface(__uuidof(IPlugin), (void**)ppv);
        }

        static void static_ctor()
        {
            CPluginFactory::RegisterPlugin(T::GUID(), T::CreateInstance);
        }

        STDMETHOD(get_GUID)(BSTR* psGUID)
        {
            if (!psGUID) return E_INVALIDARG;
            *psGUID = ::SysAllocString(CA2W(T::GUID()));
            return S_OK;
        }

        STDMETHOD(get_Application)(IApplication** ppApplication)
        {
            if (!ppApplication) return E_INVALIDARG;
            return m_pApplication.QueryInterface(ppApplication);
        }

        STDMETHOD(Init)(IApplication* pApplication)
        {
            if (!pApplication) return E_INVALIDARG;
            m_pApplication = pApplication;
            return S_OK;
        }

        STDMETHOD(get_Version)(BSTR* psVersion)
        {
            if (!psVersion)
                return E_INVALIDARG;

            *psVersion = NULL;

            TCHAR szVersionFile[MAX_PATH + 1] = {0};
            if (GetModuleFileName(_AtlBaseModule.GetModuleInstance(), szVersionFile, MAX_PATH))
            {
                DWORD  verHandle = NULL;
                if (DWORD  verSize = GetFileVersionInfoSize( szVersionFile, &verHandle))
                {
                    CTempBuffer<char, 1024> verData(verSize);
                    if (GetFileVersionInfo( szVersionFile, verHandle, verSize, verData))
                    {
                        UINT   size      = 0;
                        LPBYTE lpBuffer  = NULL;
                        if (VerQueryValue(verData,_T("\\"), (VOID FAR* FAR*)&lpBuffer, &size) && size)
                        {
                            VS_FIXEDFILEINFO *verInfo = (VS_FIXEDFILEINFO *)lpBuffer;
                            TCHAR szVersion[2000];
                            wsprintf(szVersion, _T("%d.%d.%d.%d"), HIWORD(verInfo->dwProductVersionMS), LOWORD(verInfo->dwProductVersionMS),
                                HIWORD(verInfo->dwProductVersionLS), LOWORD(verInfo->dwProductVersionLS));
                            *psVersion = OLE2BSTR(szVersion);
                            return S_OK;
                        }
                    }
                }
            }

            return E_FAIL;
        }


        STDMETHOD(get_Name)(BSTR* psName)
        {
            if (!psName)
                return E_INVALIDARG;

            *psName = NULL;

            TCHAR szVersionFile[MAX_PATH + 1] = {0};
            if (GetModuleFileName(_AtlBaseModule.GetModuleInstance(), szVersionFile, MAX_PATH))
            {
                DWORD  verHandle = NULL;
                if (DWORD  verSize = GetFileVersionInfoSize( szVersionFile, &verHandle))
                {
                    CTempBuffer<char, 1024> verData(verSize);
                    if (GetFileVersionInfo( szVersionFile, verHandle, verSize, verData))
                    {
                        DWORD *pdwTranslation;
                        UINT nLength;
                        if (::VerQueryValue(verData, _T("\\VarFileInfo\\Translation"), (void**)&pdwTranslation, &nLength))
                        {
                            TCHAR szKey[2000];
                            wsprintf(szKey, _T("\\StringFileInfo\\%04x%04x\\ProductName"), LOWORD(*pdwTranslation), HIWORD(*pdwTranslation));

                            LPTSTR lpszValue;
                            if (::VerQueryValue(verData, szKey, (void**) &lpszValue, &nLength))
                            {
                                *psName = OLE2BSTR(lpszValue);
                                return S_OK;
                            }
                        }
                    }
                }
            }

            return E_FAIL;
        }

        STDMETHOD(UnInit)()
        {
            m_pApplication.Release();
            return S_OK;
        }
    };

public:
    CPluginFactory() {};
    virtual ~CPluginFactory() {};

    BEGIN_COM_MAP(CPluginFactory)
        COM_INTERFACE_ENTRY(IPluginFactory)
    END_COM_MAP()

    // IPluginFactory Methods
public:
    STDMETHOD(GetPluginList)(unsigned long celt, BSTR * ppGUIDs, unsigned long * pceltFetched)
    {
        // provide count of registered plug-ins
        if (!celt && ppGUIDs == NULL)
        {
            if (!pceltFetched)
                return E_INVALIDARG;
            *pceltFetched = (long)GetMap().size();
            return S_OK;
        }

        if (!ppGUIDs || !pceltFetched)
            return E_INVALIDARG;

        *pceltFetched = min(celt, static_cast<ULONG>(GetMap().size()));
        size_t nCounter = 0;
        for (CreatorsMap::const_iterator iter = GetMap().begin(), iend = GetMap().end(); iter != iend && nCounter != *pceltFetched; ++iter, ++nCounter )
            ppGUIDs[nCounter] = iter->first.AllocSysString();

        return (celt == *pceltFetched)?S_OK:S_FALSE;
    }

    STDMETHOD(CreatePlugin)(BSTR psGUID, IPlugin ** ppPlugin)
    {
        CreatorsMap::const_iterator iter = GetMap().find(CStringA(psGUID));
        if (iter != GetMap().end())
            return iter->second(ppPlugin);
        return E_FAIL;
    }

    // static plug-in registration
    static HRESULT RegisterPlugin(LPCSTR pszGUID, LPCreator pfCreator)
    {
        return GetMap().insert(CreatorsMap::value_type(pszGUID, pfCreator)).second?S_OK:S_FALSE;
    }

    static CComPtr<IPluginFactory> CreateInstanceEx()
    {
        CComObject<TMe>* spFactory = NULL;
        CComObject<TMe>::CreateInstance(&spFactory);
        return spFactory;
    }
};

typedef CPluginFactory<> CPluginFactoryExt;

template <const GUID* plibid = &LIBID_HippoEditLib, WORD wMajor = 1, WORD wMinor = 0>
struct ATL_NO_VTABLE CFrameEventsImpl : public IDispatchImpl<IFrameEvents, &__uuidof(IFrameEvents), plibid, wMajor, wMinor>
{
    //////////////////////////////////////////////////////////////////////////
    // IFrameEvents Methods
    STDMETHOD(onDocumentListUpdate)() { return S_FALSE; }
    STDMETHOD(onWorkspaceOpen)(VARIANT_BOOL /*bSaveState*/) { return S_FALSE; }
    STDMETHOD(onWorkspaceClose)(VARIANT_BOOL /*bSaveState*/) { return S_FALSE; }
    STDMETHOD(onJobFinished)(IDocument* /*pDocument*/, BSTR /*sJobID*/, long /*nLineFrom*/, long /*nLineTo*/) { return S_FALSE; }
    STDMETHOD(onIdle)() { return S_FALSE; }
    STDMETHOD(onDocumentStateUpdate)(IDocument* /*pDocument*/) { return S_FALSE; }
    STDMETHOD(onDocumentNameChange)(BSTR /*pszOldName*/, BSTR /*pszNewName*/, VARIANT_BOOL /*bRename*/) { return S_FALSE; }
    STDMETHOD(CanCloseWorkspace)(VARIANT_BOOL* pbResult) { if (!pbResult) return E_INVALIDARG; *pbResult = VARIANT_TRUE; return S_OK; }
    STDMETHOD(CanCloseApplication)(VARIANT_BOOL* pbResult) { return CanCloseWorkspace(pbResult); }
    STDMETHOD(onModifiedChange)(IDocument* /*pDocument*/, VARIANT_BOOL /*bModified*/)  { return S_FALSE; }
    STDMETHOD(onDocumentClose)(IDocument* /*pDocument*/)  { return S_FALSE; }
    STDMETHOD(onDocumentOpen)(IDocument* /*pDocument*/) { return S_FALSE; }
    STDMETHOD(onDocumentSave)(IDocument* /*pDocument*/) { return S_FALSE; }
    STDMETHOD(onNewDocument)(IDocument* /*pDocument*/) { return S_FALSE; }
    STDMETHOD(onSyntaxChange)(IDocument* /*pDocument*/, ISyntax* /*pOldSyntax*/) { return S_FALSE; }
    STDMETHOD(onCursorPosChange)(IView* /*pView*/, IPosition* /*pPos*/) { return S_FALSE; }
    STDMETHOD(onTextInsert)(IDocument* /*pDocument*/, IPosition* /*pPos*/, BSTR /*strText*/, eAction /*action*/) { return S_FALSE; }
    STDMETHOD(onTextFormat)(IDocument* /*pDocument*/, IRange* /*pRange*/, VARIANT_BOOL /*bIndent*/, VARIANT_BOOL* /*pbResult*/) { return S_FALSE; }
    STDMETHOD(onQuickInfo)(IView* /*pView*/, IRange* /*pSelection*/, VARIANT_BOOL* /*pbProcessed*/) { return S_FALSE; }
    STDMETHOD(onCompletion)(IView* /*pView*/, IRange* /*pSelection*/, VARIANT_BOOL* /*pbProcessed*/) { return S_FALSE; }
    STDMETHOD(onEditOperation)(IDocument* /*pDocument*/, eAction /*action*/) { return S_FALSE; }
    STDMETHOD(onFileDrop)(IView* /*pView*/, IPosition* pDropPos, long nDropEffect, VARIANT /*pFiles*/, VARIANT_BOOL* /*pbProcessed*/) { return S_FALSE; }
    STDMETHOD(onTextDrop)(IView* /*pView*/, IPosition* pDropPos, long nDropEffect, BSTR /*sText*/, VARIANT_BOOL* /*pbProcessed*/) { return S_FALSE; }
    STDMETHOD(onSettingsChange)(ISettings* /*pSettings*/) { return S_FALSE; }
    STDMETHOD(onDocumentSwitch)(IDocument* /*pActiveDocument*/) { return S_FALSE; }
    STDMETHOD(onFocusSet)(IView* /*pView*/) { return S_FALSE; }
    STDMETHOD(onFocusLost)(IView* /*pView*/) { return S_FALSE; }
    STDMETHOD(onScroll)(IView* /*pView*/, ULONG /*nScrollX*/, ULONG /*nScrollY*/) { return S_FALSE; }
    STDMETHOD(onSelectionChange)(IView* /*pView*/, IRange* /*pSelection*/, VARIANT_BOOL /*bBlockMode*/) { return S_FALSE; }
};

typedef CFrameEventsImpl<> CFrameEventsImplExt;

static inline DWORD ReadStorage(const CComPtr<ISettingsStorage>& spStorage, LPCSTR pszValueName, DWORD defValue = 0)
{
    CComVariant var;
    if (spStorage && spStorage->read(CComBSTR(pszValueName), &var) == S_OK)
        return var.uintVal;
    return defValue;
}

static inline long ReadStorage(const CComPtr<ISettingsStorage>& spStorage, LPCSTR pszValueName, long defValue = 0)
{
    CComVariant var;
    if (spStorage && spStorage->read(CComBSTR(pszValueName), &var) == S_OK)
        return var.uintVal;
    return defValue;
}

static inline bool ReadStorage(const CComPtr<ISettingsStorage>& spStorage, LPCSTR pszValueName, bool defValue = false)
{
    CComVariant var;
    if (spStorage && spStorage->read(CComBSTR(pszValueName), &var) == S_OK)
        return var.boolVal == VARIANT_TRUE;
    return defValue;
}

static inline CString ReadStorage(const CComPtr<ISettingsStorage>& spStorage, LPCSTR pszValueName, const CString& defValue = _T(""))
{
    CComVariant var;
    if (spStorage && spStorage->read(CComBSTR(pszValueName), &var) == S_OK)
        return CString(var.bstrVal);
    return defValue;
}

static inline HRESULT WriteStorage(const CComPtr<ISettingsStorage>& spStorage, LPCSTR pszValueName, DWORD Value)
{
    return spStorage->write(CComBSTR(pszValueName), CComVariant(Value));
}

static inline HRESULT WriteStorage(const CComPtr<ISettingsStorage>& spStorage, LPCSTR pszValueName, long Value)
{
    return spStorage->write(CComBSTR(pszValueName), CComVariant(Value));
}

static inline HRESULT WriteStorage(const CComPtr<ISettingsStorage>& spStorage, LPCSTR pszValueName, const CString& Value)
{
    return spStorage->write(CComBSTR(pszValueName), CComVariant(Value));
}

static inline HRESULT WriteStorage(const CComPtr<ISettingsStorage>& spStorage, LPCSTR pszValueName, bool Value)
{
    return spStorage->write(CComBSTR(pszValueName), CComVariant(Value));
}