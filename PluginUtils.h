//////////////////////////////////////////////////////////////////////////
// This file is part of HippoEDIT application.
// Copyright 2004-2010 HippoEDIT.com
// All rights reserved.
//
//    web:    http://www.hippoedit.com
//    mailto: salesbox@hippoedit.com
//
//////////////////////////////////////////////////////////////////////////

#pragma once
#include "dispatcheximpl.h"
#include "PluginFactory.h"

template <typename I, const GUID* plibid = &LIBID_HippoEditLib, WORD wMajor = 1, WORD wMinor = 0>
struct ATL_NO_VTABLE CPluginImpl : public IDispatchImpl<I, &__uuidof(I), plibid, wMajor, wMinor> {};

template <typename I, const GUID* plibid = &LIBID_HippoEditLib, WORD wMajor = 1, WORD wMinor = 0>
struct ATL_NO_VTABLE CPluginImplEx : public Scripting::IDispatchImplEx< I, &__uuidof(I), plibid, wMajor, wMinor > {};

class ATL_NO_VTABLE CCommandHandler : public CComObjectRootEx<CComSingleThreadModel>, public CPluginImplEx< ICommandHandler >
{
protected:
    CString m_sName;
    CString m_sTitle;
    CString m_sPrompt;

public:
    CCommandHandler() {}
    ~CCommandHandler() {}

    BEGIN_COM_MAP(CCommandHandler)
        COM_INTERFACE_ENTRY(IDispatchEx)
        COM_INTERFACE_ENTRY(ICommandHandler)
    END_COM_MAP()

public:
    void Init(const CString& sName, const CString& sTitle, const CString& sPrompt)
    {
        m_sName   = sName;
        m_sTitle  = sTitle;
        m_sPrompt = sPrompt;
    }

protected:
    virtual const CString& Name() const { return m_sName; }
    virtual const CString& Title() const { return m_sTitle; }
    virtual const CString& Prompt() const { return m_sPrompt; }

    HRESULT CopyVal(const CString& sFrom, BSTR* pVal)
    {
        if (!pVal) return E_INVALIDARG;
        *pVal = sFrom.AllocSysString();
        return S_OK;
    }

    HRESULT CopyVal(const BSTR sFrom, CString& sVal)
    {
        sVal = CString(sFrom);
        return S_OK;
    }

    // ICommandHandler Methods
public:
    STDMETHOD(get_Name)(BSTR* psID) { return CopyVal(Name(), psID); }
    STDMETHOD(get_Title)(BSTR* pbstrTitle) { return CopyVal(Title(), pbstrTitle); }
    STDMETHOD(put_Title)(BSTR strTitle) { return CopyVal(strTitle, m_sTitle); }
    STDMETHOD(get_Prompt)(BSTR* pbstrPrompt) { return CopyVal(Prompt(), pbstrPrompt); }
    STDMETHOD(put_Prompt)(BSTR strPrompt) { return CopyVal(strPrompt, m_sPrompt); }

    STDMETHOD(get_Enabled)(VARIANT_BOOL* pbEnabled)
    {
        if (!pbEnabled) return E_INVALIDARG;
        *pbEnabled = VARIANT_TRUE;
        return S_OK;
    }

    STDMETHOD(put_Enabled)(VARIANT_BOOL bEnabled = VARIANT_TRUE)
    {
        return E_NOTIMPL;
    }

    STDMETHOD(get_Checked)(LONG* pnChecked)
    {
        if (!pnChecked) return E_INVALIDARG;
        *pnChecked = 0;
        return S_OK;
    }

    STDMETHOD(put_Checked)(long nChecked = 1)
    {
        return E_NOTIMPL;
    }

    STDMETHOD(get_Radio)(VARIANT_BOOL* pbRadio)
    {
        if (!pbRadio) return E_INVALIDARG;
        *pbRadio = VARIANT_FALSE;
        return S_OK;
    }

    STDMETHOD(put_Radio)(VARIANT_BOOL bRadio = VARIANT_TRUE)
    {
        return E_NOTIMPL;
    }

    STDMETHOD(get_Image)(LONG* pnImage)
    {
        if (!pnImage) return E_INVALIDARG;
        return S_FALSE;
    }

    STDMETHOD(put_Image)(long /*nImage*/)
    {
        return E_NOTIMPL;
    }

    STDMETHOD(get_TextColor)(OLE_COLOR* pclr)
    {
        if (!pclr) return E_INVALIDARG;
        return S_FALSE;
    }

    STDMETHOD(put_TextColor)(OLE_COLOR /*clr*/)
    {
        return E_NOTIMPL;
    }

    STDMETHOD(get_BackColor)(OLE_COLOR* pclr)
    {
        if (!pclr) return E_INVALIDARG;
        return S_FALSE;
    }

    STDMETHOD(put_BackColor)(OLE_COLOR /*clr*/)
    {
        return E_NOTIMPL;
    }

    STDMETHOD(Execute)() {  return E_NOTIMPL;  }
};