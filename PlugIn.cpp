#include <cstring>
#include <atlsafe.h>
#include <boost/regex.hpp>
#include <boost/regex/mfc.hpp> 
#include "PlugIn.h"

#ifndef CLR_NONE
	#define CLR_NONE                0xFFFFFFFFL
	#define CLR_DEFAULT             0xFF000000L
#endif

#ifndef RGBA
	#define RGBA(r, g, b, a)	MAKELONG(MAKEWORD(r, g), MAKEWORD(b, a))
#endif

STDMETHODIMP CLiveColors::GetInfo(BSTR* psDescription, BSTR* psAuthor, BSTR* psEmail, BSTR* psHomepage)
{
	if (psDescription)	*psDescription	= OLE2BSTR(OLESTR("Show color values for HEX or RGB defined colors inline in source code"));
	if (psAuthor)		*psAuthor		= OLE2BSTR(OLESTR("HippoEDIT"));
	if (psEmail)		*psEmail		= OLE2BSTR(OLESTR("supportbox@hippoedit.com"));
	if (psHomepage)		*psHomepage		= OLE2BSTR(OLESTR("http://wiki.hippoedit.com/plugins/live-colors"));

	return S_OK;
}

class ATL_NO_VTABLE CStyleInformer : public CPluginImpl<IStyleInformer>,
									 public CComObjectRootEx<CComMultiThreadModel>
{
	CLiveColors*				m_pPlugIn;
	CComPtr<IStyleCollector>	m_pCollector;
	boost::tregex				m_r;

public:
	CStyleInformer() { }
	virtual ~CStyleInformer() { }

	STDMETHOD(Init)(CLiveColors* pPlugin) { m_pPlugIn = pPlugin; return S_OK; }

	BEGIN_COM_MAP(CStyleInformer)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(IStyleInformer)
	END_COM_MAP()

protected:
	BYTE adopt_hex(const boost::sub_match<TCHAR const*>& submatch)
	{
		int bt = 0;
		if (submatch.matched)
		{
			if (submatch.length() == 1)
			{
				std::basic_string<TCHAR> str =  submatch.str() + submatch.str();
				_stscanf_s(str.c_str(),_T("%x"),&bt);				
			}
			else
			{
				_stscanf_s(submatch.str().c_str(),_T("%x"),&bt);				
			}
		}

		return LOBYTE(bt);
	}

	BYTE adopt(const boost::sub_match<TCHAR const*>& submatch)
	{
		int bt = 0;
		if (submatch.matched)
			_stscanf_s(submatch.str().c_str(),_T("%d"),&bt);				

		return LOBYTE(bt);
	}

public:
	// IStyleInformer
	STDMETHOD(get_JobID)(BSTR* psJobID)
	{
		if (!psJobID) return E_INVALIDARG;
		*psJobID = ::SysAllocString(L"live_colors_job");
		return S_OK;
	}

	STDMETHOD(OnJobStart)(IDocumentData* pDocument, LINE_T* nLineFrom, LINE_T* nLineTo, HANDLE_PTR* /*pUserData*/)
	{
		if (!pDocument) 
			return E_INVALIDARG;
		
		// RGB "rgb\(([01]?\d\d?|2[0-4]\d|25[0-5])\,([01]?\d\d?|2[0-4]\d|25[0-5])\,([01]?\d\d?|2[0-4]\d|25[0-5])\)"
		// WEB "#?([a-f]|[A-F]|[0-9]){3}(([a-f]|[A-F]|[0-9]){3})?"
		// ALL "(?:rgb|RGB|RGBA)\(([01]?\d\d?|2[0-4]\d|25[0-5])\,([01]?\d\d?|2[0-4]\d|25[0-5])\,([01]?\d\d?|2[0-4]\d|25[0-5])(?:,([01]?\d\d?|2[0-4]\d|25[0-5]))*\)|#((?:[a-f]|[A-F]|[0-9]){2})((?:[a-f]|[A-F]|[0-9]){2})((?:[a-f]|[A-F]|[0-9]){2})?((?:[a-f]|[A-F]|[0-9]){2})?|#([a-f]|[A-F]|[0-9])([a-f]|[A-F]|[0-9])([a-f]|[A-F]|[0-9])"

		if (m_r.empty())
		{
			CString sRegexp = _T("(?:rgb|rgba)\\(\\s*([01]?\\d\\d?|2[0-4]\\d|25[0-5])\\s*,\\s*([01]?\\d\\d?|2[0-4]\\d|25[0-5])\\s*,\\s*([01]?\\d\\d?|2[0-4]\\d|25[0-5])(?:\\s*,\\s*([01]?\\d\\d?|2[0-4]\\d|25[0-5]))?\\s*\\)");
			sRegexp += _T("|(?|(?:rgb|rgba)\\(\\s*0x((?:[a-f]|[0-9]){2})\\s*,\\s*0x((?:[a-f]|[0-9]){2})\\s*,\\s*0x((?:[a-f]|[0-9]){2})(?:\\s*,\\s*0x((?:[a-f]|[0-9]){2}))?\\s*\\)");
			sRegexp += _T("|#((?:[a-f]|[0-9]){2})((?:[a-f]|[0-9]){2})((?:[a-f]|[0-9]){2})?((?:[a-f]|[0-9]){2})?");
			sRegexp += _T("|#([a-f]|[0-9])([a-f]|[0-9])([a-f]|[0-9]))");

			m_r.set_expression(sRegexp, boost::regex_constants::normal|boost::regex_constants::icase|boost::regex_constants::optimize);
		}

		return S_FALSE;
	}

	STDMETHOD(ProcessLine)(LINE_T nLine, BSTR pszLine, IStyleCollector* pCollector, HANDLE_PTR* /*pUserData*/)
	{
		bool bFound = false;
		for(boost::tregex_iterator iter = boost::make_regex_iterator(pszLine, m_r), iend; iter != iend; ++iter)
		{
			const INT_PTR nLength = iter->length();
			const INT_PTR nPos = (*iter)[0].first - pszLine;
			if ((!nPos || (!_istalnum(pszLine[nPos - 1]) && pszLine[nPos - 1] != _T('&'))) && (!_istalnum(pszLine[nPos + nLength])))
			{
				COLORREF clr = CLR_NONE;
				if ((*iter)[1].matched) clr = RGBA(adopt((*iter)[1]), adopt((*iter)[2]), adopt((*iter)[3]), adopt((*iter)[4]));
				else if ((*iter)[5].matched) clr = RGBA(adopt_hex((*iter)[5]), adopt_hex((*iter)[6]), adopt_hex((*iter)[7]), adopt_hex((*iter)[8]));
				if (clr != CLR_NONE && clr != CLR_DEFAULT)
				{
					bFound = true;
					if (CComPtr<IStyle> pStyle = m_pPlugIn->GetColorStyle(clr))
						pCollector->Add(pStyle, nLine, nPos, nPos + nLength);
				}
			}
		}

		return bFound?S_OK:S_FALSE;
	}

	STDMETHOD(OnJobEnd)(VARIANT_BOOL bCanceled, IStyleCollector* pCollector, HANDLE_PTR* /*pUserData*/)
	{
		return S_FALSE;
	}
};

STDMETHODIMP CLiveColors::Create(ISyntax* pSyntax, IStyleInformer** pInformer)
{
	CComObject<CStyleInformer>* spInformer = NULL;
	CComObject<CStyleInformer>::CreateInstance(&spInformer);
	spInformer->Init(this);
	return spInformer->QueryInterface(__uuidof(IStyleInformer), (void**)pInformer);
}

CComPtr<IStyle> CLiveColors::GetColorStyle( COLORREF clr )
{
	ObjectLock lock(this);

	StylesMap::const_iterator iter = m_mStylesCache.find(clr);
	if (iter != m_mStylesCache.end()) return iter->second;

	wchar_t chBuffer[10];
	swprintf_s(chBuffer, _T("#%02X%02X%02X%02X"), LOBYTE(LOWORD(clr)), HIBYTE(LOWORD(clr)), LOBYTE(HIWORD(clr)), HIBYTE(HIWORD(clr)));
	CStringW sName; sName.Format(L"Color %s", chBuffer);

	CComPtr<IStyle> pStyle;
	if (SUCCEEDED(Application()->Style(CComBSTR(chBuffer), CComBSTR(sName), CComBSTR(sName), CComVariant(L"def"), &pStyle)) && pStyle)
	{
		// first reset to default
		pStyle->put_UnderlineStyle(eUnderlineStyleNone);
		pStyle->put_UnderlineColor(CComVariant(CLR_NONE));
		pStyle->put_BackColor(CComVariant(CLR_NONE));

		switch (DisplayMode())
		{
		case eDisplayModeBackground:
			pStyle->put_BackColor(CComVariant(clr));
			break;
		case eDisplayModeUnderline1x:
			pStyle->put_UnderlineStyle(eUnderlineStylePlain);
			pStyle->put_UnderlineColor(CComVariant(clr));
			break;
		case eDisplayModeUnderline2x:
			pStyle->put_UnderlineStyle(eUnderlineStylePlain2x);
			pStyle->put_UnderlineColor(CComVariant(clr));
			break;
		case eDisplayModeUnderline3x:
			pStyle->put_UnderlineStyle(eUnderlineStylePlain3x);
			pStyle->put_UnderlineColor(CComVariant(clr));
			break;
		case eDisplayModeUnderline4x:
			pStyle->put_UnderlineStyle(eUnderlineStylePlain4x);
			pStyle->put_UnderlineColor(CComVariant(clr));
			break;
		}
	}

	return m_mStylesCache.insert(StylesMap::value_type(clr, CAdapt<Style>(pStyle))).first->second;
}

STDMETHODIMP CLiveColors::onWorkspaceOpen( VARIANT_BOOL /*bSaveState*/ )
{
	ObjectLock lock(this);

	CComPtr<ISettingsStorage> spStorage;
	Application()->GetStorage(eStorageTypeWorkspace, NULL, eAccessTypeRead, &spStorage.p);

	m_eDisplayMode = (eDisplayMode)ReadStorage(spStorage, "mode",  (DWORD)DisplayMode());
	m_mStylesCache.clear();

	return S_OK;
}

void CLiveColors::DisplayMode( eDisplayMode val )
{
	if (DisplayMode() != val)
	{
		ObjectLock lock(this);

		m_eDisplayMode = val;

		CComPtr<ISettingsStorage> spStorage;
		if (SUCCEEDED(Application()->GetStorage(eStorageTypeWorkspace, NULL, eAccessTypeWrite, &spStorage.p)) && spStorage)
			WriteStorage(spStorage, "mode",  (DWORD)DisplayMode());

		m_mStylesCache.clear();

		// force views invalidation
		Application()->UpdateAll();
	}
}

class CCommandVisualMode : public CCommandHandler
{
protected:
	CLiveColors*				m_pPlugin;
	CLiveColors::eDisplayMode	m_eDisplayMode;

public:
	static CComPtr<CCommandVisualMode> CreateInstanceEx(CLiveColors* pPlugin, LPCWSTR pszID, LPCWSTR pszTitle, LPCWSTR pszPrompt, CLiveColors::eDisplayMode mode) 
	{	
		CComObject<CCommandVisualMode>* pCommand = NULL;
		CComObject<CCommandVisualMode>::CreateInstance(&pCommand);
		if (pCommand)
		{
			pCommand->m_pPlugin			= pPlugin;
			pCommand->m_eDisplayMode	= mode;
			pCommand->Init(pszID, pszTitle, pszPrompt);
			return pCommand;
		}
		return NULL;
	}

	STDMETHOD(get_Radio)(VARIANT_BOOL* pbRadio) 
	{ 
		if (!pbRadio) return E_INVALIDARG;
		*pbRadio = (m_pPlugin->DisplayMode() == m_eDisplayMode)?VARIANT_TRUE:VARIANT_FALSE;
		return S_OK; 
	}

	STDMETHOD(Execute)() 
	{  
		m_pPlugin->DisplayMode(m_eDisplayMode);
		return S_OK;
	}
};

void CLiveColors::InitCommands()
{
	if (!m_spDispModeBack)
	{
		m_spDispModeBack	= CCommandVisualMode::CreateInstanceEx(this, _T("LiveColors.ShowAsBackground"), _T("Show As Background"), _T("Visualize color as text background"), eDisplayModeBackground);
		m_spDispModePlain1x = CCommandVisualMode::CreateInstanceEx(this, _T("LiveColors.ShowAsUnderline1x"), _T("Show As 1 px Single Underline"), _T("Visualize color as 1 px underline"), eDisplayModeUnderline1x);
		m_spDispModePlain2x = CCommandVisualMode::CreateInstanceEx(this, _T("LiveColors.ShowAsUnderline2x"), _T("Show As 2 px Single Underline"), _T("Visualize color as 2 px underline"), eDisplayModeUnderline2x);
		m_spDispModePlain3x = CCommandVisualMode::CreateInstanceEx(this, _T("LiveColors.ShowAsUnderline3x"), _T("Show As 3 px Single Underline"), _T("Visualize color as 3 px underline"), eDisplayModeUnderline3x);
		m_spDispModePlain4x = CCommandVisualMode::CreateInstanceEx(this, _T("LiveColors.ShowAsUnderline4x"), _T("Show As 4 px Single Underline"), _T("Visualize color as 4 px underline"), eDisplayModeUnderline4x);
	}
}

STDMETHODIMP CLiveColors::Init( IMenuObject* pMenu, VARIANT_BOOL bUpdate )
{
	if (!pMenu) return E_INVALIDARG;

	if (bUpdate == VARIANT_FALSE)
	{
		CComPtr<IMenuLocation> spLocation;
		if (SUCCEEDED(pMenu->GetMenuItemPos(CComBSTR("View.ZoomIn"), &spLocation)) && spLocation)
		{
			CComPtr<IMenuObject> spSubMenu;
			LONG nPosition = -1;
			spLocation->get_Menu(&spSubMenu);
			spLocation->get_Position(&nPosition);
			if (spSubMenu && nPosition != -1)
			{
				CComPtr<IMenuObject> pLiveColorsMenu;
				spSubMenu->InsertSubMenu(nPosition, CComBSTR(L"&Live Colors"), &pLiveColorsMenu);
				spSubMenu->InsertItem(nPosition + 1, NULL, CComVariant());
				if (pLiveColorsMenu)
				{
					InitCommands();
					pLiveColorsMenu->AddItem(NULL, CComVariant(m_spDispModeBack),    eItemPosDesigned);
					pLiveColorsMenu->AddItem(NULL, CComVariant(m_spDispModePlain1x), eItemPosDesigned);
					pLiveColorsMenu->AddItem(NULL, CComVariant(m_spDispModePlain2x), eItemPosDesigned);
					pLiveColorsMenu->AddItem(NULL, CComVariant(m_spDispModePlain3x), eItemPosDesigned);
					pLiveColorsMenu->AddItem(NULL, CComVariant(m_spDispModePlain4x), eItemPosDesigned);
				}
				return S_OK;
			}
		}
	}

	return S_FALSE;
}