// Connect.cpp : Implementation of CConnect
#include "stdafx.h"
#include "AddIn_i.h"
#include <atlcoll.h>
#include "Connect.h"

#define DEFAULT_LANGUAGE 1031

_ATL_FUNC_INFO ClickEventInfo = { CC_STDCALL, VT_EMPTY, 2, { VT_DISPATCH, VT_BOOL|VT_BYREF } };

// event sink to handle button click events from Visio.
// Finally this simple "click" event is used as the least evil method.
class ClickEventRedirector :
	public IDispEventSimpleImpl<1, ClickEventRedirector, &__uuidof(Office::_CommandBarButtonEvents)>
{
public:
	ClickEventRedirector(IUnknownPtr punk) : m_punk(punk)
	{
		// advise to the event immediately
		DispEventAdvise(punk);
	}

	~ClickEventRedirector()
	{
		// unadvise on destruction
		DispEventUnadvise(m_punk);
	}

	// the event handler itself. 
	// Just redirect to the global command processor ("parameter" of the button contains the command itself)
	void __stdcall OnClick(IDispatch* pButton, VARIANT_BOOL* pCancel);

	// keep the reference to the item itself, otherwise VISIO destroys it for some unknown reason
	// probably because of it's "double-nature" user interface, and events are never fired.
	IUnknownPtr m_punk;

	BEGIN_SINK_MAP(ClickEventRedirector)
		SINK_ENTRY_INFO(1, __uuidof(Office::_CommandBarButtonEvents), 1, &ClickEventRedirector::OnClick, &ClickEventInfo)
	END_SINK_MAP()
};

struct CVisioConnect::Impl 
{
	int GetVisioVersion(Visio::IVApplicationPtr app)
	{
		static int result = -1;

		if (result == -1)
			result = StrToInt(app->GetVersion());

		return result;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	int GetAppLanguage()
	{
		IDispatchPtr disp_language_settings;
		if (m_app->get_LanguageSettings(&disp_language_settings))
			return DEFAULT_LANGUAGE;

		LanguageSettingsPtr language_settings;
		if (FAILED(disp_language_settings->QueryInterface(__uuidof(LanguageSettings), (void**)&language_settings)))
			return DEFAULT_LANGUAGE;

		int app_language = 0;
		if (FAILED(language_settings->get_LanguageID(msoLanguageIDUI, &app_language)))
			return DEFAULT_LANGUAGE;

		switch (app_language)
		{
		case 1033:
		case 1031:
		case 1049:
			return app_language;
		}

		return DEFAULT_LANGUAGE;
	}

	struct LanguageLock
	{
		int old_lcid;
		int old_langid;

		LanguageLock(int app_language)
		{
			HMODULE hKernel32 = GetModuleHandle(L"Kernel32.dll");

			typedef LANGID (WINAPI *FP_SetThreadUILanguage)(LANGID LangId);
			FP_SetThreadUILanguage pSetThreadUILanguage = (FP_SetThreadUILanguage)GetProcAddress(hKernel32, "SetThreadUILanguage");

			typedef LANGID (WINAPI *FP_GetThreadUILanguage)();
			FP_GetThreadUILanguage pGetThreadUILanguage = (FP_GetThreadUILanguage)GetProcAddress(hKernel32, "GetThreadUILanguage");

			old_lcid = GetThreadLocale();
			SetThreadLocale(app_language);

			old_langid = 0;
			if (pSetThreadUILanguage && pGetThreadUILanguage)
			{
				old_langid = pGetThreadUILanguage();
				pSetThreadUILanguage(app_language);
			}
		}

		~LanguageLock()
		{
			SetThreadLocale(old_lcid);

			HMODULE hKernel32 = GetModuleHandle(L"Kernel32.dll");

			typedef LANGID (WINAPI *FP_SetThreadUILanguage)(LANGID LangId);
			FP_SetThreadUILanguage pSetThreadUILanguage = (FP_SetThreadUILanguage)GetProcAddress(hKernel32, "SetThreadUILanguage");

			typedef LANGID (WINAPI *FP_GetThreadUILanguage)();
			FP_GetThreadUILanguage pGetThreadUILanguage = (FP_GetThreadUILanguage)GetProcAddress(hKernel32, "GetThreadUILanguage");

			if (pSetThreadUILanguage && pGetThreadUILanguage)
				pSetThreadUILanguage(old_langid);
		}
	};

	static CString LoadTextFile(UINT resource_id)
	{
		HMODULE hResources = _Module.GetResourceInstance();

		HRSRC rc = ::FindResource(
			hResources, MAKEINTRESOURCE(resource_id), L"TEXTFILE");

		LPWSTR content = static_cast<LPWSTR>(
			::LockResource(::LoadResource(hResources, rc)));

		DWORD content_length = 
			::SizeofResource(hResources, rc);

		return CString(content, content_length / 2);
	}

	Office::CommandBarPtr GetDrawingContextMenu(Office::_CommandBarsPtr cbs)
	{
		Office::CommandBarPtr result;
		cbs->get_Item(variant_t(L"Drawing Object Selected"), &result);
		return result;
	}

	#define ADDON_NAME			L"JumpToShapeAddin"
	#define MARKER_ADDON_NAME	L"QUEUEMARKEREVENT"

	/**-----------------------------------------------------------------------------
		Initializes menu or toolbar item's caption, icon, and command id.
	------------------------------------------------------------------------------*/
	void InitializeItem(Office::CommandBarControlPtr item, UINT command_id)
	{
		CString caption;
		caption.LoadString(command_id);
	    item->put_Caption(bstr_t(caption));

	    // The target action is marker add-on, since we work using MarkerEvent.
	    // For more information about that, see the article in MSDN: 
	    // http://msdn.microsoft.com/en-us/library/aa140366.aspx

		CString parameter;
		parameter.Format(L"%d", command_id);
	    item->put_Parameter(bstr_t(parameter));

		// Set unique tag, so that the command is not lost
		CString tag;
		tag.Format(L"%s_%d", ADDON_NAME, command_id);
		item->put_Tag(bstr_t(tag));

		m_buttons.Add(new ClickEventRedirector(item));
	}

	void FillMenuItems(long position, Office::CommandBarControlsPtr menu_items, HMENU popup_menu)
	{
		// For each items in the menu,
		bool begin_group = false;
		for (int i = 0; i < GetMenuItemCount(popup_menu); ++i)
		{
			HMENU sub_menu = GetSubMenu(popup_menu, i);

			// set item caption
			WCHAR item_caption[1024] = L"";
			GetMenuString(popup_menu, i, item_caption, 1024, MF_BYPOSITION);

			// if this item is actually a separator then process next item
			if (lstrlen(item_caption) == 0)
			{
				begin_group = true;
				continue;
			}

			// create new menu item.
			Office::CommandBarControlPtr menu_item_obj;
			menu_items->Add(
				variant_t(sub_menu ? long(Office::msoControlPopup) : long(Office::msoControlButton)), 
				vtMissing, 
				vtMissing, 
				position < 0 ? vtMissing : variant_t(position), 
				variant_t(true),
				&menu_item_obj);

			if (position > 0)
				++position;

			// obtain command id from menu
			UINT command_id = GetMenuItemID(popup_menu, i);

			// normal command; set up visio menu item
			InitializeItem(menu_item_obj, command_id);

			// if current item is first in a group, start new group
			if (begin_group)
			{
				menu_item_obj->put_BeginGroup(VARIANT_TRUE);
				begin_group = false;
			}

			// if this command has sub-menu
			if (sub_menu)
			{
				Office::CommandBarPopupPtr popup_menu_item_obj = menu_item_obj;

				Office::CommandBarControlsPtr controls;
				popup_menu_item_obj->get_Controls(&controls);

				FillMenuItems(-1, controls, sub_menu);
			}
		}
	}

	void FillMenu(long position, Office::CommandBarControlsPtr cbs, UINT menu_id)
	{
		HMENU menu = LoadMenu(_Module.GetResourceInstance(), MAKEINTRESOURCE(menu_id));

		FillMenuItems(position, cbs, GetSubMenu(menu, 0));
	}

	void CreateCommandBarsMenu(Visio::IVApplicationPtr app)
	{
		Office::_CommandBarsPtr cbs = app->CommandBars;

		Office::CommandBarPtr drawing_context_popup = GetDrawingContextMenu(cbs);

		Office::CommandBarControlsPtr controls;
		drawing_context_popup->get_Controls(&controls);

		FillMenu(1L, controls, IDR_MENU);
	}

	void DestroyCommandBarsMenu()
	{
		for (size_t i = 0; i < m_buttons.GetCount(); ++i)
			delete m_buttons[i];

		m_buttons.RemoveAll();
	}

	static bool FindItem(HWND hwnd_tree, HTREEITEM root, LPCWSTR name)
	{
		TVITEM item;

		WCHAR item_name[1024];
		item.hItem = root;
		item.mask = TVIF_TEXT;
		item.pszText = item_name;
		item.cchTextMax = 1024;

		if (TreeView_GetItem(hwnd_tree, &item))
		{
			if (!StrCmp(item_name, name))
			{
				TreeView_Select(hwnd_tree, root, TVGN_CARET);
				TreeView_EnsureVisible(hwnd_tree, root);
				return true;
			}
		}

		TreeView_Expand(hwnd_tree, root, TVE_EXPAND);

		HTREEITEM child = TreeView_GetChild(hwnd_tree, root);
		while (child)
		{
			if (FindItem(hwnd_tree, child, name))
				return true;

			child = TreeView_GetNextSibling(hwnd_tree, child);
		}

		TreeView_Expand(hwnd_tree, root, TVE_COLLAPSE);

		return false;
	}

	static HRESULT JumpToShape(Visio::IVApplicationPtr app)
	{
		using namespace Visio;

		IVWindowPtr window = app->ActiveWindow;

		IVWindowPtr drawin_explorer_window = window->GetWindows()->GetItemFromID(visWinIDDrawingExplorer);
		drawin_explorer_window->Visible = VARIANT_TRUE;

		IVSelectionPtr selection = window->Selection;
		if (selection->Count >= 1)
		{
			IVShapePtr shape = selection->GetItem(1L);

			HWND hwnd = (HWND) drawin_explorer_window->GetWindowHandle32();
			LPCWSTR shape_name = shape->Name;

			HWND hwnd_tree = GetWindow(hwnd, GW_CHILD);

			SendMessage(hwnd_tree, WM_SETREDRAW, 0, 0);

			if (FindItem(hwnd_tree, TreeView_GetRoot(hwnd_tree), shape_name))
				SetFocus(hwnd_tree);

			SendMessage(hwnd_tree, WM_SETREDRAW, 1, 0);
		}

		return S_OK;
	}

	void Create(IDispatch * pApplication, IDispatch * pAddInInst) 
	{
		pApplication->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_app);
		pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_addin);

		m_language = GetAppLanguage();

		if (GetVisioVersion(m_app) < 14)
		{
			LanguageLock lock(GetAppLanguage());
			CreateCommandBarsMenu(m_app);
		}
	}

	void Destroy() 
	{
		DestroyCommandBarsMenu();

		m_app = NULL;
		m_addin = NULL;
	}

	CString GetControlId(IDispatch* pControl)
	{
		IRibbonControlPtr control;
		pControl->QueryInterface(__uuidof(IRibbonControl), (void**)&control);

		CComBSTR bstr_control_id;
		if (FAILED(control->get_Tag(&bstr_control_id)))
			return S_OK;

		return static_cast<LPCWSTR>(bstr_control_id);
	}

	void OnRibbonButtonClicked(IDispatch * pControl) 
	{
		CString control_id = GetControlId(pControl);

		if (control_id == L"JumpToShape")
			JumpToShape(m_app);
	}

	VARIANT_BOOL IsRibbonButtonVisible(IDispatch * pControl)
	{
		CString control_id = GetControlId(pControl);

		VARIANT_BOOL result = VARIANT_TRUE;

		if (control_id == L"JumpToShape")
		{
			Visio::IVApplicationSettingsPtr app_settings;

			if (SUCCEEDED(m_app->get_Settings(&app_settings)))
				app_settings->get_DeveloperMode(&result);
		}

		return result;
	}

	CString GetRibbonLabel(IDispatch* pControl) 
	{
		CString control_id = GetControlId(pControl);

		LanguageLock lock(GetAppLanguage());

		CString result;
		if (control_id == L"JumpToShape")
			result.LoadString(IDS_JumpToShape);

		return result;
	}

	void SetupRibbon(IDispatch* disp) 
	{
		m_ribbon = disp;
	}

	Visio::IVApplicationPtr m_app;
	IDispatchPtr m_addin;
	Office::IRibbonUIPtr m_ribbon;

	DWORD m_language;

	CAtlArray<ClickEventRedirector*> m_buttons;
};

static void LogError(_com_error)
{
	// UNUSED_ALWAYS(e);
}

// the event handler itself. 
// Just redirect to the global command processor ("parameter" of the button contains the command itself)
void __stdcall ClickEventRedirector::OnClick(IDispatch* pButton, VARIANT_BOOL* pCancel)
{
	// Just in case. Visio seem to call this from some odd thread, without this all crashes down.
	try
	{
		Office::_CommandBarButtonPtr button;
		pButton->QueryInterface(__uuidof(Office::_CommandBarButton), (void**)&button);

		IDispatchPtr app;
		button->get_Application(&app);

		CVisioConnect::Impl::JumpToShape(app);
	}
	catch (_com_error& e)
	{
		LogError(e);
	}
}

// CConnect
STDMETHODIMP CVisioConnect::OnConnection(IDispatch *pApplication, ext_ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** custom)
{
	try
	{
		m_impl->Create(pApplication, pAddInInst);
	}
	catch (_com_error& e)
	{
		LogError(e);
	}

	return S_OK;
}

STDMETHODIMP CVisioConnect::OnDisconnection(ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	try
	{
		m_impl->Destroy();
	}
	catch (_com_error& e)
	{
		LogError(e);
	}

	return S_OK;
}

STDMETHODIMP CVisioConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CVisioConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CVisioConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CVisioConnect::GetCustomUI(BSTR RibbonID, BSTR * RibbonString)
{
	try
	{
		*RibbonString = Impl::LoadTextFile(IDR_RIBBON).AllocSysString();
	}
	catch (_com_error& e)
	{
		LogError(e);
	}

	return S_OK;
}

STDMETHODIMP CVisioConnect::OnRibbonButtonClicked (IDispatch * pControl)
{
	try
	{
		m_impl->OnRibbonButtonClicked(pControl);
	}
	catch (_com_error& e)
	{
		LogError(e);
	}

	return S_OK;
}


STDMETHODIMP CVisioConnect::IsRibbonButtonVisible(IDispatch * pControl, VARIANT_BOOL* pResult)
{
	*pResult = m_impl->IsRibbonButtonVisible(pControl);
	return S_OK;
}

STDMETHODIMP CVisioConnect::IsRibbonButtonEnabled(IDispatch * pControl, VARIANT_BOOL* pResult)
{
	*pResult = VARIANT_TRUE;
	return S_OK;
}

STDMETHODIMP CVisioConnect::OnRibbonLoad(IDispatch* disp)
{
	try
	{
		m_impl->SetupRibbon(disp);
	}
	catch (_com_error& e)
	{
		LogError(e);
	}

	return S_OK;
}

STDMETHODIMP CVisioConnect::GetRibbonLabel(IDispatch* pControl, BSTR *pbstrLabel)
{
	try
	{
		*pbstrLabel = m_impl->GetRibbonLabel(pControl).AllocSysString();
	}
	catch (_com_error& e)
	{
		LogError(e);
	}

	return S_OK;
}

CVisioConnect::CVisioConnect()
{
	m_impl = new Impl();
}

CVisioConnect::~CVisioConnect()
{
	delete m_impl;
}
