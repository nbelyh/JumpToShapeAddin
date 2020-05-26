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
		long app_language = 0;
		if (FAILED(m_app->get_Language(&app_language)))
			return DEFAULT_LANGUAGE;

		CComVariant v_disp_language_settings;
		CComDispatchDriver disp = m_app;
		if (SUCCEEDED(disp.GetPropertyByName(L"LanguageSettings", &v_disp_language_settings)))
		{
			LanguageSettingsPtr language_settings;
			if (SUCCEEDED(V_DISPATCH(&v_disp_language_settings)->QueryInterface(__uuidof(LanguageSettings), (void**)&language_settings)))
			{
				int ui_language = app_language;
				if (SUCCEEDED(language_settings->get_LanguageID(msoLanguageIDUI, &ui_language)))
					app_language = ui_language;
			}
		}

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

	void FillMenuItem(Office::CommandBarControlsPtr menu_items, UINT command_id)
	{
		// create new menu item.
		Office::CommandBarControlPtr menu_item_obj;
		menu_items->Add(
			variant_t(long(Office::msoControlButton)),
			vtMissing,
			vtMissing,
			variant_t(long(1)),
			variant_t(true),
			&menu_item_obj);

		// set item caption
		CString caption;
		caption.LoadString(command_id);
		menu_item_obj->put_Caption(bstr_t(caption));

		// The target action is marker add-on, since we work using MarkerEvent.
		// For more information about that, see the article in MSDN: 
		// http://msdn.microsoft.com/en-us/library/aa140366.aspx

		CString parameter;
		parameter.Format(L"%d", command_id);
		menu_item_obj->put_Parameter(bstr_t(parameter));

		// Set unique tag, so that the command is not lost
		CString tag;
		tag.Format(L"%s_%d", ADDON_NAME, command_id);
		menu_item_obj->put_Tag(bstr_t(tag));

		m_buttons.Add(new ClickEventRedirector(menu_item_obj));
	}

	void CleanupMenuItem(Office::_CommandBarsPtr cbs, UINT command_id)
	{
		CString tag;
		tag.Format(L"%s_%d", ADDON_NAME, command_id);

		Office::CommandBarControlsPtr controls;
		if (SUCCEEDED(cbs->FindControls(vtMissing, vtMissing, variant_t(tag), vtMissing, &controls)))
		{
			int count = 0;
			controls->get_Count(&count);
			for (int i = count; i > 0; --i)
			{
				Office::CommandBarControlPtr control;
				if (SUCCEEDED(controls->get_Item(variant_t(long(i)), &control)))
				{
					control->Delete();
				}
			}
		}
	}

	void CreateCommandBarsMenu(Visio::IVApplicationPtr app)
	{
		Office::_CommandBarsPtr cbs = app->CommandBars;

		CleanupMenuItem(cbs, IDS_JumpToShape);

		Office::CommandBarPtr drawing_context_popup = GetDrawingContextMenu(cbs);

		Office::CommandBarControlsPtr controls;
		drawing_context_popup->get_Controls(&controls);

		FillMenuItem(controls, IDS_JumpToShape);
	}

	void DestroyCommandBarsMenu()
	{
		for (size_t i = 0; i < m_buttons.GetCount(); ++i)
			delete m_buttons[i];

		m_buttons.RemoveAll();
	}

	static Visio::IVWindowPtr GetExplorerWindow(Visio::IVWindowPtr window)
	{
		using namespace Visio;

		switch (window->SubType)
		{
		case visPageWin:
            return window->GetWindows()->GetItemFromID(visWinIDDrawingExplorer);
		case visMasterWin:
			return window->GetWindows()->GetItemFromID(visWinIDMasterExplorer);
		default:
			return nullptr;
		}
	}

	static HTREEITEM FindItem(HWND hwnd_tree, HTREEITEM root, LPCWSTR name)
	{
		const int ImageType_Folder = 0;
		const int ImageType_Extras = 10;

		CAtlArray<HTREEITEM> queue;
		queue.Add(root);
		size_t pos = 0;

		while (pos < queue.GetCount())
		{
			HTREEITEM first = queue[pos++];

			TreeView_Expand(hwnd_tree, first, TVE_EXPAND);

			HTREEITEM child = TreeView_GetChild(hwnd_tree, first);
			while (child)
			{
				TVITEM item;

				WCHAR item_name[1024] = L"";
				item.hItem = child;
				item.mask = TVIF_TEXT | TVIF_IMAGE | TVIF_CHILDREN;
				item.pszText = item_name;
				item.cchTextMax = 1024;

				if (TreeView_GetItem(hwnd_tree, &item))
				{
					if (!StrCmp(item_name, name) && item.iImage != ImageType_Folder)
						return child;

					if (item.cChildren > 0 && item.iImage < ImageType_Extras)
						queue.Add(child);
				}

				child = TreeView_GetNextSibling(hwnd_tree, child);
			}
		}

		return NULL;
	}

	static HRESULT JumpToShape(Visio::IVApplicationPtr app)
	{
		using namespace Visio;

		IVWindowPtr window = app->ActiveWindow;

		IVWindowPtr explorer_window = GetExplorerWindow(window);

		if (explorer_window == nullptr)
			return S_FALSE;

		explorer_window->Visible = VARIANT_TRUE;

		IVSelectionPtr selection = window->Selection;
		selection->IterationMode = 0;
		IVShapePtr shape = selection->PrimaryItem;

		if (shape != nullptr)
		{
			HWND hwnd = (HWND)(LONG_PTR)explorer_window->GetWindowHandle32();

			CAtlArray<CString> node_names;
			for (IVShapePtr currentShape = shape; currentShape != nullptr; currentShape = currentShape->Parent)
				node_names.InsertAt(0, currentShape->Name);

			// drawing explorer
			IVPagePtr containing_page = shape->ContainingPage;
			if (containing_page != nullptr)
			{
				CString page_name = containing_page->Name;
				node_names.InsertAt(0, page_name);
			}

			// Master explorer
			IVMasterPtr containing_master = shape->ContainingMaster;
			if (containing_master != nullptr)
			{
				// The master explorer window shows original name
				IVMasterPtr original_master = containing_master->Original;
				CString master_name = original_master != nullptr
					? original_master->Name 
					: containing_master->Name;

				node_names.InsertAt(0, master_name);
			}

			HWND hwnd_tree = GetWindow(hwnd, GW_CHILD);

			SendMessage(hwnd_tree, WM_SETREDRAW, 0, 0);

			HTREEITEM root = TreeView_GetRoot(hwnd_tree);
			for (size_t i = 0; i < node_names.GetCount(); ++i)
				root = FindItem(hwnd_tree, root, node_names[i]);

			if (root)
			{
				TreeView_Select(hwnd_tree, root, TVGN_CARET);
				TreeView_EnsureVisible(hwnd_tree, root);
				SetFocus(hwnd_tree);
			}

			SendMessage(hwnd_tree, WM_SETREDRAW, 1, 0);
		}

		return S_OK;
	}

	void Create(IDispatch * pApplication, IDispatch * pAddInInst) 
	{
		pApplication->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_app);
		pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_addin);

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
