// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols

using namespace Office;
using namespace AddInDesignerObjects;

// CConnect
class ATL_NO_VTABLE CVisioConnect : 
	public CComObjectRootEx<CComSingleThreadModel>
	, public CComCoClass<CVisioConnect, &CLSID_Connect>
	, public IDispatchImpl<ICallbackInterface, &__uuidof(ICallbackInterface), &LIBID_AddinLib, 1, 0>
	, public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), 1, 0>
	, public IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), 12, 0>
{
public:
	CVisioConnect();

	~CVisioConnect();

DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)
DECLARE_NOT_AGGREGATABLE(CVisioConnect)

BEGIN_COM_MAP(CVisioConnect)
	COM_INTERFACE_ENTRY2(IDispatch, ICallbackInterface)
	COM_INTERFACE_ENTRY(_IDTExtensibility2)
	COM_INTERFACE_ENTRY(IRibbonExtensibility)
	COM_INTERFACE_ENTRY(ICallbackInterface)
END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

public:
	//IDTExtensibility2 implementation:
	STDMETHOD(OnConnection)(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);

	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY **custom );
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom );
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom );
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom );

	//IRibbonExtensibility implementation:
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonString);
	STDMETHOD(GetRibbonLabel)(IDispatch* pControl, BSTR *pbstrLabel);
	STDMETHOD(OnRibbonButtonClicked)(IDispatch * pControl);
	STDMETHOD(IsRibbonButtonVisible)(IDispatch * pControl, VARIANT_BOOL* pResult);
	STDMETHOD(IsRibbonButtonEnabled)(IDispatch * pControl, VARIANT_BOOL* pResult);

	STDMETHOD(OnRibbonLoad)(IDispatch* disp);

	struct Impl;
	Impl* m_impl;
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CVisioConnect)
