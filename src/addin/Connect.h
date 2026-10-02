#pragma once

#include "pch.h"
#include <memory>

class HttpServer;

using namespace ATL;

class DECLSPEC_UUID("6d274715-3f05-4505-aa63-2ae9df1e6881") NetOfficeAutomateLibraryGuid;

class ATL_NO_VTABLE CConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &__uuidof(NetOfficeAutomateLibraryGuid)>,
	public IDispatchImpl<AddInDesignerObjects::_IDTExtensibility2,
		&AddInDesignerObjects::IID__IDTExtensibility2,
		&AddInDesignerObjects::LIBID_AddInDesignerObjects, 1, 0>
{
public:
	DECLARE_NO_REGISTRY()
	DECLARE_NOT_AGGREGATABLE(CConnect)

	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(AddInDesignerObjects::_IDTExtensibility2)
	END_COM_MAP()

	STDMETHOD(OnConnection)(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *pAddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY **custom);
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom);
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom);
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom);
	void FinalRelease();

private:
	void Disconnect();
	std::shared_ptr<HttpServer> m_pHttpServer;

};

OBJECT_ENTRY_AUTO(__uuidof(NetOfficeAutomateLibraryGuid), CConnect)
