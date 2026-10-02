#include "pch.h"
#include "Connect.h"
#include "GrpcServer.h"

#include <new>
#include <utility>

STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode connectMode, IDispatch *pAddInInst, SAFEARRAY **custom)
{
	Disconnect();
	try
	{
		auto server = std::make_shared<GrpcServer>();
		HRESULT hr = server->Start(pApplication);
		if (FAILED(hr))
			return hr;
		m_pGrpcServer = std::move(server);
		return S_OK;
	}
	catch (const std::bad_alloc &)
	{
		return E_OUTOFMEMORY;
	}
}

STDMETHODIMP CConnect::OnDisconnection(AddInDesignerObjects::ext_DisconnectMode removeMode, SAFEARRAY **custom)
{
	Disconnect();
	return S_OK;
}

STDMETHODIMP CConnect::OnAddInsUpdate(SAFEARRAY **custom)
{
	return S_OK;
}

STDMETHODIMP CConnect::OnStartupComplete(SAFEARRAY **custom)
{
	return S_OK;
}

STDMETHODIMP CConnect::OnBeginShutdown(SAFEARRAY **custom)
{
	Disconnect();
	return S_OK;
}

void CConnect::FinalRelease()
{
	Disconnect();
}

void CConnect::Disconnect()
{
	// Clear our ownership first: Office can reenter callbacks while disconnecting.
	auto server = std::move(m_pGrpcServer);
	if (server)
		server->Stop();
}
