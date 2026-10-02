#include "pch.h"
#include "Connect.h"
#include "HttpServer.h"

#include <new>
#include <utility>

namespace
{
	HRESULT ReadServerPort(unsigned short &port)
	{
		port = 50051;
		DWORD value = 0;
		DWORD size = sizeof(value);
		LONG result = RegGetValueW(HKEY_CURRENT_USER,
			L"Software\\Microsoft\\Office\\PowerPoint\\Addins\\NetOffice.Automate",
			L"ServerPort", RRF_RT_REG_DWORD, nullptr, &value, &size);
		if (result == ERROR_FILE_NOT_FOUND)
			return S_OK;
		if (result != ERROR_SUCCESS)
			return HRESULT_FROM_WIN32(result);
		if (value == 0 || value > 65535)
			return E_INVALIDARG;
		port = static_cast<unsigned short>(value);
		return S_OK;
	}
}

STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode connectMode, IDispatch *pAddInInst, SAFEARRAY **custom)
{
	Disconnect();
	try
	{
		unsigned short port;
		HRESULT hr = ReadServerPort(port);
		if (FAILED(hr))
			return hr;
		auto server = std::make_shared<HttpServer>();
		hr = server->Start(pApplication, port);
		if (FAILED(hr))
			return hr;
		m_pHttpServer = std::move(server);
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
	auto server = std::move(m_pHttpServer);
	if (server)
		server->Stop();
}
