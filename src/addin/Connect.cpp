#include "pch.h"
#include "Connect.h"

STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode connectMode, IDispatch *pAddInInst, SAFEARRAY **custom)
{
	pApplication->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pApplication);
	pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pAddInInstance);

	return S_OK;
}

STDMETHODIMP CConnect::OnDisconnection(AddInDesignerObjects::ext_DisconnectMode removeMode, SAFEARRAY **custom)
{
	m_pApplication = NULL;
	m_pAddInInstance = NULL;
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
	return S_OK;
}
