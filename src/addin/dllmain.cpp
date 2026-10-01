// AddIn.cpp : Implementation of library exports.

#include "pch.h"

class CAddinModule : public ATL::CAtlDllModuleT<CAddinModule>
{
};

CAddinModule _module;

// Library Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	return _module.DllMain(dwReason, lpReserved);
}

// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow()
{
	return _module.DllCanUnloadNow();
}

// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	return _module.DllGetClassObject(rclsid, riid, ppv);
}
