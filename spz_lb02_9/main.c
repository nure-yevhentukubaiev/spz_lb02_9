#define _WIN32_DCOM
#include "pch.h"

HRESULT CoWMIInit(VOID);

VOID CoWMIUninit(VOID);

HRESULT CoWMIInit(VOID)
{
	HRESULT hr = S_OK;
	
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr)) {
		_tprintf_s(_T("CoInitializeEx failed, hr: %X\n"), hr);
		return hr;
	}

	hr = CoInitializeSecurity(
		NULL,                        // Security descriptor    
		-1,                          // COM negotiates authentication service
		NULL,                        // Authentication services
		NULL,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication level for proxies
		RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation level for proxies
		NULL,                        // Authentication info
		EOAC_NONE,                   // Additional capabilities of the client/server
		NULL                         // Reserved
	);
	if (FAILED(hr)) {
		_tprintf_s(_T("CoInitializeSecurity failed, hr: %X\n"), hr);
		return hr;
	}

	return S_OK;
}

VOID CoWMIUninit(VOID)
{
	CoUninitialize();
	return;
}

int _tmain(int argc, TCHAR **argv)
{
	HRESULT hr = CoWMIInit();
	if (SUCCEEDED(hr)) {
	}
	CoWMIUninit();
	return (int)hr;
}