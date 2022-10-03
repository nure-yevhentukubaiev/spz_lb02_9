#define _WIN32_DCOM
#include "globals.h"
#include "pch.h"

HRESULT CoInit(VOID);

VOID WMIClose(VOID);

HRESULT WMIConnect(VOID);

HRESULT CoInit(VOID)
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
	}
	return hr;
}

HRESULT WMIConnect(VOID)
{
	
	HRESULT hr = CoCreateInstance(
		CLSID_WbemLocator,
		0,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator,
		(LPVOID *)&pLoc
	);
	if (FAILED(hr)) {
		_tprintf_s(_T("CoCreateInstace failed, hr: %X\n"), hr);
		return hr;
	}
	
	hr = pLoc->ConnectServer(
		(BSTR)_T("ROOT\\CimV2"),
		NULL,
		NULL,
		0,
		NULL,
		0,
		0,
		&pSvc
	);
	if (FAILED(hr)) {
		_tprintf_s(_T("ConnectServer failed, hr: %X\n"), hr);
		pLoc->Release();
	}



	return hr;
}

VOID WMIClose(VOID)
{
	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	pclsObj->Release();
	return;
}

int _tmain(int argc, TCHAR **argv)
{
	HRESULT hr = CoInit() || WMIConnect();
	if (SUCCEEDED(hr)) {
	}
	WMIClose();
	CoUninitialize();
	return (int)hr;
}