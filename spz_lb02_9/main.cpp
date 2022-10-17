#include "globals.h"
#include "pch.h"
#include "task01.h"
#include "task02.h"
#include "task03.h"
#include "task04.h"
#include "task05.h"

HRESULT CoInit(VOID);

VOID WMIClose(VOID);

HRESULT WMIConnect(VOID);

HRESULT CoInit(VOID)
{
	HRESULT hr = S_OK;
	
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr)) {
		std::tcout << _T("CoInitializeEx failed, hr: ")
			<< std::hex << hr << "\n";
		goto fail;
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
		std::tcout << _T("CoInitializeSecurity failed, hr: ")
			<< std::hex << hr << "\n";
		goto fail;
	}

	fail:
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
		std::tcout << _T("CoCreateInstace failed, hr: ")
			<< std::hex << hr << "\n";
		goto fail;
	}
	
	hr = pLoc->ConnectServer(
		(BSTR)_T("ROOT\\CimV2"),
		NULL, NULL, NULL,
		0, 0,
		NULL, &pSvc
	);
	if (FAILED(hr)) {
		std::tcout << _T("ConnectServer failed, hr: ")
			<< std::hex << hr << "\n";
		goto fail;
	}

	// Set the proxy so that impersonation of the client occurs.
	hr = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx 
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx 
		NULL,                        // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		NULL,                        // client identity
		EOAC_NONE                    // proxy capabilities
	);

	if (FAILED(hr)) {
		std::tcout << _T("CoSetProxyBlanket failed, hr: ")
			<< std::hex << hr << "\n";
		goto fail;
	}

	fail:
	return hr;
}

VOID WMIClose(VOID)
{
	pSvc->Release();
	pLoc->Release();
	return;
}

int _tmain(int argc, TCHAR **argv)
{
	UNREFERENCED_PARAMETER(argc);
	UNREFERENCED_PARAMETER(argv);

	HRESULT hr = CoInit();
	hr = WMIConnect();
	if (SUCCEEDED(hr)) {
		Task01();
		Task02();
		Task03();
		Task04();
		Task05();
	}
	WMIClose();
	CoUninitialize();
	return (int)hr;
}