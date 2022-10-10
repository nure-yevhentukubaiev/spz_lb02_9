#include "task04.h"
#include "pch.h"
#include "globals.h"

static LPCTSTR taskProps[] = {
	_T("Name"),
	_T("ProcessId"),
	_T("Priority")
};

HRESULT Task04(VOID)
{
	HRESULT hr = S_OK;
	VARIANT v;
	VariantInit(&v);

	_tprintf_s(_T("-- %s\n"), _T(__FUNCTION__));
	
	hr = pSvc->ExecQuery(
		(BSTR)_T("WQL"),
		(BSTR)_T(
			"SELECT * "
			"FROM Win32_Process "
		),
		WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator
	);

	ULONG uRet = 0;
	ULONG uPriority = UINT_MAX;
	ULONG i = 0;
	while (1) {
		hr = pEnumerator->Next(WBEM_INFINITE, 1, &pClsObj, &uRet);
		if (uRet == 0)
			break;
		hr = pClsObj->Get(
			_T("Priority"), 0,
			&v, NULL, NULL
		);
		if (v.ulVal < uPriority) uPriority = v.ulVal;
	}
	hr = pEnumerator->Reset();

	while (1) {
		hr = pEnumerator->Next(WBEM_INFINITE, 1, &pClsObj, &uRet);
		if (uRet == 0)
			break;
		hr = pClsObj->Get(
			_T("Priority"), 0,
			&v, NULL, NULL
		);
		
		if (v.ulVal == uPriority) {
			for (LPCTSTR *prop = taskProps; prop; prop++) {
				HRESULT get_res = pClsObj->Get(
					*prop, 0,
					&v, NULL, NULL
				);
				if (FAILED(get_res)) break;
				switch (V_VT(&v)) {
				case VT_BSTR:
					_tprintf_s(_T("%s: %s\n"), *prop, v.bstrVal);
					break;
				default:
					_tprintf_s(_T("%s: %u\n"), *prop, v.ulVal);
					break;
				}
				
			}
			_tprintf_s(_T("\n"));
		}
	}
	VariantClear(&v);

	return hr;
}