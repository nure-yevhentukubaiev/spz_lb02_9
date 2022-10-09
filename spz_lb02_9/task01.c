#include "task01.h"
#include "pch.h"
#include "globals.h"

static LPCTSTR taskProps[] = {
	_T("Manufacturer"),
	_T("Name"),
	_T("Version")
};

HRESULT Task01(VOID)
{
	HRESULT hr = S_OK;

	hr = pSvc->ExecQuery(
		(BSTR)_T("WQL"),
		(BSTR)_T(
			"SELECT * "
			"FROM Win32_BIOS"
		),
		WBEM_FLAG_FORWARD_ONLY
		| WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator
	);

	
	while (pEnumerator) {
		ULONG uRet = 0;
		hr = pEnumerator->Next(WBEM_INFINITE, 1, &pClsObj, &uRet);
		if (uRet == 0)
			break;
		VARIANT vt;
		VariantInit(&vt);
		for (LPCTSTR *prop = taskProps; prop; prop++) {
			HRESULT get_res = pClsObj->Get(
				*prop,
				0,
				&vt,
				0,
				0
			);
			if (FAILED(get_res)) break;
			_tprintf_s(_T("%s: %s\n"), *prop, vt.bstrVal);
		}
		VariantClear(&vt);
	}

	return hr;
}