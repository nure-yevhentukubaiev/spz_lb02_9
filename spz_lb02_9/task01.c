#include "task01.h"
#include "pch.h"
#include "globals.h"

HRESULT Task01(VOID)
{
	HRESULT hr = S_OK;
	
	_tprintf_s(_T("-- %s\n"), _T(__FUNCTION__));

	hr = Task01_GetPropertyInfo() || Task01_GetIndividualProperties();
	return hr;
}

static HRESULT Task01_GetPropertyInfo(VOID)
{
	HRESULT hr = S_OK;
	BSTR lpszProps = NULL;

	hr = pSvc->GetObject(
		(BSTR)_T("Win32_BIOS"), 0,
		NULL, &pClsObj, NULL
	);
	if (FAILED(hr))
		goto fail;

	hr = pClsObj->GetNames(
		NULL, WBEM_FLAG_ALWAYS,
		NULL, NULL
	);
	if (FAILED(hr))
		goto fail;

	hr = pClsObj->GetObjectText(0, &lpszProps);
	if (FAILED(hr))
		goto fail;

	_tprintf_s(_T("%s\n"), lpszProps);
	
	fail:
	SysFreeString(lpszProps);
	return hr;
}

static HRESULT Task01_GetIndividualProperties(VOID)
{
	HRESULT hr = S_OK;
	static LPCTSTR taskProps[] = {
		_T("Manufacturer"),
		_T("Name"),
		_T("Version")
	};
	VARIANT vt;
	VariantInit(&vt);

	hr = pSvc->ExecQuery(
		(BSTR)_T("WQL"),
		(BSTR)_T(
			"SELECT * "
			"FROM Win32_BIOS"
		),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL, &pEnumerator
	);
	if (FAILED(hr))
		goto fail;

	while (pEnumerator) {
		ULONG uRet = 0;
		pEnumerator->Next(WBEM_INFINITE, 1, &pClsObj, &uRet);
		if (uRet == 0)
			break;
		for (LPCTSTR *prop = taskProps; prop; prop++) {
			HRESULT get_res = pClsObj->Get(
				*prop, 0,
				&vt, 0, 0
			);
			if (FAILED(get_res))
				continue;
			_tprintf_s(_T("%s: %s\n"), *prop, vt.bstrVal);
		}
	}

	fail:
	VariantClear(&vt);
	return hr;
}