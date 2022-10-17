#include "task02.h"
#include "pch.h"
#include "globals.h"

static LPCTSTR taskProps[] = {
	_T("Manufacturer"),
	_T("Name"),
	_T("Version"),
	NULL
};

HRESULT Task02(VOID)
{
	HRESULT hr = S_OK;
	IWbemClassObject *pObj = NULL;
	IEnumWbemClassObject *pEnum = NULL;
	VARIANT v;
	VariantInit(&v);

	hr = pSvc->ExecQuery(
		(BSTR)_T("WQL"),
		(BSTR)_T(
			"SELECT * "
			"FROM Win32_BIOS"
		),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL, &pEnum
	);
	if (FAILED(hr))
		goto fail;


	while (1) {
		ULONG uRet = 0;
		hr = pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);
		if (uRet == 0)
			break;
		for (LPCTSTR *prop = taskProps; *prop; prop++) {
			HRESULT get_res = pObj->Get(
				*prop, 0,
				&v, 0, 0
			);
			if (FAILED(get_res))
				break;
			std::tcout << *prop << _T(": ") << &v << _T("\n");
		}

	}

	fail:
	pEnum->Release();
	pObj->Release();
	VariantClear(&v);
	return hr;
}