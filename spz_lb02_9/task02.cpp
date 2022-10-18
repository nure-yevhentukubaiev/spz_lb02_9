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
	CIMTYPE cimtype = 0;
	VARIANT v;
	VariantInit(&v);

	std::tcout << _T("-- ") << _T(__FUNCTION__) << _T("\n");

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
				&v, &cimtype, 0
			);
			if (FAILED(get_res))
				break;
			std::tcout << *prop << _T(": ");
			switch (cimtype) {
			case CIM_STRING:
				std::tcout << V_BSTR(&v);
				break;
			case CIM_SINT16:
			case CIM_SINT32:
			case CIM_SINT64:
				std::tcout << V_I8(&v);
				break;
			case CIM_UINT16:
			case CIM_UINT32:
			case CIM_UINT64:
			default:
				std::tcout << V_UI8(&v);
				break;
			}
			std::tcout << _T("\n");
		}

	}

	fail:
	pEnum->Release();
	pObj->Release();
	VariantClear(&v);
	return hr;
}