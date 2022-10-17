#include "task04.h"
#include "pch.h"
#include "globals.h"

static LPCTSTR taskProps[] = {
	_T("Name"),
	_T("ProcessId"),
	_T("Priority"),
	NULL
};

HRESULT Task04(VOID)
{
	HRESULT hr = S_OK;
	IWbemClassObject *pObj = NULL;
	IEnumWbemClassObject *pEnum = NULL;
	ULONG uRet = 0;
	ULONG uPriority = UINT_MAX;
	CIMTYPE cimtype;
	VARIANT v;
	VariantInit(&v);

	std::tcout << _T("-- ") << _T(__FUNCTION__) << _T("\n");
	
	hr = pSvc->ExecQuery(
		(BSTR)_T("WQL"),
		(BSTR)_T("SELECT * ")
			_T("FROM Win32_Process"),
		WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnum
	);

	while (1) {
		hr = pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);
		if (uRet == 0)
			break;
		hr = pObj->Get(
			_T("Priority"), 0,
			&v, NULL, NULL
		);
		if (V_UI4(&v) < uPriority) uPriority = V_UI4(&v);
	}
	
	/*
	 * Enumerate the second time and select only processes
	 * with given priority = uPriority
	 */
	hr = pEnum->Reset();
	while (1) {
		hr = pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);
		if (uRet == 0)
			break;
		hr = pObj->Get(
			_T("Priority"), 0,
			&v, NULL, NULL
		);
		
		if (V_UI4(&v) == uPriority) {
			for (LPCTSTR *prop = taskProps; prop; prop++) {
				HRESULT get_res = pObj->Get(
					*prop, 0,
					&v, &cimtype, NULL
				);
				if (FAILED(get_res))
					break;
				std::tcout << _T("\t") << *prop << _T(": ");
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
	}

	fail:
	VariantClear(&v);
	pObj->Release();
	pEnum->Release();
	return hr;
}