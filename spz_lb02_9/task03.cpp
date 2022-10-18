#include "task03.h"
#include "pch.h"
#include "globals.h"

static LPCTSTR lpszWin32ProcessProps[] = {
	_T("Name"),                 
	_T("ExecutablePath"),       // a) шлях до виконуваного файлу процесу;
	_T("CreationDate"),         // b) час початку процесу;
	_T("Priority"),             // c) пріоритет процесу;
	_T("ProcessId"),            // d) ідентифікатор процесу;
	_T("ThreadCount"),          // e) кількість активних потоків процесу;
	NULL
};

static LPCTSTR lpszWin32ThreadProps[] = {
	_T("ProcessHandle"),        // -ідентифікатор процесу, що створив потік;
	_T("Priority"),             // -динамічний пріоритет потоку;
	_T("PriorityBase"),         // -базовий пріоритет потоку;
	_T("ElapsedTime"),          // -загальний час виконання потоку;
	_T("ExecutionState"),       // -стан потоку.
	NULL
};

static CONST BSTR lpszProgCmdLine = SysAllocString(
	_T("\"C:\\Program Files (x86)\\Microsoft Office\\OFFICE11\\MSACCESS.EXE\"")
);

static BSTR bszProcessID = NULL;

/*
 * Converts datetime in WMI format to human-readable datetime string
 * Example:
 * 20111201120159.000000+000 -> 01/12/2011 12:01:59
 * YYYYMMDDhhmmss.uuuuuu+ggg -> DD/MM/YYYY hh:mm:ss
 * See: https://learn.microsoft.com/en-us/previous-versions/tn-archive/ee156576(v=technet.10)?redirectedfrom=MSDN
 */
static VOID WMIDateStringToDate(BSTR d)
{
	/* Date string copy */
	CONST BSTR b = SysAllocString((CONST OLECHAR *)d);
	_stprintf_s(
		d, 35,
		_T("%c%c/%c%c/%c%c%c%c %c%c:%c%c:%c%c\0"),
		b[6], b[7], b[4], b[5], b[0], b[1], b[2], b[3], /* date */
		b[8], b[9], b[10], b[11], b[12], b[13]          /* time */
	);
	SysFreeString(b);
	return;
}

static HRESULT Task03_SetProcessStartup(IWbemClassObject *pClsProcessStartupParam)
{
	HRESULT hr = S_OK;
	VARIANT v;
	VariantInit(&v);
	
	VarBstrFromUint(SW_MAXIMIZE, 0, 0, &V_BSTR(&v));
	hr = pClsProcessStartupParam->Put(_T("ShowWindow"), 0, &v, 0);
	
	VarBstrFromUint(NORMAL_PRIORITY_CLASS, 0, 0, &V_BSTR(&v));
	hr = pClsProcessStartupParam->Put(_T("PriorityClass"), 0, &v, 0);
	
	VariantClear(&v);
	return hr;
}

static HRESULT Task03_SetProcessInParam(IWbemClassObject *pClsProcessInParam)
{
	HRESULT hr = S_OK;
	IWbemClassObject *pClsProcessStartupParam = NULL;
	VARIANT v;
	VariantInit(&v);
	
	hr = pSvc->GetObject(
		(BSTR)_T("Win32_ProcessStartup"), 0,
		NULL, &pClsProcessStartupParam, NULL
	);
	if (FAILED(hr))
		goto fail;

	hr = Task03_SetProcessStartup(pClsProcessStartupParam);
	
	V_VT(&v) = VT_BSTR;
	V_BSTR(&v) = lpszProgCmdLine;
	hr = pClsProcessInParam->Put(_T("CommandLine"), 0, &v, CIM_STRING);

	V_VT(&v) = VT_UNKNOWN;
	V_UNKNOWN(&v) = pClsProcessStartupParam;
	hr = pClsProcessInParam->Put(_T("ProcessStartupInformation"), 0, &v, CIM_OBJECT);

	goto fail;
	fail:
	VariantClear(&v);
	return S_OK;
}

static HRESULT Task03_CreateProcess(VOID)
{
	HRESULT hr = S_OK;
	IWbemClassObject *pClsProcessInParam = NULL;
	IWbemClassObject *pClsProcessOutParam = NULL;
	IWbemClassObject *pClsProcessDef = NULL;
	IWbemClassObject *pClsProcessInParamInst = NULL;
	static LPCTSTR lpszMethod = _T("Create");
	static LPCTSTR lpszClass = _T("Win32_Process");
	VARIANT v;
	VariantInit(&v);

	hr = pSvc->GetObject(
		(BSTR)lpszClass, 0,
		NULL, &pClsProcessDef, NULL
	);

	hr = pClsProcessDef->GetMethod(
		lpszMethod, 0,
		&pClsProcessInParam, NULL
	);
	
	hr = pClsProcessInParam->SpawnInstance(0, &pClsProcessInParamInst);

	hr = Task03_SetProcessInParam(pClsProcessInParamInst);

	hr = pSvc->ExecMethod(
		(BSTR)lpszClass, (BSTR)lpszMethod, 0,
		NULL, pClsProcessInParamInst, &pClsProcessOutParam, NULL
	);
	if (FAILED(hr)) {
		_tprintf_s(_T("ExecMethod failed, hr: %lX\n"), hr);
		goto fail;
	}
	
	/* 
	 * Acquire the created process ID for further querying
	 */
	pClsProcessOutParam->Get(
		_T("ProcessId"), 0,
		&v, 0, 0
	);
	VarBstrFromUint(V_UI4(&v), 0, 0, &bszProcessID);

	fail:
	VariantClear(&v);
	pClsProcessInParamInst->Release();
	pClsProcessInParam->Release();
	pClsProcessOutParam->Release();
	pClsProcessDef->Release();
	return hr;
}

static HRESULT Task03_GetThreadPropsOfProcess(VOID)
{
	HRESULT hr = S_OK;
	IEnumWbemClassObject *pEnum = NULL;
	IWbemClassObject *pObj = NULL;
	ULONG uRet = 0;
	CIMTYPE cimtype;
	VARIANT v;
	BSTR bszQuery = SysAllocString(
		_T("SELECT * ")
		_T("FROM Win32_Thread ")
		_T("WHERE ProcessHandle=")
	);
	VariantInit(&v);

	VarBstrCat(
		bszQuery,
		bszProcessID,
		&bszQuery
	);

	hr = pSvc->ExecQuery(
		(BSTR)_T("WQL"),
		bszQuery,
		WBEM_FLAG_FORWARD_ONLY,
		NULL, &pEnum
	);
	if (FAILED(hr))
		goto fail;

	while (pEnum) {
		pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);
		if (uRet == 0)
			break;

		std::tcout << _T("-");
		for (LPCTSTR *prop = lpszWin32ThreadProps; *prop; prop++) {
			HRESULT get_res = pObj->Get(
				*prop, 0,
				&v, &cimtype, 0
			);
			if (FAILED(get_res))
				continue;

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

	fail:
	VariantClear(&v);
	pObj->Release();
	pEnum->Release();
	SysFreeString(bszQuery);
	return hr;
}

static HRESULT Task03_GetProcessProps(VOID)
{
	HRESULT hr = S_OK;
	IWbemClassObject *pObj = NULL;
	IEnumWbemClassObject *pEnum = NULL;
	ULONG uRet = 0;
	CIMTYPE cimtype;
	BSTR bszQuery = SysAllocString(
		_T("SELECT * ")
		_T("FROM Win32_Process ")
		_T("WHERE ProcessId=")
	);
	VARIANT v;
	VariantInit(&v);

	VarBstrCat(
		bszQuery,
		bszProcessID,
		&bszQuery
	);

	hr = pSvc->ExecQuery(
		(BSTR)_T("WQL"),
		bszQuery,
		WBEM_FLAG_FORWARD_ONLY,
		NULL, &pEnum
	);
	if (FAILED(hr))
		goto fail;

	while (pEnum) {
		pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);
		if (uRet == 0)
			break;
		for (LPCTSTR *prop = lpszWin32ProcessProps; *prop; prop++) {
			HRESULT get_res = pObj->Get(
				*prop, 0,
				&v, &cimtype, 0
			);
			if (FAILED(get_res))
				continue;
			std::tcout << *prop << _T(": ");
			switch (cimtype) {
			case CIM_DATETIME:
				WMIDateStringToDate(V_BSTR(&v));
				/* nobreak; */
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
		std::tcout << _T("Threads:") << _T("\n");
		pObj->Get(
			_T("ProcessId"), 0,
			&v, 0, 0
		);
		Task03_GetThreadPropsOfProcess();
	}

	fail:
	VariantClear(&v);
	pEnum->Release();
	pObj->Release();
	return hr;
}

HRESULT Task03(VOID)
{
	HRESULT hr = S_OK;
	
	std::tcout << _T("-- ") << _T(__FUNCTION__) << _T("\n");

	hr = Task03_CreateProcess();
	if (FAILED(hr))
		goto fail;

	hr = Task03_GetProcessProps();
	if (FAILED(hr))
		goto fail;

	fail:
	return hr;
}