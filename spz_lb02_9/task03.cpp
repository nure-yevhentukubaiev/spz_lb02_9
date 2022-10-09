#include "task03.h"
#include "pch.h"
#include "globals.h"

static LPCTSTR lpszWin32ProcessProps[] = {
	_T("ExecutablePath")        // a) шлях до виконуваного файлу процесу;
	_T("CreationDate"),         // b) час початку процесу
	_T("Priority"),             // c) пріоритет процесу;
	_T("ID"),                   // d) ідентифікатор процесу;
	_T("ThreadCount")           // e) кількість активних потоків процесу;
};



static LPCTSTR lpszWin32ThreadProps[] = {
	_T("ID"),                   // -ідентифікатор процесу, що створив потік;
	_T("Priority"),             // -динамічний пріоритет потоку;
	_T("PriorityBase"),         // -базовий пріоритет потоку;
	_T("ElapsedTime"),          // -загальний час виконання потоку;
	_T("ExecutionState")        // -стан потоку.
};

/* TODO: how to launch MS Access. Like that? */
static CONST BSTR lpszProgCmdLine = SysAllocString(_T("C:\\Program Files(x86)\\Microsoft Office\\OFFICE11\\MSACCESS.EXE"));

static struct WMICreateProcessStruct {
	CONST BSTR lpszArg;
	CONST BSTR lpszVal;
	CIMTYPE cimType;
} WMICreateProcess[] = {
	{(BSTR)_T("CommandLine"),               CIM_STRING},
	{(BSTR)_T("ProcessStartupInformation"), CIM_OBJECT}
};

/* TODO: Is this a WMIDateToStringDate function? */
/*
 Function WMIDateStringToDate(dtmInstallDate)
 WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
 Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
 & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
 Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
 13, 2))
End Function
 */

static HRESULT Task03_CreateProcess(IWbemClassObject *pClsProcessObj)
{
	HRESULT hr = S_OK;
	IWbemClassObject *pClsProcessInParam = NULL;
	IWbemClassObject *pClsProcessOutParam = NULL;
	IWbemCallResult *pRet = NULL;

	hr = pClsProcessObj->GetMethod(
		_T("Create"), 0,
		&pClsProcessInParam, NULL
	);

	VARIANT v;
	VariantInit(&v);
	V_VT(&v) = VT_BSTR;
	//V_BSTR(&v) = lpszProgPath;
	for (WMICreateProcessStruct *i = WMICreateProcess; i; i++) {
		V_BSTR(&v) = i->lpszVal;
		hr = pClsProcessInParam->Put(
			i->lpszArg,
			0,
			&v,
			i->cimType
		);
	}

	hr = pSvc->ExecMethod(
		(BSTR)_T("Win32_Process"),
		(BSTR)_T("Create"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		NULL,
		&pClsProcessObj,
		&pRet
	);

	if (FAILED(hr) || pRet == 0) {
		_tprintf_s(_T("ExecMethod failed/no process found, hr: %lX\n"), hr);
		goto fail;
	}
	while (TRUE) {
		long lStatus = 0;
		hr = pRet->GetCallStatus(5000, &lStatus);
		if (hr == WBEM_S_NO_ERROR || hr != WBEM_S_TIMEDOUT)
			break;
	}

	hr = pRet->GetResultObject(5000, &pClsProcessObj);
}

HRESULT Task03(VOID)
{
	HRESULT hr = S_OK;
	BSTR lpszWQLQuery = NULL;
	IWbemClassObject *pClsProcessObj = 0;
	
	if (FAILED(hr))
		goto fail;

	lpszWQLQuery = (BSTR)LocalAlloc(LMEM_ZEROINIT, lstrlen(lpszProgName) + 256);
	_stprintf_s(
		lpszWQLQuery, 1,
		_T("SELECT * "
		"FROM Win32_Process"
		"WHERE Name=%s"),
		lpszProgName
	);
	hr = pSvc->ExecQuery(
		(BSTR)_T("WQL"),
		lpszWQLQuery,
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator
	);
	if (FAILED(hr))
		goto fail;

	while (pEnumerator) {
		ULONG uRet = 0;
		hr = pEnumerator->Next(WBEM_INFINITE, 1, &pClsObj, &uRet);
		if (uRet == 0)
			break;
		VARIANT vt;
		VariantInit(&vt);
		for (LPCTSTR *prop = lpszWin32ProcessProps; prop; prop++) {
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
	
fail:
	pClsProcessObj->Release();
	pRet->Release();
	LocalFree(lpszWQLQuery);
	return hr;
}