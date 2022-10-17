#include "task01.h"
#include "pch.h"
#include "globals.h"

static CONST BSTR bszClass = SysAllocString(_T("Win32_BIOS"));

HRESULT Task01(VOID)
{
	HRESULT hr = S_OK;
	IWbemClassObject *pObj = NULL;
	BSTR bszAllProps = NULL;

	std::tcout << _T("-- ") << _T(__FUNCTION__) << _T("\n");

	hr = pSvc->GetObject(
		bszClass, 0,
		NULL, &pObj, NULL
	);
	if (FAILED(hr))
		goto fail;

	hr = pObj->GetObjectText(0, &bszAllProps);
	if (FAILED(hr))
		goto fail;

	std::tcout << bszAllProps;
	
	fail:
	SysFreeString(bszAllProps);
	pObj->Release();
	return hr;
}