#include "task05.h"
#include "pch.h"
#include "globals.h"

static HRESULT Task05_01(VOID)
{
	HRESULT hr = S_OK;
	VARIANT v;
	VariantInit(&v);

	std::tcout << _T("-- ") << _T(__FUNCTION__) << _T("\n");

	goto fail;
	fail:
	VariantClear(&v);
	return hr;
}

HRESULT Task05(VOID)
{
	HRESULT hr = S_OK;
	hr = Task05_01();
	return hr;
}