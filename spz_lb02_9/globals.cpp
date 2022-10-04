#include "globals.h"

#include "pch.h"

IWbemServices *pSvc = NULL;
IWbemLocator *pLoc = NULL;
IEnumWbemClassObject *pEnumerator = NULL;
IWbemClassObject *pclsObj = NULL;