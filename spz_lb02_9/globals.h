#pragma once
#include "pch.h"

/* WMI Services instance */
extern IWbemServices *pSvc;
/* WMI Locator instance */
extern IWbemLocator *pLoc;

extern IEnumWbemClassObject *pEnumerator;

extern IWbemClassObject *pClsObj;