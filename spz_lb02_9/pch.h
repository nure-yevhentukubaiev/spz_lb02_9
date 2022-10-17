#pragma once

#define _WIN32_DCOM

#include <iostream>
#include <sstream>
#include <tchar.h>
#include <windows.h>
#include <wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")

#ifdef UNICODE
#define tcout wcout
#else
#define tcout cout
#endif

#ifdef UNICODE
#define tcin wcin
#else
#define tcin cin
#endif

#ifdef UNICODE
#define tcerr wcerr
#else
#define tcerr cerr
#endif