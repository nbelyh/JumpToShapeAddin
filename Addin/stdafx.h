// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#ifndef _WIN32_WINNT		// Allow use of features specific to Windows NT 4 or later.
#define _WIN32_WINNT 0x0501	// Change this to the appropriate value to target Windows 2000 or later.
#endif						

#ifndef _WIN32_IE			// Allow use of features specific to IE 4.0 or later.
#define _WIN32_IE 0x0700	// Change this to the appropriate value to target IE 7.0 or later.
#endif

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

#define _ATL_ALL_WARNINGS	// turns off ATL's hiding of some common and often safely ignored warning messages

#include "resource.h"

#include <atlbase.h>
#include <atlcom.h>
#include <atlstr.h>
#include <CommCtrl.h>

#pragma warning( disable : 4278 )
#pragma warning( disable : 4146 )
                                              
#include "import\MSADDNDR.tlh"
#include "import\MSO.tlh"
#include "import\VISLIB.tlh"

#pragma warning( default : 4146 )
#pragma warning( default : 4278 )

using namespace ATL;
extern CComModule _Module;
