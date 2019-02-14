/******************************************************************************
* IpConfig plugin v0.1														  *
*																			  *
*																			  *
* Started on 06-18-2009														  *
******************************************************************************/

#ifndef _WIN32_DCOM
#define _WIN32_DCOM
#endif

#ifdef _UNICODE
	#define UNICODE
	#undef _MBCS
#endif

/******************************************************************************
/* Includes																	  *
******************************************************************************/
#include <windows.h>
#include <string.h>
#include <objbase.h>
#include <atlbase.h>
#include <iostream>
#include <stdio.h>
#include <Wbemidl.h>
#include <comutil.h>
#include <comdef.h>

#include "ConvFunc.h"

#pragma comment(lib, "wbemuuid.lib")

/******************************************************************************
/* Defines																	  *
******************************************************************************/
// for NSIS stack
#define EXDLL_INIT()				\
{									\
	g_stacktop		= stacktop;		\
	g_variables		= variables;	\
	g_stringsize	= string_size;	\
}
#define	MAX_STRLEN	1024
/******************************************************************************
/* NSIS stack structure														  *
******************************************************************************/
typedef struct _stack_t {
	struct	_stack_t *next;
	TCHAR	text[MAX_STRLEN];
} stack_t;

#define NSISPIAPIVER_1_0 0x00010000
#define NSISPIAPIVER_CURR NSISPIAPIVER_1_0

#ifndef NSISCALL
#  define NSISCALL __stdcall
#endif

typedef struct
{
  int autoclose;
  int all_user_var;
  int exec_error;
  int abort;
  int exec_reboot; // NSIS_SUPPORT_REBOOT
  int reboot_called; // NSIS_SUPPORT_REBOOT
  int XXX_cur_insttype; // depreacted
  int plugin_api_version; // see NSISPIAPIVER_CURR
                          // used to be XXX_insttype_changed
  int silent; // NSIS_CONFIG_SILENT_SUPPORT
  int instdir_error;
  int rtl;
  int errlvl;
  int alter_reg_view;
  int status_update;
} exec_flags_t;

typedef UINT_PTR (*NSISPLUGINCALLBACK)(enum NSPIM);

typedef struct {
  exec_flags_t *exec_flags;
  int (NSISCALL *ExecuteCodeSegment)(int, HWND);
  void (NSISCALL *validate_filename)(TCHAR *);
  BOOL (NSISCALL *RegisterPluginCallback)(HMODULE, NSISPLUGINCALLBACK);
} extra_parameters;


enum NSPIM 
{
	NSPIM_UNLOAD,    // This is the last message a plugin gets, do final cleanup
	NSPIM_GUIUNLOAD, // Called after .onGUIEnd
};

/******************************************************************************
/* Global values															  *
******************************************************************************/
// for NSIS stack
stack_t			**g_stacktop;
TCHAR			*g_variables;
unsigned int	g_stringsize;
// for plugin
HINSTANCE		g_hInstance;
TCHAR			wmiError[MAX_STRLEN]={0};
TCHAR			wmiResult[MAX_STRLEN]={0};
_variant_t		wmiArrayResult=0;
_variant_t		WMIResults[1024]={0};
int				WMIResultsNum = 0;

/******************************************************************************
* FUNCTION NAME: pushstring													  *
*     PARAMETER: STR														  *
*		PURPOSE: Removes the element from the top of the NSIS stack and puts  *
*				 it in the buffer.											  * 
*		 RETURN: TRUE if stack is empty, FALSE if value is loaded.			  *
******************************************************************************/
int popstring(TCHAR *str)
{
	stack_t *th;

	if (!g_stacktop || !*g_stacktop) return 1;

	th=(*g_stacktop);
	lstrcpy(str,th->text);
	*g_stacktop = th->next;
	GlobalFree((HGLOBAL)th);

	return 0;
}
/******************************************************************************
* FUNCTION NAME: pushstring													  *
*	  PARAMETER: STR														  *
*		PURPOSE: Adds an element to the top of the NSIS stack				  * 
******************************************************************************/
void pushstring(const TCHAR *str)
{
	stack_t *th;

	if (!g_stacktop) return;
	
	th=(stack_t*)GlobalAlloc(GPTR, sizeof(stack_t)+g_stringsize);
	lstrcpyn(th->text,str,g_stringsize);
	th->next=*g_stacktop;
	
	*g_stacktop=th;
}
/******************************************************************************
* FUNCTION NAME: PluginCallback												  *
*	 PARAMETERS: MSG														  *
******************************************************************************/
static UINT_PTR PluginCallback(enum NSPIM msg)
{
  return 0;
}
/******************************************************************************
* FUNCTION NAME: SetError													  *
*	 PARAMETERS: MESSAGE, HRES												  *
*		PURPOSE: Compose the error message with the error code returned by	  *
*				 WMI.														  *
******************************************************************************/
void SetError(const _bstr_t message, const long hres)
{
	TCHAR buf[MAX_STRLEN]=TEXT("");

	_tcscpy_s(wmiError, message);
	if(hres != 0)
	{
		wsprintf(buf,TEXT(" Error code = 0x%lX"),hres);
		_tcscat_s(wmiError,buf);
	}
}
/******************************************************************************
* INTERNAL FUNCTION: wmiRequest												  *
*		 PARAMETERS: NAMESPACE, TABLE, VALUE, WHERE (optional)				  *
*			PURPOSE: Retrieve the requested value of the table in the correct *
*					 namespace with WMI calls.						   		  *
*			 RETURN: TRUE if value is found, FALSE if error occurs.			  *
******************************************************************************/
bool wmiRequest(const _bstr_t& _namespace, const _bstr_t& _table, const _bstr_t& _value, const _bstr_t& _where)
{
	HRESULT	hres;
    // Obtain the initial locator to WMI.
    CComPtr< IWbemLocator > locator;
    hres = CoCreateInstance	(
							CLSID_WbemLocator,             
							NULL, 
							CLSCTX_INPROC_SERVER, 
							IID_IWbemLocator, 
							reinterpret_cast< void** >( &locator )
							);
    if (FAILED(hres))
    {
        SetError(TEXT("Failed to create IWbemLocator object."), hres);
        return false;
    }
    // Connect to WMI through the IWbemLocator::ConnectServer method.
    CComPtr< IWbemServices > service;
    const _bstr_t _path = TEXT("ROOT\\") + _namespace;
	hres = locator->ConnectServer	(
								_path,							// Object path of WMI namespace
								NULL,							// User name. NULL = current user
								NULL,							// User password. NULL = current
								NULL,							// Locale. NULL indicates current
								WBEM_FLAG_CONNECT_USE_MAX_WAIT, // Security flags.
								NULL,							// Authority (e.g. Kerberos)
								NULL,							// Context object 
								&service						// pointer to IWbemServices proxy
								);
	if (FAILED(hres))
	{
		SetError(TEXT("Could not connect."), hres);
		return false;
	}
	// Set security levels on the proxy
    hres = CoSetProxyBlanket	(
								service,						// Indicates the proxy to set
								RPC_C_AUTHN_WINNT,				// RPC_C_AUTHN_xxx
								RPC_C_AUTHZ_NONE,				// RPC_C_AUTHZ_xxx
								NULL,							// Server principal name 
								RPC_C_AUTHN_LEVEL_CALL,			// RPC_C_AUTHN_LEVEL_xxx 
								RPC_C_IMP_LEVEL_IMPERSONATE,	// RPC_C_IMP_LEVEL_xxx
								NULL,							// client identity
								EOAC_NONE						// proxy capabilities 
								);
    if (FAILED(hres))
    {
        SetError(TEXT("Could not set proxy blanket."), hres);
		return false;
    }
	// Use the IWbemServices pointer to make requests of WMI.
    CComPtr< IEnumWbemClassObject > enumerator;
	_bstr_t _request;
	if(_where != _bstr_t(TEXT("")))
	{
		_request = TEXT("SELECT * FROM ") + _table + TEXT(" WHERE ") + _where;
	}
	else
	{
		_request = TEXT("SELECT * FROM ") + _table;
	}
	hres = service->ExecQuery	(
								bstr_t(TEXT("WQL")), 
								_request,
								WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
								NULL,
								&enumerator
								);    
    if (FAILED(hres))
    {
        SetError(TEXT("Query failed."), hres);
        return false;
    }
    // Get the data from the query.
    CComPtr< IWbemClassObject > value = NULL;
	ULONG uReturn = 0;
	HRESULT hEnumRes = WBEM_S_NO_ERROR;
	_stprintf_s(wmiResult, TEXT(""));
	while (WBEM_S_NO_ERROR == hEnumRes)
    {
		uReturn = 0;
		value = NULL;
		hEnumRes = enumerator->Next(WBEM_INFINITE, 1L, &value, &uReturn);
        if(FAILED(hEnumRes) || value == NULL)
        {
            break;
        }
		_variant_t var = new _variant_t();
        // Get the value of the property
        hres = value->Get(_value, 0, &var, 0, 0);
		if(FAILED(hres))
		{
			SetError(TEXT("Unable to find requested value."), hres);
			return false;
		}
		_bstr_t str;
		TCHAR *cvar;
		switch (var.vt)
		{
			case VT_EMPTY: // nothing
			case VT_NULL: // SQL style Null
				_stprintf_s(wmiResult, TEXT(""));
				break;
			case VT_UI1: // unsigned char
				_stprintf_s(wmiResult, TEXT("%u"), var.cVal);
				break;
			case VT_UI2: // unsigned short	
				_stprintf_s(wmiResult, TEXT("%u"), var.iVal);
				break;
			case VT_UI4: // unsigned long
				_stprintf_s(wmiResult, TEXT("%u"), var.ulVal);
				break;
			case VT_I2: // 2 byte signed int
				
			case VT_I4: // 4 byte signed int
				_stprintf_s(wmiResult, TEXT("%i"), var.lVal);
				break;
			case VT_BOOL: // True=-1, False=0
				_stprintf_s(wmiResult, TEXT("%s"), (var.boolVal) ? TEXT("Yes") : TEXT("No"));
				break;
			case VT_BSTR|VT_ARRAY: // OLE Automation string array
				wmiArrayResult = var;
				break;
			case VT_BSTR: // OLE Automation string 
				str = var.bstrVal;
				cvar=str;
				_tcscpy_s(wmiResult,cvar);
				break;
			default: 
				_stprintf_s(wmiResult, TEXT("UNHANDLED VARIANT TYPE (%i)"), (int)var.vt);
				break;
        }
    }
	return true;
}
/******************************************************************************
* INTERNAL FUNCTION: wmiMultiRequest										  *
*		 PARAMETERS: NAMESPACE, TABLE, VALUE, WHERE (optional)				  *
*			PURPOSE: Retrieve an array of requested values of the table in	  *
*					 the correct namespace with WMI calls.			   		  *
*			 RETURN: TRUE if value is found, FALSE if error occurs.			  *
******************************************************************************/
bool wmiMultiRequest(const _bstr_t& _namespace, const _bstr_t& _table, const _bstr_t& _value, const _bstr_t& _where)
{
	HRESULT	hres;
    // Obtain the initial locator to WMI.
    CComPtr< IWbemLocator > locator;
    hres = CoCreateInstance	(
							CLSID_WbemLocator,             
							NULL, 
							CLSCTX_INPROC_SERVER, 
							IID_IWbemLocator, 
							reinterpret_cast< void** >( &locator )
							);
    if (FAILED(hres))
    {
        SetError(TEXT("Failed to create IWbemLocator object."), hres);
        return false;
    }
    // Connect to WMI through the IWbemLocator::ConnectServer method.
    CComPtr< IWbemServices > service;
    const _bstr_t _path = TEXT("ROOT\\") + _namespace;
	hres = locator->ConnectServer	(
								_path,							// Object path of WMI namespace
								NULL,							// User name. NULL = current user
								NULL,							// User password. NULL = current
								NULL,							// Locale. NULL indicates current
								WBEM_FLAG_CONNECT_USE_MAX_WAIT, // Security flags.
								NULL,							// Authority (e.g. Kerberos)
								NULL,							// Context object 
								&service						// pointer to IWbemServices proxy
								);
	if (FAILED(hres))
	{
		SetError(TEXT("Could not connect."), hres);
		return false;
	}
	// Set security levels on the proxy
    hres = CoSetProxyBlanket	(
								service,						// Indicates the proxy to set
								RPC_C_AUTHN_WINNT,				// RPC_C_AUTHN_xxx
								RPC_C_AUTHZ_NONE,				// RPC_C_AUTHZ_xxx
								NULL,							// Server principal name 
								RPC_C_AUTHN_LEVEL_CALL,			// RPC_C_AUTHN_LEVEL_xxx 
								RPC_C_IMP_LEVEL_IMPERSONATE,	// RPC_C_IMP_LEVEL_xxx
								NULL,							// client identity
								EOAC_NONE						// proxy capabilities 
								);
    if (FAILED(hres))
    {
        SetError(TEXT("Could not set proxy blanket."), hres);
		return false;
    }
	// Use the IWbemServices pointer to make requests of WMI.
    CComPtr< IEnumWbemClassObject > enumerator;
	_bstr_t _request;
	if(_where != _bstr_t(TEXT("")))
	{
		_request = TEXT("SELECT * FROM ") + _table + TEXT(" WHERE ") + _where;
	}
	else
	{
		_request = TEXT("SELECT * FROM ") + _table;
	}
	hres = service->ExecQuery	(
								bstr_t(TEXT("WQL")), 
								_request,
								WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
								NULL,
								&enumerator
								);    
    if (FAILED(hres))
    {
        SetError(TEXT("Query failed."), hres);
        return false;
    }
    // Get the data from the query.
    CComPtr< IWbemClassObject > value = NULL;
	ULONG uReturn = 0;
	HRESULT hEnumRes = WBEM_S_NO_ERROR;
	WMIResultsNum = 0;
	while (WBEM_S_NO_ERROR == hEnumRes)
    {
		uReturn = 0;
		value = NULL;
		hEnumRes = enumerator->Next(WBEM_INFINITE, 1L, &value, &uReturn);
        if(FAILED(hEnumRes) || value == NULL)
        {
            break;
        }
		_variant_t var = new _variant_t();
        // Get the value of the property
        hres = value->Get(_value, 0, &var, 0, 0);
		if(FAILED(hres))
		{
			SetError(TEXT("Unable to find requested value."), hres);
			return false;
		}
		WMIResults[WMIResultsNum]=var;
		WMIResultsNum++;
    }
	return true;
}
/******************************************************************************
* INTERNAL FUNCTION: wmiSearchInStringArray									  *
*		 PARAMETERS: NAMESPACE, TABLE, SEARCHVALUENAME, WHERENAME, WHEREVALUE *
*			PURPOSE: Search for a value in an array (VT_ARRAY).		   		  *
*			 RETURN: TRUE if value is found, FALSE if error occurs.			  *
******************************************************************************/
bool wmiSearchInStringArray (const _bstr_t& _namespace, const _bstr_t& _table, const _bstr_t& _searchvaluename, const _bstr_t& _wherename, const _bstr_t& _wherevalue)
{
	HRESULT	hres;
    // Obtain the initial locator to WMI.
    CComPtr< IWbemLocator > locator;
    hres = CoCreateInstance	(
							CLSID_WbemLocator,             
							NULL, 
							CLSCTX_INPROC_SERVER, 
							IID_IWbemLocator, 
							reinterpret_cast< void** >( &locator )
							);
    if (FAILED(hres))
    {
        SetError(TEXT("Failed to create IWbemLocator object."), hres);
        return false;
    }
    // Connect to WMI through the IWbemLocator::ConnectServer method.
    CComPtr< IWbemServices > service;
    const _bstr_t _path = TEXT("ROOT\\") + _namespace;
	hres = locator->ConnectServer	(
								_path,							// Object path of WMI namespace
								NULL,							// User name. NULL = current user
								NULL,							// User password. NULL = current
								NULL,							// Locale. NULL indicates current
								WBEM_FLAG_CONNECT_USE_MAX_WAIT, // Security flags.
								NULL,							// Authority (e.g. Kerberos)
								NULL,							// Context object 
								&service						// pointer to IWbemServices proxy
								);
	if (FAILED(hres))
	{
		SetError(TEXT("Could not connect."), hres);
		return false;
	}
	// Set security levels on the proxy
    hres = CoSetProxyBlanket	(
								service,						// Indicates the proxy to set
								RPC_C_AUTHN_WINNT,				// RPC_C_AUTHN_xxx
								RPC_C_AUTHZ_NONE,				// RPC_C_AUTHZ_xxx
								NULL,							// Server principal name 
								RPC_C_AUTHN_LEVEL_CALL,			// RPC_C_AUTHN_LEVEL_xxx 
								RPC_C_IMP_LEVEL_IMPERSONATE,	// RPC_C_IMP_LEVEL_xxx
								NULL,							// client identity
								EOAC_NONE						// proxy capabilities 
								);
    if (FAILED(hres))
    {
        SetError(TEXT("Could not set proxy blanket."), hres);
		return false;
    }
	// Use the IWbemServices pointer to make requests of WMI.
    CComPtr< IEnumWbemClassObject > enumerator;
	_bstr_t _request;
	_request = TEXT("SELECT * FROM ") + _table;
	hres = service->ExecQuery	(
								bstr_t(TEXT("WQL")), 
								_request,
								WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
								NULL,
								&enumerator
								);    
    if (FAILED(hres))
    {
        SetError(TEXT("Query failed."), hres);
        return false;
    }
    // Get the data from the query.
    CComPtr< IWbemClassObject > value = NULL;
	ULONG uReturn = 0;
	HRESULT hEnumRes = WBEM_S_NO_ERROR;
	WMIResultsNum = 0;
	while (WBEM_S_NO_ERROR == hEnumRes)
    {
		uReturn = 0;
		value = NULL;
		hEnumRes = enumerator->Next(WBEM_INFINITE, 1L, &value, &uReturn);
        if(FAILED(hEnumRes) || value == NULL)
        {
            break;
        }
		_variant_t var = new _variant_t();
        // Get the value of the property
        hres = value->Get(_wherename, 0, &var, 0, 0);
		if(FAILED(hres))
		{
			SetError(TEXT("Unable to find requested value."), hres);
			return false;
		}
		if(var.vt == (VT_BSTR|VT_ARRAY))
		{
			LONG lBound = 0;
			LONG uBound = 0;
			SafeArrayGetLBound(var.parray, 1, &lBound);
			SafeArrayGetUBound(var.parray, 1, &uBound);
			for(LONG j = lBound; j <= uBound; j++)
			{
				BSTR arrayitemvalue;
				SafeArrayGetElement(var.parray, &j, &arrayitemvalue);
				if(_tcscmp(_bstr_t(arrayitemvalue),_wherevalue) == 0)
				{
					hres = value->Get(_searchvaluename, 0, &var, 0, 0);
					if(FAILED(hres))
					{
						SetError(TEXT("Unable to find requested value."), hres);
						return false;
					}
					_bstr_t str;
					TCHAR *cvar;
					switch (var.vt)
					{
						case VT_EMPTY: // nothing
						case VT_NULL: // SQL style Null
							_stprintf_s(wmiResult, TEXT(""));
							break;
						case VT_UI1: // unsigned char
							_stprintf_s(wmiResult, TEXT("%u"), var.cVal);
							break;
						case VT_UI2: // unsigned short	
							_stprintf_s(wmiResult, TEXT("%u"), var.iVal);
							break;
						case VT_UI4: // unsigned long
							_stprintf_s(wmiResult, TEXT("%u"), var.ulVal);
							break;
						case VT_I2: // 2 byte signed int
							
						case VT_I4: // 4 byte signed int
							_stprintf_s(wmiResult, TEXT("%i"), var.lVal);
							break;
						case VT_BOOL: // True=-1, False=0
							_stprintf_s(wmiResult, TEXT("%s"), (var.boolVal) ? TEXT("Yes") : TEXT("No"));
							break;
						case VT_BSTR|VT_ARRAY: // OLE Automation string array
							wmiArrayResult = var;
							break;
						case VT_BSTR: // OLE Automation string 
							str = var.bstrVal;
							cvar=str;
							_tcscpy_s(wmiResult,cvar);
							break;
						default: 
							_stprintf_s(wmiResult, TEXT("UNHANDLED VARIANT TYPE (%i)"), (int)var.vt);
							break;
					}
					return true;
				}
			}
		}
		else if (var.vt != VT_NULL)
		{
			SetError(TEXT("The requested value is not a string array."), 0);
			return false;
		}
    }
	SetError(TEXT("Requested value not found."), 0);
	return false;
}
/******************************************************************************
* INTERNAL FUNCTION: GetRegistryStringValue									  *
*		 PARAMETERS: SUBKEYNAME, VALUENAME									  *
*			PURPOSE: Retrieve the requested string value the registry using	  *
*			WMI calls.												   		  *
*			 RETURN: TRUE if value is found, FALSE if error occurs.			  *
******************************************************************************/
bool GetRegistryStringValue(const _bstr_t SubKeyName, const _bstr_t ValueName)
{
	HRESULT	hres;
	IWbemLocator *ppiWmiLoc = NULL;

	hres = CoCreateInstance(CLSID_WbemLocator,0,CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID *) &ppiWmiLoc);

	IWbemServices *pSvc = NULL;
	CComBSTR bstrserv(TEXT("\\\\.\\root\\default"));

	hres = ppiWmiLoc->ConnectServer(bstrserv, NULL, NULL,0, NULL, 0, 0, &pSvc);
	if(SUCCEEDED(hres))
	{
		hres = CoSetProxyBlanket(pSvc, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE,	NULL, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE);

		if(SUCCEEDED(hres))
		{
			BSTR MethodName = SysAllocString(L"GetStringValue");
			BSTR ClassName = SysAllocString(L"StdRegProv");

			IWbemClassObject* pClass = NULL;
			hres = pSvc->GetObject(ClassName, 0, NULL, &pClass, NULL);

			if(SUCCEEDED(hres))
			{
				IWbemClassObject* pInParamsDefinition = NULL;
				hres = pClass->GetMethod(MethodName, 0, &pInParamsDefinition, NULL);

				if(SUCCEEDED(hres))
				{
					IWbemClassObject* pClassInstance = NULL;
					hres = pInParamsDefinition->SpawnInstance(0, &pClassInstance);

					if(SUCCEEDED(hres))
					{
						// Create the values for the in parameters
						VARIANT varCommand;
						varCommand.vt = VT_UINT;
						varCommand.uintVal = 0x80000002;

						// Store the value for the in parameters
						hres = pClassInstance->Put(L"hDefKey", 0, &varCommand, 0);

						// Create the values for the in parameters
						varCommand.vt = VT_BSTR;
						varCommand.bstrVal = SubKeyName;

						// Store the value for the in parameters
						hres = pClassInstance->Put(L"sSubKeyName", 0, &varCommand, 0);

						// Create the values for the in parameters

						varCommand.vt = VT_BSTR;
						varCommand.bstrVal = ValueName;

						// Store the value for the in parameters
						hres = pClassInstance->Put(L"sValueName", 0, &varCommand, 0);

						// Execute Method
						IWbemClassObject* pOutParams = NULL;
						hres = pSvc->ExecMethod(ClassName, MethodName, 0, NULL, pClassInstance, &pOutParams, NULL);

						if (FAILED(hres))
						{
							SetError(TEXT("ExecMethod failed."), hres);
						}
						else
						{
							//BSTR ret;
							//hres = pOutParams->GetObjectText(0, &ret);
							_variant_t varReturnValue;
							CComBSTR bstrret(TEXT("ReturnValue"));
							hres = pOutParams->Get(bstrret, 0, &varReturnValue, NULL, 0);
							if(varReturnValue.ulVal == 0)
							{
								CComBSTR bstrret(TEXT("sValue"));
								hres = pOutParams->Get(bstrret, 0, &varReturnValue, NULL, 0);
								_bstr_t str;
								TCHAR *cvar;
								str = varReturnValue.bstrVal;
								cvar=str;
								_tcscpy_s(wmiResult,cvar);
								return true;
							}
							else if(varReturnValue.ulVal == 1)
							{
								_tcscpy_s(wmiResult,TEXT(""));
								return true;
							}
							else
							{
								SetError(TEXT("GetStringValue failed."), varReturnValue.ulVal);
							}
							pOutParams->Release();
						}	
						pClassInstance->Release();
					}
					else SetError(TEXT("SpawnInstance failed."), hres);
					pInParamsDefinition->Release();
				}
				else SetError(TEXT("GetMethod failed."), hres);
				pClass->Release();
			}
			else SetError(TEXT("GetObject failed."), hres);
		}
		else SetError(TEXT("CoSetProxyBlanket failed."), hres);
		pSvc->Release();
	}
	else SetError(TEXT("ConnectServer failed."), hres);
	return false;
}
/******************************************************************************
* INTERNAL FUNCTION: GetRegistryDWORDValue									  *
*		 PARAMETERS: SUBKEYNAME, VALUENAME									  *
*			PURPOSE: Retrieve the requested DWORD value the registry using	  *
*			WMI calls.												   		  *
*			 RETURN: TRUE if value is found, FALSE if error occurs.			  *
******************************************************************************/
DWORD GetRegistryDWORDValue (const _bstr_t SubKeyName, const _bstr_t ValueName)
{
	HRESULT	hres;
	IWbemLocator *ppiWmiLoc = NULL;

	hres = CoCreateInstance(CLSID_WbemLocator,0,CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID *) &ppiWmiLoc);

	IWbemServices *pSvc = NULL;
	CComBSTR bstrserv(TEXT("\\\\.\\root\\default"));

	hres = ppiWmiLoc->ConnectServer(bstrserv, NULL, NULL,0, NULL, 0, 0, &pSvc);
	if(SUCCEEDED(hres))
	{
		hres = CoSetProxyBlanket(pSvc, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE,	NULL, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE);

		if(SUCCEEDED(hres))
		{
			BSTR MethodName = SysAllocString(L"GetDWORDValue");
			BSTR ClassName = SysAllocString(L"StdRegProv");

			IWbemClassObject* pClass = NULL;
			hres = pSvc->GetObject(ClassName, 0, NULL, &pClass, NULL);

			if(SUCCEEDED(hres))
			{
				IWbemClassObject* pInParamsDefinition = NULL;
				hres = pClass->GetMethod(MethodName, 0, &pInParamsDefinition, NULL);

				if(SUCCEEDED(hres))
				{
					IWbemClassObject* pClassInstance = NULL;
					hres = pInParamsDefinition->SpawnInstance(0, &pClassInstance);

					if(SUCCEEDED(hres))
					{
						// Create the values for the in parameters
						VARIANT varCommand;
						varCommand.vt = VT_UINT;
						varCommand.uintVal = 0x80000002;

						// Store the value for the in parameters
						hres = pClassInstance->Put(L"hDefKey", 0, &varCommand, 0);

						// Create the values for the in parameters
						varCommand.vt = VT_BSTR;
						varCommand.bstrVal = SubKeyName;

						// Store the value for the in parameters
						hres = pClassInstance->Put(L"sSubKeyName", 0, &varCommand, 0);

						// Create the values for the in parameters

						varCommand.vt = VT_BSTR;
						varCommand.bstrVal = ValueName;

						// Store the value for the in parameters
						hres = pClassInstance->Put(L"sValueName", 0, &varCommand, 0);

						// Execute Method
						IWbemClassObject* pOutParams = NULL;
						hres = pSvc->ExecMethod(ClassName, MethodName, 0, NULL, pClassInstance, &pOutParams, NULL);

						if (FAILED(hres))
						{
							SetError(TEXT("ExecMethod failed."), hres);
						}
						else
						{
							_variant_t varReturnValue;
							CComBSTR bstrret(TEXT("ReturnValue"));
							hres = pOutParams->Get(bstrret, 0, &varReturnValue, NULL, 0);
							if(varReturnValue.ulVal == 0)
							{
								CComBSTR bstrret(TEXT("sValue"));
								hres = pOutParams->Get(bstrret, 0, &varReturnValue, NULL, 0);
								return varReturnValue.ulVal;
							}
							else if(varReturnValue.ulVal == 1)
							{
								return -2;
							}
							else
							{
								SetError(TEXT("GetStringValue failed."), varReturnValue.ulVal);
							}
							pOutParams->Release();
						}	
						pClassInstance->Release();
					}
					else SetError(TEXT("SpawnInstance failed."), hres);
					pInParamsDefinition->Release();
				}
				else SetError(TEXT("GetMethod failed."), hres);
				pClass->Release();
			}
			else SetError(TEXT("GetObject failed."), hres);
		}
		else SetError(TEXT("CoSetProxyBlanket failed."), hres);
		pSvc->Release();
	}
	else SetError(TEXT("ConnectServer failed."), hres);
	return -1;
}
/******************************************************************************
* INTERNAL FUNCTION: StackResult											  *
*		  PARAMETER: RESULT													  *
*			PURPOSE: Stack returned value for NSIS plugin		 			  * 
******************************************************************************/
void StackResult(bool result)
{
	switch(result)
		{
			case true: 
					pushstring(wmiResult);
					pushstring(TEXT("ok"));		
					break;
			case false: 
					pushstring(wmiError);
					pushstring(TEXT("error"));
				break;
		}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetHostName											  *
*		 PARAMETERS: 														  *
*			PURPOSE: Return the hostname of the computer.		  		 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetHostName(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		StackResult(GetRegistryStringValue(TEXT("SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters"), TEXT("Hostname")));
	}
}

/******************************************************************************
* EXTERNAL FUNCTION: GetPrimaryDNSSuffix				  					  *
*		 PARAMETERS: 														  *
*			PURPOSE: Return the Primary DNS Suffix of the computer.		 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetPrimaryDNSSuffix(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		StackResult(GetRegistryStringValue(TEXT("SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters"), TEXT("Domain")));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNodeType						  					  *
*		 PARAMETERS: 														  *
*			PURPOSE: Return the Node Type of the computer.		  		 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNodeType(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		DWORD dwNodeType = GetRegistryDWORDValue(TEXT("SYSTEM\\CurrentControlSet\\Services\\NetBT\\Parameters"), TEXT("DhcpNodeType"));
		if(dwNodeType == -1)StackResult(false);
		else 
		{
			if(dwNodeType == 1)_stprintf_s(wmiResult, TEXT("Broadcast"));
			else if(dwNodeType == 2)_stprintf_s(wmiResult, TEXT("Peer-Peer"));
			else if(dwNodeType == 4)_stprintf_s(wmiResult, TEXT("Mixed"));
			else if(dwNodeType == 8)_stprintf_s(wmiResult, TEXT("Hybrid"));
			else _stprintf_s(wmiResult, TEXT("Unknown"));
			StackResult(true);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: IsIPRoutingEnabled							  			  *
*		 PARAMETERS: 														  *
*			PURPOSE: Check if IP Routing is enabled on the computer.		  *
******************************************************************************/
extern "C"
void __declspec(dllexport) IsIPRoutingEnabled(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		DWORD dwNodeType = GetRegistryDWORDValue(TEXT("SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters"), TEXT("IPEnableRouter"));
		if(dwNodeType == -1)StackResult(false);
		else 
		{
			if(dwNodeType == 0)_stprintf_s(wmiResult, TEXT("No"));
			else if(dwNodeType == 1)_stprintf_s(wmiResult, TEXT("Yes"));
			else _stprintf_s(wmiResult, TEXT("?"));
			StackResult(true);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: IsWINSProxyEnabled					  					  *
*		 PARAMETERS: 														  *
*			PURPOSE: Check if WINS Proxy is enabled on the computer.		  *
******************************************************************************/
extern "C"
void __declspec(dllexport) IsWINSProxyEnabled(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("DNSEnabledForWINSResolution"), TEXT("IPEnabled = True")));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetDNSSuffixSearchList				  					  *
*		 PARAMETERS: 														  *
*			PURPOSE: Return the DNS Suffix SearchList of the computer.	 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetDNSSuffixSearchList(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		StackResult(GetRegistryStringValue(TEXT("SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters"), TEXT("SearchList")));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetAllNetworkAdaptersIDsCB			  					  *
*		 PARAMETERS: CALLBACKADDRESS										  *
*			PURPOSE: Find all ID's from installed Networkadapters.		 	  *
*					 Calls the Callback function for every result.			  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetAllNetworkAdaptersIDsCB(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szCallbackAddress[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	int CallbackAddress = 0;
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szCallbackAddress);
		CallbackAddress=_tstoi(szCallbackAddress);
		if(wmiMultiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("Index"), TEXT("")) == true)
		{
			int results = 0;
			for(int i = 0; i < WMIResultsNum; i++)
			{
				if(WMIResults[i].vt == VT_I4)
				{
					results++;
					_ltot_s(WMIResults[i].lVal, szBuf, 10);
					pushstring(szBuf);
					extra->ExecuteCodeSegment(CallbackAddress - 1, 0);
					if (extra->exec_flags->exec_error){}
				}
			}
			_stprintf_s(wmiResult, TEXT("%i"), results);
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetAllNetworkAdaptersIDs			  					  *
*		 PARAMETERS:														  *
*			PURPOSE: Find all  ID's from installed Networkadapters.		   	  *
*					 All results are returned in a string (space seperated).  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetAllNetworkAdaptersIDs(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szBuf[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		if(wmiMultiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("Index"), TEXT("")) == true)
		{
			for(int i = 0; i < WMIResultsNum; i++)
			{
				if(WMIResults[i].vt == VT_I4)
				{
					_ltot_s(WMIResults[i].lVal, szBuf, 10);
					if(i == 0)_stprintf_s(wmiResult, TEXT("%s"), szBuf);
					else _stprintf_s(wmiResult, TEXT("%s %s"), wmiResult, szBuf);
				}
			}
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetEnabledNetworkAdaptersIDsCB		  					  *
*		 PARAMETERS: CALLBACKADDRESS										  *
*			PURPOSE: Find all ID's from enabled installed Networkadapters.    *
*					 Calls the Callback function for every result.			  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetEnabledNetworkAdaptersIDsCB(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szCallbackAddress[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	int CallbackAddress = 0;
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szCallbackAddress);
		CallbackAddress=_tstoi(szCallbackAddress);
		if(wmiMultiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("Index"), TEXT("IPEnabled = True")) == true)
		{
			int results = 0;
			for(int i = 0; i < WMIResultsNum; i++)
			{
				if(WMIResults[i].vt == VT_I4)
				{
					results++;
					_ltot_s(WMIResults[i].lVal, szBuf, 10);
					pushstring(szBuf);
					extra->ExecuteCodeSegment(CallbackAddress - 1, 0);
					if (extra->exec_flags->exec_error){}
				}
			}
			_stprintf_s(wmiResult, TEXT("Number of results: %i"), results);
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetEnabledNetworkAdaptersIDs		  					  *
*		 PARAMETERS:														  *
*			PURPOSE: Find all ID's from enabled installed Networkadapters. 	  *
*					 All results are returned in a string (space seperated).  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetEnabledNetworkAdaptersIDs(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szBuf[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		if(wmiMultiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("Index"), TEXT("IPEnabled = True")) == true)
		{
			for(int i = 0; i < WMIResultsNum; i++)
			{
				if(WMIResults[i].vt == VT_I4)
				{
					_ltot_s(WMIResults[i].lVal, szBuf, 10);
					if(i == 0)_stprintf_s(wmiResult, TEXT("%s"), szBuf);
					else _stprintf_s(wmiResult, TEXT("%s %s"), wmiResult, szBuf);
				}
			}
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterIDFromDescription						  *
*		 PARAMETERS: DESCRIPTION											  *
*			PURPOSE: Return the ID from the NetworkAdapter with the given	  *
*					 Description								  		 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterIDFromDescription(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szSearchString[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szSearchString);
		_bstr_t _where = TEXT("Description = '") + _bstr_t(szSearchString) + TEXT("'");
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapter"), TEXT("DeviceID"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterIDFromMACAddress						  *
*		 PARAMETERS: MACADDRESS												  *
*			PURPOSE: Return the ID from the NetworkAdapter with the given	  *
*					 MAC Address								  		 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterIDFromMACAddress(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szSearchString[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szSearchString);
		_bstr_t _where = TEXT("MACAddress = '") + _bstr_t(szSearchString) + TEXT("'");
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("Index"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterIDFromIPAddress						  *
*		 PARAMETERS: IPADDRESS												  *
*			PURPOSE: Return the ID from the NetworkAdapter with the given	  *
*					 IP Address									  		 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterIDFromIPAddress(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szSearchString[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szSearchString);
		StackResult(wmiSearchInStringArray(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("Index"), TEXT("IPAddress"), szSearchString));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterType				  					  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Return the Type of the given Networkadapter.		 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterType(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("DeviceID = ") + _bstr_t(szNetworkAdapterID);
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapter"), TEXT("AdapterType"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterConnectionID		  					  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Return the Connection ID of the given Networkadapter. 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterConnectionID(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("DeviceID = ") + _bstr_t(szNetworkAdapterID);
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapter"), TEXT("NetConnectionID"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterConnectionSpecificDNSSuffix			  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Return the DNS Suffix of the given Networkadapter. 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterConnectionSpecificDNSSuffix(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("DNSDomain"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterDescription							  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Return the Description of the given Networkadapter. 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterDescription(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("DeviceID = ") + _bstr_t(szNetworkAdapterID);
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapter"), TEXT("Description"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterMACAddress							  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Return the MAC Address of the given Networkadapter. 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterMACAddress(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("MACAddress"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: IsNetworkAdapterDHCPEnabled							  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Check if DHCP is enabled for the given Networkadapter.   *
******************************************************************************/
extern "C"
void __declspec(dllexport) IsNetworkAdapterDHCPEnabled(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("DHCPEnabled"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: IsNetworkAdapterAutoSense								  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Check if AutoSense is enabled for the given			  *
*					 Networkadapter.			  							  *
******************************************************************************/
extern "C"
void __declspec(dllexport) IsNetworkAdapterAutoSense(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("DeviceID = ") + _bstr_t(szNetworkAdapterID);
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapter"), TEXT("AutoSense"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterIPAddressesCB		  					  *
*		 PARAMETERS: NETWORKADAPTERID, CALLBACKADDRESS						  *
*			PURPOSE: Find all IP addresses's for the given Networkadapter.	  *
*					 Calls the Callback function for every result.			  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterIPAddressesCB(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szCallbackAddress[MAX_STRLEN]={0};
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	int CallbackAddress = 0;
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		popstring(szCallbackAddress);
		CallbackAddress=_tstoi(szCallbackAddress);
		if(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("IPAddress"), _where) == true)
		{
			int results = 0;
			if(wmiArrayResult.vt == (VT_BSTR|VT_ARRAY))
			{
				LONG lBound = 0;
				LONG uBound = 0;
				SafeArrayGetLBound(wmiArrayResult.parray, 1, &lBound);
				SafeArrayGetUBound(wmiArrayResult.parray, 1, &uBound);
				for(LONG j = lBound; j <= uBound; j++)
				{
					results++;
					BSTR var;
					SafeArrayGetElement(wmiArrayResult.parray, &j, &var);
					pushstring(_bstr_t(var));
					extra->ExecuteCodeSegment(CallbackAddress - 1, 0);
					if (extra->exec_flags->exec_error){}
				}
			}
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterIPAddresses		  					  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Find all IP addresses's for the given Networkadapter.	  *
*					 All results are returned in a string (space seperated).  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterIPAddresses(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		if(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("IPAddress"), _where) == true)
		{
			if(wmiArrayResult.vt == (VT_BSTR|VT_ARRAY))
			{
				LONG lBound = 0;
				LONG uBound = 0;
				SafeArrayGetLBound(wmiArrayResult.parray, 1, &lBound);
				SafeArrayGetUBound(wmiArrayResult.parray, 1, &uBound);
				for(LONG j = lBound; j <= uBound; j++)
				{
					BSTR var;
					SafeArrayGetElement(wmiArrayResult.parray, &j, &var);
					_bstr_t _var = _bstr_t(var);
					if(j == lBound)_tcscpy_s(wmiResult,(TCHAR*)_var);
					else _stprintf_s(wmiResult, TEXT("%s %s"), wmiResult, (TCHAR*)_var);
					BSTR i = 0;
				}
			}
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterIPSubNetsCB		  					  *
*		 PARAMETERS: NETWORKADAPTERID, CALLBACKADDRESS						  *
*			PURPOSE: Find all IP subnets for the given Networkadapter.		  *
*					 Calls the Callback function for every result.			  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterIPSubNetsCB(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szCallbackAddress[MAX_STRLEN]={0};
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	int CallbackAddress = 0;
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		popstring(szCallbackAddress);
		CallbackAddress=_tstoi(szCallbackAddress);
		if(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("IPSubnet"), _where) == true)
		{
			int results = 0;
			if(wmiArrayResult.vt == (VT_BSTR|VT_ARRAY))
			{
				LONG lBound = 0;
				LONG uBound = 0;
				SafeArrayGetLBound(wmiArrayResult.parray, 1, &lBound);
				SafeArrayGetUBound(wmiArrayResult.parray, 1, &uBound);
				for(LONG j = lBound; j <= uBound; j++)
				{
					results++;
					BSTR var;
					SafeArrayGetElement(wmiArrayResult.parray, &j, &var);
					pushstring(_bstr_t(var));
					extra->ExecuteCodeSegment(CallbackAddress - 1, 0);
					if (extra->exec_flags->exec_error){}
				}
			}
			_stprintf_s(wmiResult, TEXT("Number of results: %i"), results);
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterIPSubNets			  					  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Find all IP subnets for the given Networkadapter.		  *
*					 All results are returned in a string (space seperated).  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterIPSubNets(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		if(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("IPSubnet"), _where) == true)
		{
			if(wmiArrayResult.vt == (VT_BSTR|VT_ARRAY))
			{
				LONG lBound = 0;
				LONG uBound = 0;
				SafeArrayGetLBound(wmiArrayResult.parray, 1, &lBound);
				SafeArrayGetUBound(wmiArrayResult.parray, 1, &uBound);
				for(LONG j = lBound; j <= uBound; j++)
				{
					BSTR var;
					SafeArrayGetElement(wmiArrayResult.parray, &j, &var);
					_bstr_t _var = _bstr_t(var);
					if(j == lBound)_tcscpy_s(wmiResult,(TCHAR*)_var);
					else _stprintf_s(wmiResult, TEXT("%s %s"), wmiResult, (TCHAR*)_var);
					BSTR i = 0;
				}
			}
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterDefaultIPGatewaysCB		  			  *
*		 PARAMETERS: NETWORKADAPTERID, CALLBACKADDRESS						  *
*			PURPOSE: Find all Default IP Gateways for the given				  * 
*					 Networkadapter.										  *
*					 Calls the Callback function for every result.			  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterDefaultIPGatewaysCB(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szCallbackAddress[MAX_STRLEN]={0};
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	int CallbackAddress = 0;
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		popstring(szCallbackAddress);
		CallbackAddress=_tstoi(szCallbackAddress);
		if(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("DefaultIPGateway"), _where) == true)
		{
			int results = 0;
			if(wmiArrayResult.vt == (VT_BSTR|VT_ARRAY))
			{
				LONG lBound = 0;
				LONG uBound = 0;
				SafeArrayGetLBound(wmiArrayResult.parray, 1, &lBound);
				SafeArrayGetUBound(wmiArrayResult.parray, 1, &uBound);
				for(LONG j = lBound; j <= uBound; j++)
				{
					results++;
					BSTR var;
					SafeArrayGetElement(wmiArrayResult.parray, &j, &var);
					pushstring(_bstr_t(var));
					extra->ExecuteCodeSegment(CallbackAddress - 1, 0);
					if (extra->exec_flags->exec_error){}
				}
			}
			_stprintf_s(wmiResult, TEXT("Number of results: %i"), results);
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterDefaultIPGateways	  					  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Find all Default IP Gateways for the given				  * 
*					 Networkadapter.										  *
*					 All results are returned in a string (space seperated).  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterDefaultIPGateways(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		if(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("DefaultIPGateway"), _where) == true)
		{
			if(wmiArrayResult.vt == (VT_BSTR|VT_ARRAY))
			{
				LONG lBound = 0;
				LONG uBound = 0;
				SafeArrayGetLBound(wmiArrayResult.parray, 1, &lBound);
				SafeArrayGetUBound(wmiArrayResult.parray, 1, &uBound);
				for(LONG j = lBound; j <= uBound; j++)
				{
					BSTR var;
					SafeArrayGetElement(wmiArrayResult.parray, &j, &var);
					_bstr_t _var = _bstr_t(var);
					if(j == lBound)_tcscpy_s(wmiResult,(TCHAR*)_var);
					else _stprintf_s(wmiResult, TEXT("%s %s"), wmiResult, (TCHAR*)_var);
					BSTR i = 0;
				}
			}
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterDHCPServer							  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Return the DHCP Server of the given Networkadapter. 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterDHCPServer(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("DHCPServer"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterDNSServersCB				  			  *
*		 PARAMETERS: NETWORKADAPTERID, CALLBACKADDRESS						  *
*			PURPOSE: Find all DNS Servers for the given	Networkadapter.		  * 
*					 Calls the Callback function for every result.			  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterDNSServersCB(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szCallbackAddress[MAX_STRLEN]={0};
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	int CallbackAddress = 0;
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		popstring(szCallbackAddress);
		CallbackAddress=_tstoi(szCallbackAddress);
		if(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("DNSServerSearchOrder"), _where) == true)
		{
			int results = 0;
			if(wmiArrayResult.vt == (VT_BSTR|VT_ARRAY))
			{
				LONG lBound = 0;
				LONG uBound = 0;
				SafeArrayGetLBound(wmiArrayResult.parray, 1, &lBound);
				SafeArrayGetUBound(wmiArrayResult.parray, 1, &uBound);
				for(LONG j = lBound; j <= uBound; j++)
				{
					results++;
					BSTR var;
					SafeArrayGetElement(wmiArrayResult.parray, &j, &var);
					pushstring(_bstr_t(var));
					extra->ExecuteCodeSegment(CallbackAddress - 1, 0);
					if (extra->exec_flags->exec_error){}
				}
			}
			_stprintf_s(wmiResult, TEXT("Number of results: %i"), results);
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterDefaultIPGateways	  					  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Find all DNS Servers for the given	Networkadapter.		  * 
*					 All results are returned in a string (space seperated).  *	
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterDNSServers(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	TCHAR szBuf[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		if(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("DNSServerSearchOrder"), _where) == true)
		{
			if(wmiArrayResult.vt == (VT_BSTR|VT_ARRAY))
			{
				LONG lBound = 0;
				LONG uBound = 0;
				SafeArrayGetLBound(wmiArrayResult.parray, 1, &lBound);
				SafeArrayGetUBound(wmiArrayResult.parray, 1, &uBound);
				for(LONG j = lBound; j <= uBound; j++)
				{
					BSTR var;
					SafeArrayGetElement(wmiArrayResult.parray, &j, &var);
					_bstr_t _var = _bstr_t(var);
					if(j == lBound)_tcscpy_s(wmiResult,(TCHAR*)_var);
					else _stprintf_s(wmiResult, TEXT("%s %s"), wmiResult, (TCHAR*)_var);
					BSTR i = 0;
				}
			}
			StackResult(true);
		}
		else
		{
			StackResult(false);
		}
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterPrimaryWINSServer						  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Return the Primary WINS Server of the given			  *
*					 Networkadapter.									 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterPrimaryWINSServer(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("WINSPrimaryServer"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterSecondaryWINSServer					  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Return the Secondary WINS Server of the given			  *
*					 Networkadapter.									 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterSecondaryWINSServer(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		StackResult(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("WINSSecondaryServer"), _where));
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterDHCPLeaseObtained						  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Return the date and time when DHCP was obtained for the  *
*					 given Networkadapter.								 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterDHCPLeaseObtained(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		if(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("DHCPLeaseObtained"), _where) == true)
		{
			_stprintf_s(wmiResult, TEXT("%c%c.%c%c.%c%c%c%c %c%c:%c%c:%c%c"), wmiResult[6], wmiResult[7], wmiResult[4], wmiResult[5], wmiResult[0], wmiResult[1], wmiResult[2], wmiResult[3], wmiResult[8], wmiResult[9], wmiResult[10], wmiResult[11], wmiResult[12], wmiResult[13]);
			StackResult(true);
		}
		else StackResult(false);
	}
}
/******************************************************************************
* EXTERNAL FUNCTION: GetNetworkAdapterDHCPLeaseExpires						  *
*		 PARAMETERS: NETWORKADAPTERID										  *
*			PURPOSE: Return the date and time when DHCP will expire for the   *
*					 given Networkadapter.								 	  *
******************************************************************************/
extern "C"
void __declspec(dllexport) GetNetworkAdapterDHCPLeaseExpires(HWND hwndParent, int string_size, TCHAR *variables, stack_t **stacktop, extra_parameters *extra)
{
	TCHAR szNetworkAdapterID[MAX_STRLEN]={0};
	EXDLL_INIT();
	extra->RegisterPluginCallback(g_hInstance, PluginCallback);
	{
		popstring(szNetworkAdapterID);
		_bstr_t _where = TEXT("Index = ") + _bstr_t(szNetworkAdapterID);
		if(wmiRequest(TEXT("CIMV2"), TEXT("Win32_NetworkAdapterConfiguration"), TEXT("DHCPLeaseExpires"), _where) == true)
		{
			_stprintf_s(wmiResult, TEXT("%c%c.%c%c.%c%c%c%c %c%c:%c%c:%c%c"), wmiResult[6], wmiResult[7], wmiResult[4], wmiResult[5], wmiResult[0], wmiResult[1], wmiResult[2], wmiResult[3], wmiResult[8], wmiResult[9], wmiResult[10], wmiResult[11], wmiResult[12], wmiResult[13]);
			StackResult(true);
		}
		else StackResult(false);
	}
}
/******************************************************************************
* FUNCTION:	DllMain															  *
*  PURPOSE:	Entry point.													  *
******************************************************************************/
BOOL WINAPI DllMain(HINSTANCE hInst, ULONG ul_reason_for_call, LPVOID lpReserved)
{
	g_hInstance=hInst;

	return TRUE;
}

