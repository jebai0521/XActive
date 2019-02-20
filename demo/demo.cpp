// demo.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "EventSink.h"

#include <atlbase.h>
#include <atlconv.h>

#include <ctype.h>

#include <windows.h>
#include <ocidl.h>
#include <olectl.h>
#include <ole2.h>
#include <objidl.h>

HRESULT registEvent2(IDispatch* idisp, CEventSink* pEvnetSink)
{
	LPCONNECTIONPOINTCONTAINER	pCPCont = NULL;
	HRESULT						hr;
	
	hr = idisp->QueryInterface(IID_IConnectionPointContainer, (void**) &pCPCont);
	if (FAILED(hr))
	{
		printf("QueryInterface IID_IConnectionPointContainer Failed!\n");
		return hr;
	}

	printf("QueryInterface IID_IConnectionPointContainer Success!\n");


	IProvideClassInfo2 *pProvideClassInfo2;
    void *p;

	hr = idisp->QueryInterface(IID_IProvideClassInfo2,
                                           &p);
	if (FAILED(hr))
	{
		printf("QueryInterface IID_IProvideClassInfo2 Failed!\n");
		return hr;
	}
	printf("QueryInterface IID_IProvideClassInfo2 Success!\n");

	IID clsid;
	pProvideClassInfo2 = (IProvideClassInfo2*)p;
	hr = pProvideClassInfo2->GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID,
		&clsid);
	pProvideClassInfo2->Release();
	if (FAILED(hr)) 
	{
		printf("QueryInterface GetGUID Failed!\n");
		return hr;
    }
	printf("clsid ==> %08x-%04x-%04x-%02x-%02x_%02x-%02x-%02x-%02x-%02x-%02x\n", clsid.Data1, clsid.Data2, clsid.Data3, clsid.Data4[0], clsid.Data4[1], clsid.Data4[2], clsid.Data4[3], clsid.Data4[4], clsid.Data4[5], clsid.Data4[6], clsid.Data4[7]);

	pEvnetSink->m_iid = clsid;

	ITypeInfo* pTypeInfo;
	hr = idisp->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (FAILED(hr))
	{
		printf("GetTypeInfo Failed!\n");
		return hr;
	}
	printf("GetTypeInfo Success!\n");
	
    ITypeLib *pTypeLib;
    unsigned int index;
	hr = pTypeInfo->GetContainingTypeLib( &pTypeLib,
		&index);
	
	if (FAILED(hr))
	{
		printf("GetContainingTypeLib Failed!\n");
		return hr;
	}
	printf("GetContainingTypeLib Success!\n");
	
	

	hr = pTypeLib->GetTypeInfoOfGuid(clsid, &pEvnetSink->pTypeInfo);
	if (FAILED(hr))
	{
		printf("GetTypeInfoOfGuid Failed!\n");
		return hr;
	}
	printf("GetTypeInfoOfGuid Success!\n");

	LPCONNECTIONPOINT	m_pMyCP;
	hr = pCPCont->FindConnectionPoint(clsid, &m_pMyCP);
	pCPCont->Release();
	if (FAILED(hr))
	{
		printf("FindConnectionPoint Failed!\n");
		return 0;
	}
	printf("FindConnectionPoint Success!\n");
	
	DWORD				m_dwMyCookie;
	hr = m_pMyCP->Advise(pEvnetSink, &m_dwMyCookie);

	if (FAILED(hr))
	{
		printf("Advise Failed!\n");
		return hr;
	}
	printf("Advise Success!\n");

	return hr;
}

short testActionA(IDispatch* idisp)
{
	HRESULT hr;

	BSTR PropName[1];  
	DISPID PropertyID[1] = {0}; 
	
	PropName[0] = SysAllocString(L"ActionA"); 
	
	hr = idisp->GetIDsOfNames(IID_NULL, PropName, 1, LOCALE_USER_DEFAULT, PropertyID);
	
	if (FAILED(hr))
	{
		printf("CLSIDFromProgID Failed!\n");
		return 0;
	}
	
	printf("dspid ==> %d\n", PropertyID[0]);
	
	VARIANT Result;
	
	DISPPARAMS params = { NULL,  
		NULL,              // Dispatch identifiers of named arguments.   
		0,                 // Number of arguments.  
		0 };                // Number of named arguments.  
	
	
	EXCEPINFO excepinfo;
	hr = idisp->Invoke(
		PropertyID[0],
		IID_NULL, 
		LOCALE_USER_DEFAULT, 
		DISPATCH_METHOD,
		&params,
		&Result,
		&excepinfo,
		NULL);
	if (FAILED(hr))
	{
		printf("Invoke Failed!\n");
		return 0;
	}
	
	switch (Result.vt)
	{
	case VT_I1:
		break;
	case VT_I2:
		printf("Result ==> %d!\n", Result.iVal);
		break;
	default:
		break;
	}
	
	printf("Execute ActionB Success!\n");

	return 0;
}

short testActionB(IDispatch* idisp)
{
	HRESULT hr;
	
	BSTR PropName[1];  
	DISPID PropertyID[1] = {0}; 
	
	PropName[0] = SysAllocString(L"ActionB"); 
	
	hr = idisp->GetIDsOfNames(IID_NULL, PropName, 1, LOCALE_USER_DEFAULT, PropertyID);
	
	if (FAILED(hr))
	{
		printf("CLSIDFromProgID Failed!\n");
		return 0;
	}
	
	printf("dspid ==> %d\n", PropertyID[0]);
	
	VARIANT Result;
	
	DISPPARAMS params = { new VARIANTARG[1],  
		NULL,              // Dispatch identifiers of named arguments.   
		0,                 // Number of arguments.  
		0 };                // Number of named arguments.  

	params.cArgs = 1;
	params.cNamedArgs = 0;
	
	params.rgvarg[0].vt = VT_BSTR;
	
    params.rgvarg[0].bstrVal = SysAllocString(L"Jebai");
	
	EXCEPINFO excepinfo;
	hr = idisp->Invoke(
		PropertyID[0],
		IID_NULL, 
		LOCALE_USER_DEFAULT, 
		DISPATCH_METHOD,
		&params,
		&Result,
		&excepinfo,
		NULL);
	if (FAILED(hr))
	{
		printf("Invoke Failed!\n");
		return 0;
	}

	switch (Result.vt)
	{
	case VT_I1:
		break;
	case VT_I2:
		break;
	case VT_BSTR:{
		USES_CONVERSION;
		printf("Result ==> %s!\n", W2A(Result.bstrVal));
		break;
				 }
	default:
		break;
	}
	
	printf("Execute ActionA Success!\n");
	return 0;
}


int main(int argc, char* argv[])
{
	CoInitialize(NULL); 

	HRESULT hr;

	CLSID clsid;
	hr = CLSIDFromProgID(OLESTR("OLEA.OLEACtrl.1"), &clsid);

	if (FAILED(hr))
	{
		printf("CLSIDFromProgID Failed!\n");
		return 0;
	}

	printf("clsid ==> %08x-%04x-%04x-%02x-%02x_%02x-%02x-%02x-%02x-%02x-%02x\n", clsid.Data1, clsid.Data2, clsid.Data3, clsid.Data4[0], clsid.Data4[1], clsid.Data4[2], clsid.Data4[3], clsid.Data4[4], clsid.Data4[5], clsid.Data4[6], clsid.Data4[7]);

	IDispatch* idisp;

	// ´´½¨ÊµÀý
	hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
                IID_IDispatch, (LPVOID*)&idisp);
	if (FAILED(hr))
	{
		printf("CoCreateInstance Failed!\n");
		return 0;
	}
	printf("idisp ==> %d\n", idisp);

	CEventSink* pEvnetSink = new CEventSink();

	registEvent2(idisp, pEvnetSink);

	testActionA(idisp);

	testActionB(idisp);


	MSG msg;
    while(PeekMessage(&msg,NULL,0,0,PM_REMOVE)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
	
	idisp->Release();

	return 0;
}


