// EventSink.cpp: implementation of the CEventSink class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "EventSink.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CEventSink::CEventSink()
{
	printf("--->CEventSink Construct!\n");
	m_cRef = 0;
	pTypeInfo = NULL;
}

CEventSink::~CEventSink()
{
	printf("--->CEventSink Destruct!\n");
	pTypeInfo->Release();
}

HRESULT STDMETHODCALLTYPE CEventSink::QueryInterface( /* [in] */ REFIID riid, /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject )
{
	printf("--->CEventSink QueryInterface!\n");
    if (IsEqualIID(riid, IID_IUnknown) ||
        IsEqualIID(riid, IID_IDispatch) ||
        IsEqualIID(riid, this->m_iid)) {
        *ppvObject = this;
    }
    else {
        *ppvObject = NULL;
        return E_NOINTERFACE;
    }
	this->AddRef();
    return NOERROR;
}

ULONG STDMETHODCALLTYPE CEventSink::AddRef( void )
{
	printf("--->CEventSink AddRef!\n");
	return ++m_cRef;
}

ULONG STDMETHODCALLTYPE CEventSink::Release( void )
{
	printf("--->CEventSink Release!\n");
    ULONG ulRefCount = --m_cRef;
    if (0 == m_cRef)
    {
        delete this;
    }
    return ulRefCount;
}

HRESULT STDMETHODCALLTYPE CEventSink::GetTypeInfoCount( /* [out] */ UINT *pctinfo )
{
	printf("--->CEventSink GetTypeInfoCount!\n");
	*pctinfo = 0;
    return NOERROR;
}

HRESULT STDMETHODCALLTYPE CEventSink::GetTypeInfo( /* [in] */ UINT iTInfo, /* [in] */ LCID lcid, /* [out] */ ITypeInfo **ppTInfo )
{
	printf("--->CEventSink GetTypeInfo!\n");
    *ppTInfo = NULL;
    return DISP_E_BADINDEX;
}

HRESULT STDMETHODCALLTYPE CEventSink::GetIDsOfNames( /* [in] */ REFIID riid, /* [size_is][in] */ LPOLESTR *rgszNames, /* [in] */ UINT cNames, /* [in] */ LCID lcid, /* [size_is][out] */ DISPID *rgDispId )
{
	printf("--->CEventSink GetIDsOfNames!\n");
    if (pTypeInfo) {
        return pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
    }
    return DISP_E_UNKNOWNNAME;
}

/* [local] */ HRESULT STDMETHODCALLTYPE CEventSink::Invoke( /* [in] */ DISPID dispIdMember, /* [in] */ REFIID riid, /* [in] */ LCID lcid, /* [in] */ WORD wFlags, /* [out][in] */ DISPPARAMS *pDispParams, /* [out] */ VARIANT *pVarResult, /* [out] */ EXCEPINFO *pExcepInfo, /* [out] */ UINT *puArgErr )
{
	printf("--->CEventSink Invoke!\n");
	printf("Invoke Arrived!!!!\n");

	HRESULT hr;
    BSTR bstr;
    unsigned int count;
	hr = pTypeInfo->GetNames(dispIdMember, &bstr, 1, &count);

	if (FAILED(hr))
	{
		printf("--->CEventSink Invoke GetNames Failed!\n");
		return E_NOINTERFACE;
	}
	
	USES_CONVERSION;
	printf("--->CEventSink Invoke EventName %s!\n", W2A(bstr));

	if (pDispParams != NULL)
	{
		printf("--->CEventSink Invoke pDispParams cArgs %d!\n", pDispParams->cArgs);
		printf("--->CEventSink Invoke pDispParams cNamedArgs %d!\n", pDispParams->cNamedArgs);

		for (UINT i = 0; i < pDispParams->cArgs; i++)
		{
			VARIANT Result = pDispParams->rgvarg[0];
			switch (Result.vt)
			{
			case VT_I1:
				printf("Result ==> %d!\n", Result.iVal);
				break;
			case VT_I2:
				printf("Result ==> %d!\n", Result.iVal);
				break;
			case VT_BSTR:{
				USES_CONVERSION;
				printf("Result ==> %s!\n", W2A(Result.bstrVal));
				break;
			}
			default:
				break;
			}
		}

		for (UINT j = 0; j < pDispParams->cNamedArgs; j++)
		{
			VARIANT Result = pDispParams->rgvarg[0];
			switch (Result.vt)
			{
			case VT_I1:
				printf("Result ==> %d!\n", Result.iVal);
				break;
			case VT_I2:
				printf("Result ==> %d!\n", Result.iVal);
				break;
			case VT_BSTR:{
				USES_CONVERSION;
				printf("Result ==> %s!\n", W2A(Result.bstrVal));
				break;
			}
			default:
				break;
			}
		}
	}

	return NOERROR;
}
