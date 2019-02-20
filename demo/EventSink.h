// EventSink.h: interface for the CEventSink class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_EVENTSINK_H__1B47B203_59E1_417B_A6E0_2F367C983E7A__INCLUDED_)
#define AFX_EVENTSINK_H__1B47B203_59E1_417B_A6E0_2F367C983E7A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <atlbase.h>
#include <atlconv.h>

#include <ctype.h>

#include <windows.h>
#include <ocidl.h>
#include <olectl.h>
#include <ole2.h>
#include <objidl.h>

class CEventSink : public IDispatch  
{
public:
	CEventSink();
	virtual ~CEventSink();

public:
    long m_cRef;
    IID m_iid;
    long  m_event_id;
    ITypeInfo *pTypeInfo;

public:

	virtual HRESULT STDMETHODCALLTYPE QueryInterface( 
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
	virtual ULONG STDMETHODCALLTYPE AddRef( void);
	
	virtual ULONG STDMETHODCALLTYPE Release( void);

	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount( 
		/* [out] */ UINT *pctinfo);
        
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfo( 
		/* [in] */ UINT iTInfo,
		/* [in] */ LCID lcid,
		/* [out] */ ITypeInfo **ppTInfo);
        
    virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames( 
		/* [in] */ REFIID riid,
		/* [size_is][in] */ LPOLESTR *rgszNames,
		/* [in] */ UINT cNames,
		/* [in] */ LCID lcid,
		/* [size_is][out] */ DISPID *rgDispId);
        
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Invoke( 
		/* [in] */ DISPID dispIdMember,
		/* [in] */ REFIID riid,
		/* [in] */ LCID lcid,
		/* [in] */ WORD wFlags,
		/* [out][in] */ DISPPARAMS *pDispParams,
		/* [out] */ VARIANT *pVarResult,
		/* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);

public:



};

#endif // !defined(AFX_EVENTSINK_H__1B47B203_59E1_417B_A6E0_2F367C983E7A__INCLUDED_)
