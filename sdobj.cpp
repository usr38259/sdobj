
#include <conio.h>
#include "sdobj.h"

ULONG g_objCount = 0;

CoScriptDispatch::CoScriptDispatch ()
{
	m_refCount = 0;
	g_objCount++;
}

CoScriptDispatch::~CoScriptDispatch ()
{
	g_objCount--;
}

STDMETHODIMP CoScriptDispatch::QueryInterface (REFIID riid, void** ppv)
{
	LPOLESTR lpsz;
	StringFromIID (riid, &lpsz);
	cprintf ("CoScriptDispatch::QueryInterface (%S)\n", lpsz);
	CoTaskMemFree (lpsz);
	if (ppv == NULL)
		return E_POINTER;
	*ppv = NULL;

	if (riid == IID_IUnknown)
		*ppv = (IUnknown*)this;
	else if (riid == IID_IDispatch)
		*ppv = (IDispatch*)this;
	if (*ppv)
	{
		((IUnknown*)(*ppv))->AddRef ();
		return S_OK;
	}
	cputs ("\tE_NOTIMPL\n");
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CoScriptDispatch::AddRef ()
{
	return ++m_refCount;	// TODO: InterlockedIncrement ()
}

STDMETHODIMP_(ULONG) CoScriptDispatch::Release ()
{
	if (--m_refCount == 0)
	{
		cputs ("CoScriptDispatch::Release() -> No more objects\n");
		delete this;
		return 0;
	}
	return m_refCount;
}

STDMETHODIMP CoScriptDispatch::GetTypeInfoCount (unsigned int FAR* pctinfo)
{
	_cputs ("CoScriptDispatch::GetTypeInfoCount()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CoScriptDispatch::GetTypeInfo (unsigned int iTInfo, LCID lcid, ITypeInfo FAR* FAR* ppTInfo)
{
	_cputs ("CoScriptDispatch::GetTypeInfoCount()\n");
	return E_NOTIMPL;
}

STDMETHODIMP CoScriptDispatch::GetIDsOfNames (REFIID riid, OLECHAR FAR* FAR* rgszNames, unsigned int cNames,
	LCID lcid, DISPID FAR* rgDispId)
{
	cprintf ("CoScriptDispatch::GetIDsOfNames(%u):\n", cNames);
	HRESULT hr = S_OK;
	unsigned int i;
	for (i = 0; i < cNames; i++)
	{
		cprintf ("\t#%-2u %10S\n", i, rgszNames [i]);
		if (wcsicmp (rgszNames [i], L"Invoke") == 0)
			rgDispId [i] = 1;
		else {
			rgDispId [i] = DISPID_UNKNOWN;
			hr = DISP_E_UNKNOWNNAME;
		}
	}
	return hr;
}

STDMETHODIMP CoScriptDispatch::Invoke (DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS FAR* pDispParams,
	VARIANT FAR* pVarResult, EXCEPINFO FAR* pExcepInfo, unsigned int FAR* puArgErr)
{
	cprintf ("CoScriptDispatch::Invoke(dispIdMember = %d, ", dispIdMember);
	if (wFlags & DISPATCH_METHOD) cprintf ("DISPATCH_METHOD | ");
	if (wFlags & DISPATCH_PROPERTYGET) cprintf ("DISPATCH_PROPERTYGET | ");
	if (wFlags & DISPATCH_PROPERTYPUT) cprintf ("DISPATCH_PROPERTYPUT | ");
	if (wFlags & DISPATCH_PROPERTYPUTREF) cprintf ("DISPATCH_PROPERTYPUTREF");
	cprintf (")\n");
	cprintf ("\tcArgs = %u, cNamedArgs = %u\n", pDispParams->cArgs, pDispParams->cNamedArgs);
	switch (dispIdMember)
	{
	case 1:
		break;
	}
	return S_OK;
}

ULONG g_lockCount = 0;

CoScriptDispatchFactory::CoScriptDispatchFactory ()
{
	m_refCount = 0;
	g_objCount++;
}

CoScriptDispatchFactory::~CoScriptDispatchFactory ()
{
	g_objCount--;
}

STDMETHODIMP CoScriptDispatchFactory::QueryInterface (REFIID riid, void** pIFace)
{
	if (pIFace == NULL)
		return E_POINTER;
	*pIFace = NULL;

	if (riid == IID_IUnknown)
		*pIFace = (IUnknown*)this;
	else if (riid == IID_IClassFactory)
		*pIFace = (IClassFactory*)this;

	if (*pIFace)
	{
		((IUnknown*)(*pIFace))->AddRef ();
		return S_OK;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CoScriptDispatchFactory::AddRef ()
{
	return ++m_refCount;
}

STDMETHODIMP_(ULONG) CoScriptDispatchFactory::Release ()
{
	if (--m_refCount == 0)
	{
		delete this;
		return 0;
	}
	return m_refCount;
}

STDMETHODIMP CoScriptDispatchFactory::CreateInstance (LPUNKNOWN pUnk, REFIID riid, void **pIFace)
{
	if (pUnk != NULL)
		return CLASS_E_NOAGGREGATION;
	HRESULT hr;
	CoScriptDispatch* pcsd;
	pcsd = new CoScriptDispatch;
	hr = pcsd->QueryInterface (riid, pIFace);
	if (FAILED(hr))
		delete pcsd;
	return hr;
}

STDMETHODIMP CoScriptDispatchFactory::LockServer (BOOL fLock)
{
	if (fLock)
		++g_lockCount;
	else
		--g_lockCount;
	cprintf ("CoScriptDispatchFactory::LockServer (%s). Lock count = %u\n",
		fLock ? "1" : "0", g_lockCount);
	return S_OK;
}
