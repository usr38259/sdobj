
#ifndef __SDOBJ_H__
#define __SDOBJ_H__

#include <objbase.h>

class CoScriptDispatch : public IDispatch
{
public:
	CoScriptDispatch ();
	virtual ~CoScriptDispatch ();

	STDMETHODIMP QueryInterface (REFIID riid, void** pIface);
	STDMETHODIMP_(ULONG) AddRef ();
	STDMETHODIMP_(ULONG) Release ();

	STDMETHODIMP GetTypeInfoCount (unsigned int FAR* pctinfo);
	STDMETHODIMP GetTypeInfo (unsigned int iTInfo, LCID lcid, ITypeInfo FAR* FAR* ppTInfo);
	STDMETHODIMP GetIDsOfNames (REFIID riid, OLECHAR FAR* FAR* rgszNames, unsigned int cNames,
		LCID lcid, DISPID FAR* rgDispId);
	STDMETHODIMP Invoke (DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS FAR* pDispParams,
		VARIANT FAR* pVarResult, EXCEPINFO FAR* pExcepInfo, unsigned int FAR* puArgErr);
private:
	ULONG m_refCount;
};

extern ULONG g_objCount;

class CoScriptDispatchFactory : public IClassFactory
{
public:
	CoScriptDispatchFactory ();
	~CoScriptDispatchFactory ();

	STDMETHODIMP QueryInterface (REFIID riid, void** pIFace);
	STDMETHODIMP_(ULONG) AddRef ();
	STDMETHODIMP_(ULONG) Release ();

	STDMETHODIMP CreateInstance (LPUNKNOWN pUnk, REFIID riid, void **pIFace);
	STDMETHODIMP LockServer (BOOL fLock);
private:
	ULONG m_refCount;
};

extern ULONG g_lockCount;

// {8C6E39B0-BC52-43c0-B3E2-A5F83FAEB897}
DEFINE_GUID(CLSID_CoScriptDispatch, 
0x8c6e39b0, 0xbc52, 0x43c0, 0xb3, 0xe2, 0xa5, 0xf8, 0x3f, 0xae, 0xb8, 0x97);

#endif /* __SDOBJ_H__ */
