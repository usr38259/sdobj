
#define INITGUID

#include <stdio.h>
#include <conio.h>
#include <objbase.h>

#include "sdobj.h"

HMODULE hModuleHandle;

BOOL bConsoleAllocd;

const char szProgIDScriptDispatch [] = "Scripting.{8C6E39B0-BC52-43C0-B3E2-A5F83FAEB897}";
const char szCLSID [] = "CLSID";
const char szInprocServer32 [] = "InprocServer32";

BOOL WINAPI DllMain (HINSTANCE hinstDll, DWORD fdwReason, LPVOID lpvReserved)
{
	if (fdwReason == DLL_PROCESS_ATTACH) {
		hModuleHandle = hinstDll;
		DisableThreadLibraryCalls (hinstDll);
		bConsoleAllocd = AllocConsole ();
	} else if (fdwReason == DLL_PROCESS_DETACH) {
		if (bConsoleAllocd) {
			FreeConsole ();
			bConsoleAllocd = FALSE;
		}
	}
	return TRUE;
}

STDAPI DllGetClassObject (REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	HRESULT hr;
	CoScriptDispatchFactory* lpcsdf;
	LPOLESTR psz1, lpsz2;

	StringFromCLSID (rclsid, &psz1);
	StringFromIID (riid, &lpsz2);
	cprintf ("DllGetClassObject:\n\tREFCLSID: %S\n\tREFIID: %S\n",
		psz1, lpsz2);
	CoTaskMemFree (psz1);
	CoTaskMemFree (lpsz2);

	if (rclsid != CLSID_CoScriptDispatch)
		return CLASS_E_CLASSNOTAVAILABLE;
	lpcsdf = new CoScriptDispatchFactory;
	hr = lpcsdf->QueryInterface (riid, ppv);
	if (FAILED(hr))
		delete lpcsdf;
	return hr;
}

STDAPI DllCanUnloadNow ()
{
	cprintf ("sdobj::DllCanUnloadNow()\n");
	if (g_objCount == 0 && g_lockCount == 0)
		return S_OK;
	return S_FALSE;
}

HRESULT HResultFromWin32 (LONG r);

STDAPI DllRegisterServer (void)
{
	HKEY hk, hsubk, hsubk2;
	LONG r;
	HRESULT hr;
	LPOLESTR strclsiid;
	DWORD dwLength;
	char strModulePath [MAX_PATH];

	dwLength = GetModuleFileName (hModuleHandle, strModulePath, sizeof (strModulePath));
	if (dwLength == sizeof (strModulePath))
		return ERROR_INSUFFICIENT_BUFFER;
	else if (dwLength == 0)
		return GetLastError ();

	r = RegCreateKeyEx (HKEY_CLASSES_ROOT, szProgIDScriptDispatch, 0, NULL, 0, KEY_ALL_ACCESS, NULL, &hk, NULL);
	if (r) return HResultFromWin32 (r);
	r = RegCreateKeyEx (hk, szCLSID, 0, NULL, 0, KEY_ALL_ACCESS, NULL, &hsubk, NULL);
	if (r) {
		RegCloseKey (hk);
		return HResultFromWin32 (r);
	}
	hr = StringFromCLSID (CLSID_CoScriptDispatch, &strclsiid);
	if (FAILED(hr)) {
		RegCloseKey (hsubk);
		RegCloseKey (hk);
		return hr;
	}
	r = RegSetValueExW (hsubk, NULL, 0, REG_SZ, (const BYTE*)strclsiid, (wcslen (strclsiid) + 1) * sizeof (wchar_t));
	if (r) {
		CoTaskMemFree (strclsiid);
		RegCloseKey (hsubk);
		RegCloseKey (hk);
		return HResultFromWin32 (r);
	}
	RegCloseKey (hsubk);
	RegCloseKey (hk);
	r = RegOpenKeyEx (HKEY_CLASSES_ROOT, szCLSID, 0, KEY_ALL_ACCESS, &hk);
	if (r) {
		CoTaskMemFree (strclsiid);
		return HResultFromWin32 (r);
	}
	r = RegCreateKeyExW (hk, strclsiid, 0, NULL, 0, KEY_ALL_ACCESS, NULL, &hsubk, NULL);
	if (r) {
		RegCloseKey (hk);
		CoTaskMemFree (strclsiid);
		return HResultFromWin32 (r);
	}
	r = RegCreateKeyEx (hsubk, szInprocServer32, 0, NULL, 0, KEY_ALL_ACCESS, NULL, &hsubk2, NULL);
	if (r) {
		RegCloseKey (hsubk);
		RegCloseKey (hk);
		CoTaskMemFree (strclsiid);
		return HResultFromWin32 (r);
	}
	r = RegSetValueEx (hsubk2, NULL, 0, REG_SZ, (const BYTE*)strModulePath, strlen (strModulePath) + 1);
	if (r) {
		RegCloseKey (hsubk2);
		RegCloseKey (hsubk);
		RegCloseKey (hk);
		CoTaskMemFree (strclsiid);
		return HResultFromWin32 (r);
	}
	return S_OK;
}

STDAPI DllUnregisterServer (void)
{
	HKEY hk, hsubk;
	LONG r;
	HRESULT hr;
	LPOLESTR strclsiid;

	hr = StringFromCLSID (CLSID_CoScriptDispatch, &strclsiid);
	if (FAILED(hr)) return hr;
	r = RegOpenKeyEx (HKEY_CLASSES_ROOT, szProgIDScriptDispatch, 0, KEY_ALL_ACCESS, &hsubk);
	if (!r) {
		RegDeleteKey (hsubk, szCLSID);
		RegCloseKey (hsubk);
		RegDeleteKey (HKEY_CLASSES_ROOT, szProgIDScriptDispatch);
	}
	r = RegOpenKeyEx (HKEY_CLASSES_ROOT, szCLSID, 0, KEY_ALL_ACCESS, &hk);
	if (!r) {
		r = RegOpenKeyExW (hk, strclsiid, 0, KEY_ALL_ACCESS, &hsubk);
		if (!r) {
			RegDeleteKey (hsubk, szInprocServer32);
			RegCloseKey (hsubk);
			RegDeleteKeyW (hk, strclsiid);
			RegCloseKey (hk);
		}
	}
	CoTaskMemFree (strclsiid);
	return S_OK;
}

HRESULT HResultFromWin32 (LONG r)
{
	return HRESULT_FROM_WIN32 (r);
}
