/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    MyCOM Impl Source File                                  * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/TT-ReBORN/Georgia-ReBORN             * //
// * Version:        0.1                                                     * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    22-12-2024                                              * //
/////////////////////////////////////////////////////////////////////////////////


#include "stdafx.h"
#include "MyCOM.h"
#include "MinHook.h"
#include <map>


///////////////////
// * VARIABLES * //
///////////////////
DWORD g_dwRegister = 0;
CComPtr<ITypeLib> g_typelib;
std::wstring g_regMethod = L"RegFree"; // "RegFree" or "RegEntry"


/////////////////////
// * MyCOM CLASS * //
/////////////////////
STDMETHODIMP MyCOM::QueryInterface(REFIID riid, void** ppvObject) {
	if (riid == IID_IUnknown || riid == IID_IDispatch || riid == IID_IMyCOM) {
		*ppvObject = static_cast<IDispatch*>(this);
		AddRef();
		return S_OK;
	}
	*ppvObject = nullptr;
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) MyCOM::AddRef() {
	return InterlockedIncrement(&refCount);
}

STDMETHODIMP_(ULONG) MyCOM::Release() {
	LONG count = InterlockedDecrement(&refCount);
	if (count == 0) {
		delete this;
	}
	return count;
}

HRESULT MyCOM::InitTypeInfo() {
	if (m_pTypeInfo == nullptr && g_typelib != nullptr) {
		HRESULT hr = g_typelib->GetTypeInfoOfGuid(IID_IMyCOM, &m_pTypeInfo);
		if (FAILED(hr)) {
			console::print("GetTypeInfoOfGuid failed with HRESULT:", hr);
			return hr;
		}
	}
	return S_OK;
}

STDMETHODIMP MyCOM::GetTypeInfoCount(UINT* pctinfo) {
	*pctinfo = (m_pTypeInfo != nullptr) ? 1 : 0;
	return S_OK;
}

STDMETHODIMP MyCOM::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) {
	if (iTInfo != 0) {
		return DISP_E_BADINDEX;
	}
	if (m_pTypeInfo == nullptr) {
		HRESULT hr = InitTypeInfo();
		if (FAILED(hr)) {
			console::print("InitTypeInfo failed with HRESULT:", hr);
			return hr;
		}
	}
	*ppTInfo = m_pTypeInfo;
	if (*ppTInfo) {
		(*ppTInfo)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP MyCOM::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) {
	if (m_pTypeInfo == nullptr) {
		return E_NOTIMPL;
	}
	return m_pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

STDMETHODIMP MyCOM::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) {
	if (m_pTypeInfo == nullptr) {
		return E_NOTIMPL;
	}
	return m_pTypeInfo->Invoke(static_cast<IDispatch*>(this), dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}

STDMETHODIMP MyCOM::PrintMessage() {
	MessageBox(nullptr, L"COM/ActiveX interface is working!", L"Message", MB_OK);
	return S_OK;
}


///////////////////////////
// * MyCOMClassFactory * //
///////////////////////////
STDMETHODIMP MyCOMClassFactory::QueryInterface(REFIID riid, void** ppvObject) {
	if (riid == IID_IUnknown || riid == IID_IClassFactory) {
		*ppvObject = static_cast<IClassFactory*>(this);
		AddRef();
		return S_OK;
	}
	*ppvObject = nullptr;
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) MyCOMClassFactory::AddRef() {
	return InterlockedIncrement(&refCount);
}

STDMETHODIMP_(ULONG) MyCOMClassFactory::Release() {
	LONG count = InterlockedDecrement(&refCount);
	if (count == 0) {
		delete this;
	}
	return count;
}

STDMETHODIMP MyCOMClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) {
	if (pUnkOuter != nullptr) {
		return CLASS_E_NOAGGREGATION;
	}

	CComPtr<MyCOM> pMyCOM = new MyCOM();
	return pMyCOM->QueryInterface(riid, ppvObject);
}

STDMETHODIMP MyCOMClassFactory::LockServer(BOOL fLock) {
	if (fLock) {
		InterlockedIncrement(&refCount);
	}
	else {
		InterlockedDecrement(&refCount);
	}
	return S_OK;
}


////////////////////////////////////////////
// * COM REGISTRATION-FREE WITH MINHOOK * //
////////////////////////////////////////////
typedef HRESULT(WINAPI* CLSIDFromProgID_t)(LPCOLESTR, LPCLSID);
CLSIDFromProgID_t Original_CLSIDFromProgID = nullptr;

HRESULT WINAPI MyCLSIDFromProgID(LPCOLESTR lpszProgID, LPCLSID lpclsid) {
	static std::map<std::wstring, CLSID> progidToClsidMap = {
		{ ProgID_MyCOM, CLSID_MyCOM }
	};

	auto it = progidToClsidMap.find(lpszProgID);
	if (it != progidToClsidMap.end()) {
		*lpclsid = it->second;
		return S_OK;
	}

	// Call the original function for other ProgIDs
	if (Original_CLSIDFromProgID != nullptr) {
		return Original_CLSIDFromProgID(lpszProgID, lpclsid);
	}

	return REGDB_E_CLASSNOTREG;
}

void HookCLSIDFromProgID() {
	if (MH_Initialize() == MH_OK &&
		MH_CreateHook(&CLSIDFromProgID, &MyCLSIDFromProgID, reinterpret_cast<LPVOID*>(&Original_CLSIDFromProgID)) == MH_OK &&
		MH_EnableHook(&CLSIDFromProgID) == MH_OK) {
		return;
	}
}


/////////////////////////////////////////
// * COM REGISTRATION WITH REG ENTRY * //
/////////////////////////////////////////
HRESULT SetRegistry(HKEY rootKey, const wchar_t* subKey, const wchar_t* valueName, const wchar_t* value) {
	CRegKey hKey;
	LONG lResult = hKey.Create(rootKey, subKey);
	if (lResult != ERROR_SUCCESS) {
		console::print("Failed to create/open key. Error: ", lResult);
		return HRESULT_FROM_WIN32(lResult);
	}

	lResult = hKey.SetStringValue(valueName, value);
	if (lResult != ERROR_SUCCESS) {
		console::print("Failed to set key value. Error: ", lResult);
		return HRESULT_FROM_WIN32(lResult);
	}

	return S_OK;
}

HRESULT RegisterCLSID() {
	HRESULT hr;
	LPOLESTR clsidString = nullptr;

	hr = StringFromCLSID(CLSID_MyCOM, &clsidString);
	hr = SetRegistry(HKEY_CLASSES_ROOT, L"MyCOM\\CLSID", nullptr, clsidString);
	CoTaskMemFree(clsidString);

	if (FAILED(hr)) return hr;

	return S_OK;
}

HRESULT UnregisterCLSID() {
	RegDeleteTree(HKEY_CLASSES_ROOT, L"MyCOM");
	return S_OK;
}


//////////////////////////
// * COM REGISTRATION * //
//////////////////////////
HRESULT RegisterMyCOM() {
	HRESULT hr;
	LPOLESTR clsidString = nullptr;

	hr = CoInitialize(nullptr);
	if (FAILED(hr)) {
		MessageBox(nullptr, L"Failed to initialize COM library", L"Error", MB_OK | MB_ICONERROR);
		return hr;
	}

	CComPtr<MyCOMClassFactory> pClassFactory = new MyCOMClassFactory();
	if (!pClassFactory) {
		MessageBox(nullptr, L"Failed to create MyCOMClassFactory.", L"Error", MB_OK | MB_ICONERROR);
		return E_OUTOFMEMORY;
	}

	hr = CoRegisterClassObject(CLSID_MyCOM, pClassFactory, CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE, &g_dwRegister);
	if (FAILED(hr)) {
		MessageBox(nullptr, L"CoRegisterClassObject failed", L"Error", MB_OK | MB_ICONERROR);
		return hr;
	}

	if (g_regMethod == L"RegEntry") {
		hr = RegisterCLSID();
		if (FAILED(hr)) {
			MessageBox(nullptr, L"Failed to register CLSID", L"Error", MB_OK | MB_ICONERROR);
			CoRevokeClassObject(g_dwRegister);
			return hr;
		}
	}
	else {
		HookCLSIDFromProgID();
	}

	return S_OK;
}

HRESULT UnregisterMyCOM() {
	if (g_dwRegister != 0) {
		HRESULT hr = CoRevokeClassObject(g_dwRegister);
		if (FAILED(hr)) {
			console::print("CoRevokeClassObject failed with HRESULT:", hr);
			return hr;
		}
		else {
			if (g_regMethod == L"RegEntry") {
				UnregisterCLSID(); // COM UN-REGISTRATION WITH REG ENTRY METHOD
			}
		}
		g_dwRegister = 0;
		CoUninitialize();
	}
	return S_OK;
}


////////////////////////
// * DLL MAIN ENTRY * //
////////////////////////
BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
	switch (ul_reason_for_call) {
		case DLL_PROCESS_ATTACH: {
			DisableThreadLibraryCalls(hModule);
			(VOID)OleInitialize(nullptr);

			wchar_t module_path[MAX_PATH] = {};
			DWORD len = GetModuleFileNameW(hModule, module_path, MAX_PATH);
			if (len == 0 || len >= MAX_PATH) {
				console::print("Failed to retrieve module path.");
				return FALSE;
			}

			HRESULT hr = LoadTypeLibEx(module_path, REGKIND_NONE, &g_typelib);
			if (FAILED(hr)) {
				console::print("LoadTypeLibEx failed with HRESULT:", hr);
				return FALSE;
			}
		}
		break;

		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}
