/////////////////////////////////////////////////////////////////////////////////
// * FB2K Component: COM Automation and ActiveX Interface                    * //
// * Description:    MyCOM Impl Header File                                  * //
// * Author:         TT                                                      * //
// * Website:        https://github.com/TT-ReBORN/Georgia-ReBORN             * //
// * Version:        0.1                                                     * //
// * Dev. started:   12-12-2024                                              * //
// * Last change:    22-12-2024                                              * //
/////////////////////////////////////////////////////////////////////////////////


#pragma once
#include <Windows.h>
#include <comdef.h>
#include <atlbase.h>


//////////////
// * GUID * //
//////////////
const IID IID_IMyCOM = { 0xB2C3D4E5, 0xF678, 0x40A1, { 0x82, 0x34, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x01 } };
const IID LIBID_MyCOMLib = { 0xA1B2C3D4, 0xE5F6, 0x4780, { 0x92, 0x34, 0x56, 0x78, 0x9A, 0xBC, 0xDE, 0xF1 } };
const CLSID CLSID_MyCOM = { 0xC3D4E5F6, 0x4780, 0x91B2, { 0x34, 0x56, 0x78, 0x9A, 0xBC, 0xDE, 0xFF, 0x02 } };
const LPCOLESTR ProgID_MyCOM = OLESTR("MyCOM");


/////////////////////
// * MyCOM CLASS * //
/////////////////////
class MyCOM : public IDispatch {
public:
	// IUnknown Methods
	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override;
	STDMETHODIMP_(ULONG) Release() override;

	// IDispatch Methods
	STDMETHODIMP GetTypeInfoCount(UINT* pctinfo) override;
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override;
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) override;
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) override;

	// Custom methods
	STDMETHOD(PrintMessage)();

	// Class methods
	MyCOM() : refCount(1) {}
	~MyCOM() = default;

	HRESULT InitTypeInfo();

private:
	LONG refCount;
	CComPtr<ITypeInfo> m_pTypeInfo;
};


///////////////////////////
// * MyCOMClassFactory * //
///////////////////////////
class MyCOMClassFactory : public IClassFactory {
public:
	// IUnknown Methods
	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override;
	STDMETHODIMP_(ULONG) Release() override;

	// IClassFactory Methods
	STDMETHODIMP CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override;
	STDMETHODIMP LockServer(BOOL fLock) override;

	// Class methods
	MyCOMClassFactory() : refCount(1) {}
	~MyCOMClassFactory() = default;

private:
	LONG refCount;
};

HRESULT RegisterMyCOM();
HRESULT UnregisterMyCOM();
