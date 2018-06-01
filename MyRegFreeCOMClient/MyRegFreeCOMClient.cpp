// MyRegFreeCOMClient.cpp : アプリケーションのエントリ ポイントを定義します。
//

#include "stdafx.h"

#include <Windows.h>
#include <comdef.h>

#include "..\MyRegFreeCOMSrv\MyRegFreeCOMSrv_h.h"
#include "..\MyRegFreeCOMSrv\MyRegFreeCOMSrv_i.c"

// ISimpleCOMExampleSrvのスマートポインタ(ISimpleCOMExampleSrvPtr)の定義
_COM_SMARTPTR_TYPEDEF(IMyRegFreeCOMSrv, __uuidof(IMyRegFreeCOMSrv));


/**
 * イベントシンク
 */
class MyRegFreeCOMSrvEventsSink
	: public _IMyRegFreeCOMSrvEvents
{
public:
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppvObject)
	{
		if (riid == IID_IUnknown ||
			riid == IID_IDispatch ||
			riid == DIID__IMyRegFreeCOMSrvEvents) {
			*ppvObject = static_cast<_IMyRegFreeCOMSrvEvents *>(this);
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef(void)
	{
		return 2;
	}

	virtual ULONG STDMETHODCALLTYPE Release(void)
	{
		return 1;
	}

	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo)
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
	{
		return E_NOTIMPL;
	}

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE Invoke(
		DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
		DISPPARAMS *pDispParams, VARIANT *pVarResult,
		EXCEPINFO *pExcepInfo, UINT *puArgErr)
	{
		switch (dispIdMember) {
			case 1: {
				_variant_t arg1 = pDispParams->rgvarg[1]; // 引数は逆順
				_variant_t arg2 = pDispParams->rgvarg[0];
				arg1.ChangeType(VT_BSTR);
				int ret = MessageBoxW(NULL, arg1.bstrVal, L"NamePropertyChanging", MB_ICONQUESTION | MB_YESNO);
				if (ret != IDYES) {
					if (arg2.vt & VT_BYREF) {
						*arg2.pboolVal = VARIANT_TRUE;
					}
				}
				return S_OK;
			}

			case 2: {
				_variant_t arg1 = pDispParams->rgvarg[0];
				arg1.ChangeType(VT_BSTR);
				MessageBoxW(NULL, arg1.bstrVal, L"NamePropertyChanged", MB_ICONINFORMATION| MB_OK);
				return S_OK;
			}
		}
		return DISP_E_MEMBERNOTFOUND;
	}

};

// 参照カウンタを持っていないため。
// (またDotNETで実装したCOMのオブジェクトにイベントシンクを参照させた場合は
// Releaseのタイミングが不定になるようなのでプログラム終了まで生かしておく？)
MyRegFreeCOMSrvEventsSink sink;

#define TEST_NATIVE_COM
#define TEST_DOTNET_COM

int main()
{
	CoInitialize(nullptr);

#ifdef TEST_NATIVE_COM
	// C++ネイティブ版 COMオブジェクトの呼び出し
	{
		IMyRegFreeCOMSrvPtr pSrv;
		HRESULT hr = pSrv.CreateInstance("MyRegFreeCOMSrv.1");
		if (SUCCEEDED(hr)) {
			IConnectionPointContainerPtr pCPC(pSrv);
			if (pCPC) {
				IConnectionPointPtr pCP;
				hr = pCPC->FindConnectionPoint(DIID__IMyRegFreeCOMSrvEvents, &pCP);
				if (SUCCEEDED(hr)) {
					MyRegFreeCOMSrvEventsSink *pSinkImpl = new MyRegFreeCOMSrvEventsSink();
					IUnknownPtr pSink(pSinkImpl);

					DWORD cookie = 0;
					hr = pCP->Advise(pSink, &cookie);
					if (SUCCEEDED(hr)) {
						_bstr_t name(L"PiyoPiyo");
						pSrv->put_Name(name);
						pSrv->ShowHello();

						pCP->Unadvise(cookie);
					}
				}
			}
		}
	}
#endif

#ifdef TEST_DOTNET_COM
	// Dotnet版 COMオブジェクトの呼び出しテスト
	{
		CLSID clsid = { 0 };
		HRESULT hr = CLSIDFromProgID(L"MyRegFreeCOMDotnetSrv.1", &clsid);
		if (SUCCEEDED(hr)) {
			IDispatchPtr pDisp;
			hr = pDisp.CreateInstance(clsid);

			IConnectionPointContainerPtr pCPC(pDisp);
			if (pCPC) {
				IConnectionPointPtr pCP;
				hr = pCPC->FindConnectionPoint(DIID__IMyRegFreeCOMSrvEvents, &pCP);
				if (SUCCEEDED(hr)) {
					IUnknownPtr pSink(&sink);
					DWORD cookie = 0;
					hr = pCP->Advise(pSink, &cookie);
					if (SUCCEEDED(hr)) {

						// プロパティの書き込みテスト
						{
							LPCWSTR name = L"Name";
							DISPID dispid = 0;
							hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
							if (SUCCEEDED(hr)) {
								_variant_t varParams1[1];
								varParams1[0] = L"PiyoPiyo2";
								DISPPARAMS params = { varParams1, nullptr, 1, 0 }; // 引数の積み方は逆順
								_variant_t ret;
								hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, &ret, nullptr, nullptr);
							}
						}

						// メソッドの呼び出しテスト
						{
							LPCWSTR name = L"ShowHello";
							DISPID dispid = 0;
							hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
							if (SUCCEEDED(hr)) {
								DISPPARAMS params = { nullptr , nullptr, 0, 0 };
								_variant_t ret;
								hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &ret, nullptr, nullptr);
							}
						}

						hr = pCP->Unadvise(cookie); // DotNETのCOM実装だと、このタイミングでイベントシンクのReleaseが呼ばれない？
					}
				}
			}
		}
	}
#endif
	CoUninitialize();
    return 0;
}

