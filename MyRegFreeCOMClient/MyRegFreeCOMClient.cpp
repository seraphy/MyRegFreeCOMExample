// MyRegFreeCOMClient.cpp : アプリケーションのエントリ ポイントを定義します。
//

#include "stdafx.h"

#include <Windows.h>
#include <comdef.h>

#include "..\MyRegFreeCOMSrv\MyRegFreeCOMSrv_h.h"
#include "..\MyRegFreeCOMSrv\MyRegFreeCOMSrv_i.c"

// ISimpleCOMExampleSrvのスマートポインタ(ISimpleCOMExampleSrvPtr)の定義
_COM_SMARTPTR_TYPEDEF(IMyRegFreeCOMSrv, __uuidof(IMyRegFreeCOMSrv));

int main()
{
	CoInitialize(nullptr);
	
	// C++ネイティブ版 COMオブジェクトの呼び出し
	{
		IMyRegFreeCOMSrvPtr pSrv;
		HRESULT hr = pSrv.CreateInstance("MyRegFreeCOMSrv.1");
		if (SUCCEEDED(hr)) {
			_bstr_t name(L"PiyoPiyo");
			pSrv->put_Name(name);
			pSrv->ShowHello();
		}
	}

	// Dotnet版 COMオブジェクトの呼び出しテスト
	{
		CLSID clsid = { 0 };
		HRESULT hr = CLSIDFromProgID(L"MyRegFreeCOMDotnetSrv.1", &clsid);
		if (SUCCEEDED(hr)) {
			IDispatchPtr pDisp;
			hr = pDisp.CreateInstance(clsid);

			// DotNETでIProvideClassInfo2の実装の確認
			// (明示的に実装しないと、インターフェイスは存在しないようだ？(なので接続できない。))
			if (SUCCEEDED(hr)) {
				IProvideClassInfo2Ptr pClassInfo(pDisp);
				if (pClassInfo != nullptr) {
					ITypeInfoPtr pTypeInfo;
					hr = pClassInfo->GetClassInfoW(&pTypeInfo);
					if (SUCCEEDED(hr)) {

					}
					GUID guid = { 0 };
					hr = pClassInfo->GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID, &guid);
					if (SUCCEEDED(hr)) {

					}
				}
			}

			// プロパティの書き込みテスト
			if (SUCCEEDED(hr)) {
				LPCWSTR name = L"Name";
				DISPID dispid = 0;
				hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*) &name, 1, LOCALE_USER_DEFAULT, &dispid);
				if (SUCCEEDED(hr)) {
					_variant_t varParams1[1];
					varParams1[0] = L"PiyoPiyo2";
					DISPPARAMS params = { varParams1, nullptr, 1, 0 }; // 引数の積み方は逆順
					_variant_t ret;
					hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, &ret, nullptr, nullptr);
				}
			}

			// メソッドの呼び出しテスト
			if (SUCCEEDED(hr)) {
				LPCWSTR name = L"ShowHello";
				DISPID dispid = 0;
				hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1, LOCALE_USER_DEFAULT, &dispid);
				if (SUCCEEDED(hr)) {
					DISPPARAMS params = { nullptr , nullptr, 0, 0 };
					_variant_t ret;
					hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &ret, nullptr, nullptr);
				}
			}
		}
	}
	CoUninitialize();
    return 0;
}

