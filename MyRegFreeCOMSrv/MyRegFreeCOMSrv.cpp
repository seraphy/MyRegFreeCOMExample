// MyRegFreeCOMSrv.cpp : DLL アプリケーション用にエクスポートされる関数を定義します。
//

#include "stdafx.h"

#include <comdef.h>
#include <map>

#include<vcclr.h>  

#include "MyRegFreeCOMSrv_h.h"

extern HMODULE g_hModule;

class MyRegFreeCOMSrv
	: public IMyRegFreeCOMSrv
	, public IConnectionPointContainer
	, public IConnectionPoint
	, public IProvideClassInfo2
{
private:
	static ITypeLibPtr m_pTypeLib;
	static ITypeInfoPtr m_pTypeInfo;

	ULONG m_refCount;

	_bstr_t m_bstrName;

	// クッキーがかぶらないように管理するクッキーの最終値
	ULONG m_lastCookie;

	// コネクションポイントのアドバイスしているイベントシンク
	std::map<ULONG, IDispatch*> m_adviseMap;

	/**
	* デストラクタ。
	* ここでstaticで参照しているタイプライブラリの参照カウントを減する。
	*/
	virtual ~MyRegFreeCOMSrv()
	{
		if (m_pTypeInfo) {
			m_pTypeInfo.Release();
		}
		if (m_pTypeLib) {
			m_pTypeLib.Release();
		}
		for (auto ite = m_adviseMap.begin(); ite != m_adviseMap.end(); ++ite) {
			ite->second->Release();
		}
		m_adviseMap.clear();
	}

public:
	MyRegFreeCOMSrv()
		: m_refCount(0)
	{
	}

	/**
	* タイプライブラリのロードまたはaddRefなど、
	* クラスの実行に事前に必要な初期化処理を行う。
	* (クラス生成失敗理由をHRESULTで返せるようにする)
	*/
	HRESULT Init()
	{
		HRESULT hr = S_OK;

		// このDLLのリソースとして埋め込まれているタイプライブラリをロードして
		// Dispatchの実装に利用する.
		if (m_pTypeLib == nullptr) {
			TCHAR szModulePath[MAX_PATH];
			GetModuleFileName(g_hModule, szModulePath, MAX_PATH);

			hr = LoadTypeLib(szModulePath, &m_pTypeLib);
			if (SUCCEEDED(hr)) {
				// IMyRegFreeCOMSrvのディスパッチインターフェイスのタイプライブラリ取得
				hr = m_pTypeLib->GetTypeInfoOfGuid(IID_IMyRegFreeCOMSrv, &m_pTypeInfo);
			}
		}
		else {
			m_pTypeLib.AddRef();
			if (m_pTypeInfo) {
				m_pTypeInfo.AddRef();
			}
		}
		return hr;
	}

#pragma region IUnknownの実装
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void **ppvObject)
	{
		if (riid == IID_IUnknown ||
			riid == IID_IDispatch ||
			riid == IID_IMyRegFreeCOMSrv) {
			*ppvObject = static_cast<IMyRegFreeCOMSrv *>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == IID_IConnectionPointContainer) {
			*ppvObject = static_cast<IConnectionPointContainer *>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == IID_IConnectionPoint) {
			*ppvObject = static_cast<IConnectionPoint *>(this);
			AddRef();
			return S_OK;
		}
		else if (riid == IID_IProvideClassInfo || riid == IID_IProvideClassInfo2) {
			*ppvObject = static_cast<IProvideClassInfo2 *>(this);
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	virtual ULONG __stdcall AddRef()
	{
		return InterlockedIncrement(&m_refCount);
	}

	virtual ULONG __stdcall Release()
	{
		if (InterlockedDecrement(&m_refCount) == 0) {
			delete this;
			return 0;
		}

		return m_refCount;
	}
#pragma endregion
#pragma region IDispatchの実装(型情報から既定の処理に転送するだけ)

	virtual HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo)
	{
		if (pctinfo == nullptr) {
			return E_POINTER;
		}
		*pctinfo = m_pTypeLib->GetTypeInfoCount();
		return S_OK;
	}

	virtual HRESULT __stdcall GetTypeInfo(
		UINT iTInfo,
		LCID lcid,
		ITypeInfo **ppTInfo)
	{
		return m_pTypeLib->GetTypeInfo(iTInfo, ppTInfo);
	}

	virtual HRESULT __stdcall GetIDsOfNames(
		REFIID riid,
		LPOLESTR *rgszNames,
		UINT cNames,
		LCID lcid,
		DISPID *rgDispId)
	{
		return DispGetIDsOfNames(m_pTypeInfo, rgszNames, cNames, rgDispId);
	}

	virtual HRESULT __stdcall Invoke(
		DISPID dispIdMember,
		REFIID riid,
		LCID lcid,
		WORD wFlags,
		DISPPARAMS *pDispParams,
		VARIANT *pVarResult,
		EXCEPINFO *pExcepInfo,
		UINT *puArgErr)
	{
		return DispInvoke(this, // 呼び出しを、このオブジェクトのメソッドに転送する
			m_pTypeInfo, dispIdMember, wFlags,
			pDispParams, pVarResult, pExcepInfo, puArgErr);
	}
#pragma endregion
#pragma region ConnectionPointContainer
	virtual HRESULT _stdcall EnumConnectionPoints(IEnumConnectionPoints **ppEnum)
	{
		return E_NOTIMPL; // クライアントはコンパイル時に既知のDIIDを知っているはずなので省略可。
	}

	virtual HRESULT _stdcall FindConnectionPoint(REFIID riid, IConnectionPoint **ppCP)
	{
		if (riid == DIID__IMyRegFreeCOMSrvEvents) {
			return QueryInterface(IID_IConnectionPoint, (void**)ppCP);
		}
		return E_NOINTERFACE;
	}
#pragma endregion
#pragma region ConnectionPoint
	virtual HRESULT __stdcall GetConnectionInterface(IID *pIID)
	{
		*pIID = DIID__IMyRegFreeCOMSrvEvents;
		return S_OK;
	}

	virtual HRESULT __stdcall GetConnectionPointContainer(IConnectionPointContainer **ppCPC)
	{
		return QueryInterface(IID_IConnectionPointContainer, (void**)ppCPC);
	}

	virtual HRESULT __stdcall Advise(IUnknown *pUnkSink, DWORD *pdwCookie)
	{
		if (!pdwCookie || !pUnkSink) {
			return E_POINTER;
		}
		IDispatch *pDisp = nullptr;
		HRESULT hr = pUnkSink->QueryInterface(IID_IDispatch, (void **)&pDisp);
		if (SUCCEEDED(hr)) {
			ULONG cookie = InterlockedIncrement(&m_lastCookie);
			*pdwCookie = cookie;
			m_adviseMap[cookie] = pDisp;
			return S_OK;
		}
		return hr;
	}

	virtual HRESULT __stdcall Unadvise(DWORD dwCookie)
	{
		auto ite = m_adviseMap.find(dwCookie);
		if (ite != m_adviseMap.end()) {
			IDispatch *pDisp = ite->second;
			pDisp->Release();
			m_adviseMap.erase(ite);
		}
		return S_OK;
	}

	virtual HRESULT __stdcall EnumConnections(IEnumConnections **ppEnum)
	{
		return E_NOTIMPL; // 列挙子はサポートしなくても可
	}

	/**
	* すべてのイベントシンクに対してイベントを送信するヘルパ
	*/
	HRESULT FireEvent(DISPID dispid, DISPPARAMS *pParams, VARIANT *pVarResult) {
		HRESULT hr = S_OK;
		for (auto ite = m_adviseMap.begin(); ite != m_adviseMap.end(); ++ite) {
			IDispatch *pDisp = ite->second;
			hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
				DISPATCH_METHOD, pParams, pVarResult, NULL, NULL);
		}
		return hr;
	}
#pragma endregion
#pragma region IProviderClassInfo
	virtual HRESULT __stdcall GetClassInfo(ITypeInfo **ppTI)
	{
		// 既定のCoClass(CLSID)の型情報を返す
		return m_pTypeLib->GetTypeInfoOfGuid(CLSID_MyRegFreeCOMSrv, ppTI);
	}

	virtual HRESULT __stdcall GetGUID(DWORD dwGuidKind, GUID *pGUID)
	{
		if (dwGuidKind == GUIDKIND_DEFAULT_SOURCE_DISP_IID) {
			// 既定のイベントソースのDIIDを返却する
			*pGUID = DIID__IMyRegFreeCOMSrvEvents;
		}
		return E_INVALIDARG;
	}
#pragma endregion
#pragma region オブジェクト固有のインターフェイスの実装
	virtual HRESULT __stdcall get_Name(BSTR *pVal)
	{
		*pVal = m_bstrName.copy();
		return S_OK;
	}

	virtual HRESULT __stdcall put_Name(BSTR newVal)
	{
		HRESULT hr;

		VARIANT_BOOL cancel = VARIANT_FALSE;
		{
			// NamePropertyChangingイベントの送信
			_variant_t varParams1[2];
			varParams1[1] = newVal;
			varParams1[0].byref = &cancel;
			varParams1[0].vt = VT_BOOL | VT_BYREF;
			DISPPARAMS params = { varParams1, nullptr, 2, 0 }; // 引数の積み方は逆順
			_variant_t ret;

			hr = FireEvent(1, &params, &ret);
		}
		if (SUCCEEDED(hr)) {
			if (!cancel) {
				m_bstrName = newVal;

				{
					// NamePropertyChangedイベントの送信
					_variant_t varParams2[1];
					varParams2[0] = newVal;
					DISPPARAMS params2 = { varParams2, nullptr, 1, 0 };
					_variant_t ret;
					hr = FireEvent(2, &params2, &ret);
				}
			}
		}

		return hr;
	}

	virtual HRESULT __stdcall ShowHello(void)
	{
		// C++/CLIでDotNETオブジェクトを使ってみる
		System::Text::StringBuilder buf;
		buf.Append(L"Name: ");
		buf.Append(gcnew System::String((LPCWSTR)m_bstrName));
		pin_ptr<const wchar_t> wname = PtrToStringChars(buf.ToString());

		MessageBoxW(NULL, wname, L"MyRegFreeCOMSrv", MB_ICONINFORMATION | MB_OK);
		return S_OK;
	}
#pragma endregion
};

ITypeLibPtr MyRegFreeCOMSrv::m_pTypeLib;
ITypeInfoPtr MyRegFreeCOMSrv::m_pTypeInfo;

#pragma region クラスファクトリの実装
// IMyRegFreeCOMSrvのスマートポインタ(IMyRegFreeCOMSrvPtr)の定義
_COM_SMARTPTR_TYPEDEF(IMyRegFreeCOMSrv, __uuidof(IMyRegFreeCOMSrv));

/**
 * クラスファクトリの共通実装
 */
class MyRegFreeCOMCsFactoryBase : public IClassFactory
{
private:
	static ULONG m_lockCount;

public:
#pragma region IUnknownの特殊実装(特殊な参照カウンタ)
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void **ppvObject)
	{
		if (riid == IID_IUnknown ||
			riid == IID_IDispatch ||
			riid == IID_IClassFactory) {
			*ppvObject = static_cast<IClassFactory *>(this);
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	virtual ULONG __stdcall AddRef()
	{
		LockModule(TRUE);
		return 2;
	}

	virtual ULONG __stdcall Release()
	{
		LockModule(FALSE);
		return 1;
	}
#pragma endregion

	virtual HRESULT __stdcall CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject) = 0;

	virtual HRESULT __stdcall LockServer(BOOL fLock)
	{
		LockModule(fLock);
		return S_OK;
	}

#pragma region クラスファクトリのロックカウント制御
public:
	static void LockModule(BOOL fLock)
	{
		if (fLock) {
			InterlockedIncrement(&m_lockCount);
		}
		else {
			InterlockedDecrement(&m_lockCount);
		}
	}

	static ULONG GetLockCount()
	{
		return m_lockCount;
	}
#pragma endregion
};

ULONG MyRegFreeCOMCsFactoryBase::m_lockCount = 0;

/**
* MyRegFreeCOMSrvのクラスファクトリ。
* 参照カウンタでは寿命を管理しない。
* DLLで一度だけ構築される、staticな生存期間をもつ。
*/
class MyRegFreeCOMSrvFactory : public MyRegFreeCOMCsFactoryBase
{
public:
	virtual HRESULT __stdcall CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
	{
		if (pUnkOuter != NULL) {
			return CLASS_E_NOAGGREGATION;
		}
		MyRegFreeCOMSrv *pImpl = new MyRegFreeCOMSrv();
		IMyRegFreeCOMSrvPtr pIntf(pImpl);
		HRESULT hr = pImpl->Init(); // 初期化処理
		if (SUCCEEDED(hr)) {
			hr = pIntf.QueryInterface(riid, ppvObject);
		}
		return hr;
	}
};
#pragma endregion


/*****************************************************/
#pragma region DotNETのマネージドクラスによるCOMオブジェクトの作成例 (イベントは実装してみたが、接続ができないので使えない。)

using namespace System::Runtime::InteropServices;

/**
 * マネージドのインターフェイスの定義
 */
[ComVisible(true)]
[InterfaceType(System::Runtime::InteropServices::ComInterfaceType::InterfaceIsDual)]
[Guid("63E19300-A485-4F5C-A9A6-5CC3009145A6")]
public interface class IMyRegFreeCOMDotnetSrv {
	[DispId(1)]
	virtual property System::String ^ Name;

	[DispId(2)]
	virtual void ShowHello();
};

// {63E19300-A485-4F5C-A9A6-5CC3009145A6}
static const GUID IID_IMyRegFreeCOMDotnetSrv =
{ 0x63e19300, 0xa485, 0x4f5c,{ 0xa9, 0xa6, 0x5c, 0xc3, 0x0, 0x91, 0x45, 0xa6 } };

/**
 * マネージドのイベントソースの定義
 */
[ComVisible(true)]
[InterfaceType(System::Runtime::InteropServices::ComInterfaceType::InterfaceIsIDispatch)]
[Guid("D46EBB6F-9F7E-479D-9D4A-BD5E85766B88")]
public interface class _IMyRegFreeCOMDotnetSrvEvents {
	[DispId(1)]
	virtual HRESULT NamePropertyChanging([In] System::String^ name, [In, Out] System::Boolean *pCancel);

	[DispId(2)]
	virtual HRESULT NamePropertyChanged([In] System::String^ name);
};

// {D46EBB6F-9F7E-479D-9D4A-BD5E85766B88}
static const GUID DIID__IMyRegFreeCOMDotnetSrvEvents =
{ 0xd46ebb6f, 0x9f7e, 0x479d,{ 0x9d, 0x4a, 0xbd, 0x5e, 0x85, 0x76, 0x6b, 0x88 } };

public delegate void NamePropertyChangingDelegate(System::String^ name, System::Boolean *pCancel);
public delegate void NamePropertyChangedDelegate(System::String^ name);

// DotNETのCOMでIProvideClassInfo2の明示的実装を試みるが、うまく行かなかったのでコメントアウト
//namespace Interop
//{
//	[ComImport()]
//	[InterfaceType(System::Runtime::InteropServices::ComInterfaceType::InterfaceIsIUnknown)]
//	[Guid("B196B286-BAB4-101A-B69C-00AA00341D07")]
//	public interface class IProvideClassInfo
//	{
//		virtual int GetClassInfo([Out] System::IntPtr *ppTI);
//	};
//
//	[ComImport()]
//	[Guid("A6BC3AC0-DBAA-11CE-9DE3-00AA004BB851")]
//	public interface class IProvideClassInfo2 : public IProvideClassInfo
//	{
//		virtual int GetGUID(DWORD dwGuidKind, [Out] GUID *pGuid);
//	};
//}

/**
 * マネージドのCOMクラスの実装
 */
[ComVisible(true)]
[ClassInterface(System::Runtime::InteropServices::ClassInterfaceType::None)]
[Guid("3E607578-2BAB-40B2-97C4-662B7675E194")]
[ComSourceInterfaces(_IMyRegFreeCOMDotnetSrvEvents::typeid)] // イベント
public ref class MyRegFreeCOMDotnetSrv
	: public IMyRegFreeCOMDotnetSrv
	//, public Interop::IProvideClassInfo2
{
private:
	System::String^ m_Name;

public:
	virtual property System::String ^Name {
		System::String^ get() {
			return m_Name;
		}
		void set(System::String^ val) {
			System::Boolean cancel = false;
			NamePropertyChanging(val, &cancel);
			if (!cancel) {
				m_Name = val;
				NamePropertyChanged(val);
			}
		}
	}

	virtual void ShowHello()
	{
		System::Text::StringBuilder buf;
		buf.Append(L"Name: ");
		buf.Append(Name);
		pin_ptr<const wchar_t> wname = PtrToStringChars(buf.ToString());

		MessageBoxW(NULL, wname, L"MyRegFreeCOMDotnetSrv", MB_ICONINFORMATION | MB_OK);
	}

	//virtual int GetClassInfo([Out] System::IntPtr *ppTI)
	//{
	//	return E_NOTIMPL;
	//}

	//virtual int GetGUID(DWORD dwGuidKind, GUID *pGuid)
	//{
	//	return E_NOTIMPL;
	//}

	event NamePropertyChangingDelegate ^NamePropertyChanging;
	event NamePropertyChangedDelegate ^NamePropertyChanged;
};

// {3E607578-2BAB-40B2-97C4-662B7675E194}
static const GUID CLSID_MyRegFreeCOMDotnetSrv =
{ 0x3e607578, 0x2bab, 0x40b2,{ 0x97, 0xc4, 0x66, 0x2b, 0x76, 0x75, 0xe1, 0x94 } };

/**
 * マネージドのCOM用のクラスファクトリ
 */
class MyRegFreeCOMCsSrvFactory : public MyRegFreeCOMCsFactoryBase {
private:
	static System::Guid GUID_IDispatch;
	static System::Guid GUID_IUnknown;

public:
	virtual HRESULT __stdcall CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
	{
		if (pUnkOuter != NULL) {
			return CLASS_E_NOAGGREGATION;
		}
		
		System::Guid guid = System::Guid(
			riid.Data1, riid.Data2, riid.Data3,
			riid.Data4[0], riid.Data4[1],
			riid.Data4[2], riid.Data4[3],
			riid.Data4[4], riid.Data4[5],
			riid.Data4[6], riid.Data4[7]);

		System::Type^ intf = IMyRegFreeCOMDotnetSrv::typeid;
		if (guid == intf->GUID ||
			guid == GUID_IDispatch ||
			guid == GUID_IUnknown) {
			// IUnknown, IDispatch、もしくは実装しているインターフェイスに合致すれば
			// DotNETのMyRegFreeCOMDotnetSrvオブジェクトをgcnewで作成する。
			MyRegFreeCOMDotnetSrv ^pImpl = gcnew MyRegFreeCOMDotnetSrv();

			// DotNETのオブジェクトから、COMインターフェイスを取得して返す.
			*ppvObject = (void *)Marshal::GetComInterfaceForObject(pImpl, intf);
			return S_OK;
		}

		return E_NOINTERFACE;
	}
};
System::Guid MyRegFreeCOMCsSrvFactory::GUID_IDispatch = System::Guid("00020400-0000-0000-C000-000000000046");
System::Guid MyRegFreeCOMCsSrvFactory::GUID_IUnknown = System::Guid("00000000-0000-0000-C000-000000000046");

#pragma endregion
/*****************************************************/


#pragma region COM DLLのエクスポート関数
extern "C" HRESULT __stdcall DllCanUnloadNow(void)
{
	// クラスファクトリのロックカウントが0であれば解放可能
	return MyRegFreeCOMSrvFactory::GetLockCount() == 0 ? S_OK : S_FALSE;
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	// クラスファクトリオブジェクトをstaticな寿命をもつオブジェクトとして構築
	static MyRegFreeCOMSrvFactory factory;
	static MyRegFreeCOMCsSrvFactory factory2;

	if (rclsid == CLSID_MyRegFreeCOMSrv) {
		return factory.QueryInterface(riid, ppv);
	}
	if (rclsid == CLSID_MyRegFreeCOMDotnetSrv) {
		return factory2.QueryInterface(riid, ppv);
	}

	return CLASS_E_CLASSNOTAVAILABLE;
}
#pragma endregion
