# C++でATLを使わずにレジストリフリーのCOMサーバーを作成してWSHから利用する方法

## 概要

- C++で、ATLを使わずにCOM DLLを作成する。ただし、レジストリは一切使わない。
- 作成したCOM DLLはレジストフリーで利用できるように、Side by Side Assemblyのマニフェストをつける
- WSHから、**Microsoft.Windows.ActCtx**コントロールを使ってSxSでCOMを呼び出す

COM DLLを作成するにはATLでウィザードを使ってプロジェクトを作成すると、とても簡単にできる。

しかし、残念ながら**Visual Studio Express 2017 For Windows Desktop**にはATLは付属していない。

だが、レジストリフリーのCOMを作成するのであれば、ATLの力を借りずともVC++の言語サポートと、Expressの機能だけで、比較的容易に作成できる。

以下は、VS Express 2017のC++でレジストリフリーのCOM DLLを作成する手順である。

## 手順1: プロジェクトの作成

プロジェクトの種類には「C++」の「Windows DLL」を選択する。

"MyRegFreeComSrv"というプロジェクト名にする。

ソリューションも同時に作成して、ソリューション名は「MyRegFreeCom」にしておく。
(あとでCOMを利用するクライアントのプロジェクトも、このソリューションに作るため)

Gitリポジトリの生成も行うと、ウィザードが完了したときに、はじめからローカルリポジトリが作成済みになる。
(VS用のgitignoreの設定もされている。)

動くことを確認したら、ときどきコミットして、いろいろ試行錯誤して間違えても一発で戻れるようにしておくと気が楽である。

## 手順2: DllMainでモジュールハンドルを待避する

ウィザードで生成した直後の状態では、Windows DLLのエントリポイントの「DllMain」関数だけが定義されている。

``DllMain``` 関数は、DLLがプロセスにロードされたとき、スレッドにアタッチされたとき等々のイベントで、この関数を呼び出す。

ここで重要なのはDLLのモジュールハンドルが取得されることである。

これは後で使うのでグローバル変数として保存しておく。(自分自身のDLLからリソースを取得する場合などで)

```cpp
HMODULE g_hModule;

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
		g_hModule = hModule;
		break;

	case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}
```

## 手順3: DEFファイルによるDLL関数のエクスポート

In-Process COMサーバーとして機能するDLLを作成するには、最低でも以下の2つの関数をエクスポートする必要がある。

- DllCanUnloadNow (DLLをアンロードしてよいか判定するためのもの)
- DllGetClassObject (CLSIDを指定してCOMオブジェクトを生成して返すもの)

(このほかにレジストリへの登録、登録解除のためのエントリもあるが、これは今回は不要である)

この2つの関数をエクスポートするために、昔ながらのDEFファイルによるエクスポートを行う。


```
LIBRARY "MyRegFreeCOMSrv"

EXPORTS
	DllCanUnloadNow PRIVATE
	DllGetClassObject PRIVATE
```

LIBRARYには、作成されるDLL名を指定する。
(これはプロジェクトプロパティで指定するDLL名と一致していなければならない。
ウィザードにより生成されたあとは既定では"プロジェクト名.DLL"として設定されている。)

この2つの関数は、(どこでもいいが、今回は)"MyRegFreeComSrv.cpp"で実装するものとする。

```cpp
#pragma region COM DLLのエクスポート関数

extern "C" HRESULT __stdcall DllCanUnloadNow(void)
{
	return S_OK;
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	return CLASS_E_CLASSNOTAVAILABLE;
}
#pragma endregion
```
まだ実装の中身はないので、とりあえず、いまは常に定数値を返しておく。

ちなみに、VC++では```pragma region```でソースコードの折りたたみ範囲を指定できる。

(この時点で、まずコンパイルできるか確認する。)


## 手順4: IDLファイルの作成

次に、IDLファイルを作成する。

IDLとは、インタフェース定義言語(Interface Definition Language)であり、COMインターフェイスの定義と、COMオブジェクトが、どのCOMインターフェイスを実装しているか、などの「型情報」を記述するものである。

これはコンパイルすることで、C++のソースとヘッダを生成し、タイプライブラリという型情報を記録したバイナリ形式のファイルも出力する。

特に「タイプライブラリ」はCOMをスクリプトから利用する場合には型情報を知るためにきわめて重要なものである。

生成されたC++のソースとヘッダはCOMの実装時に使う他、クライアントがCOMインターフェイスを参照する場合にも使える。

idlに、以下のIMyRegFreeCOMSrvというインターフェイスを定義する。

(一般的にインターフェイス名は「I」で始まることになっている。)

これが、今回作成するCOMのインターフェイス定義となる。

```idl
import "oaidl.idl";
import "ocidl.idl";

[dual, uuid(E17C4111-F731-44E6-B262-D45D1241DD75)]
interface IMyRegFreeCOMSrv : IDispatch
{
	[propget, id(1)] HRESULT Name([out, retval] BSTR* pVal);
	[propput, id(1)] HRESULT Name([in] BSTR newVal);
	[id(2)] HRESULT ShowHello();
};
```

この定義が意味するところは、


```[dual]``` というのは、これがスクリプトから呼び出すことのできるデュアルインターフェイスであることを示している。

```[uuid(xxxx)]``` には、このインターフェイスを識別するユニークなIDを割り当てる。

これはVisual Studioのツールメニューから「GUIDの作成」を選んで、そこからランダムなGUIDをもらってくる。

(COMではオブジェクト、インターフェイス、ライブラリに、それぞれ世界で一意になるようなランダムな値をつけ、
このGUIDによって識別するようになっている。異なるものに同じ値は使ってはならない。)

```[propget]```, ```[propput]``` は、プロパティのgetter/setterの定義を表す。

COMではプロパティもメソッドとして実装されるので、このような形で定義される。

(C#などでプロパティを扱っていれば分かると思う。)

```[id(1)]``` は、DISPIDの定義である。スクリプトからCOMを呼び出す場合の番号となる。

このDISPIDは1から始まる任意の番号をつけてよいが、いくつか特殊な番号がある。

たとえば「DISPID_VALUE(0)」はデフォルトを表すもので、DISPIDが0のメソッドまたはプロパティは、スクリプト側でメソッド名またはプロパティ名を省略した場合に暗黙で使われるようになったり、あるいは、For Eachで列挙子を返すためのメソッドを表す、DISPID_NEWENUM(-4)などの定義がある。

今回は、とくに使わないので、1以上の適当な番号をふっておけばよい。

また、COMでは全て結果コードをHRESULTとして返すので、すべてのプロパティ、メソッドはHRESULTを返す。

引数にはIn/Outの区別があるので、```[in]```または```[in, out]```のように指定し、戻り値は```[out, retval]```のように指定する。

```[out]``` で指定した引数には、返却用のアドレスをクライアントから渡されるので、そこに書き込めば良い。
(返却用のアドレスを用意するのはクライアントの責務である。)

### 手順4-2: タイプライブラリとイベントソースの定義

つぎに、このインターフェイスを実装するCOMクラス(CoClass)とイベントソース、タイプライブラリの定義を行う。

まず、COMクラスとイベントソースはタイプライブラリの中に包含するように記述する。

```idl
[uuid(CE20ECF4-B344-49DE-AD01-CDCF7098108D), version(1.0)]
library MyRegFreeCOMSrvLib
{
	importlib("stdole32.tlb");

	[
		uuid(8C11D374-E2BF-4DEF-89AB-81756137C1D0)
	]
	dispinterface _IMyRegFreeCOMSrvEvents
	{
	properties:
	methods:
		[id(1)] HRESULT NamePropertyChanging([in] BSTR Name,[in, out] VARIANT_BOOL *pCancel);
		[id(2)] HRESULT NamePropertyChanged([in] BSTR Name);
	};

	[uuid(4475E395-FA5A-42A3-901C-F3062802B9B2)]
	coclass MyRegFreeCOMSrv
	{
		[default] interface IMyRegFreeCOMSrv;
		[default, source] dispinterface _IMyRegFreeCOMSrvEvents;
	};
}
```

```library``` がタイプライブラリの指定であり、タイプライブラリのIDをuuidで指定している。

バージョンも指定できるが、今回はレジストリに登録しないCOMであるし、とりあえず```1.0```にしている。


```importlib("stdole32.tlb");``` はシステム標準のIUnknownやIDispatchなどの既定のタイプライブラリをインポートするもので常に必要である。


この下に、まずイベントソースを指定する。

イベントソースは、WSHで```WScript.ConnectObject``` のようにしてオブジェクトからのイベント通知を受けるためのもので、そのイベントについて定義するものである。

COMにおいてはイベントは、(通常は)IDispatchインターフェイスで実装するので、```dispinterface```で定義する。

イベントソースもインターフェイスの一種であるから、識別子としてuuidを必要とする。

今回は以下の2つのイベントをもつものとする。

- NamePropertyChanging
- NamePropertyChanged

### 手順4-3: COMクラスの定義

つぎに、インターフェイスを実装するCOMクラス(CoClass)の定義を行う。

COMクラスは、libraryの中の```coclass```で定義する。

こちらはクラスを識別するCLSIDとなるuuidを指定する必要がある。

中身には、このクラスが実装するインターフェイスと、このクラスが送信するイベントのインターフェイスの指定となる。

```[source]``` と指定されているインターフェイスがイベントソースとなるインターフェイスである。

### 手順4-4: IDLのコンパイル

ここまでIDLを定義すれば、```MIDL```によってIDLをビルドすることができる。

ソリューションエクスプローラからidlファイルを右クリックして「コンパイル」を選択すると、コンパイルが行われる。

すると、以下のファイルが生成される。

- MyRegFreeCOMSrv_h.h C++用に翻訳されたCOMインターフェイス、COMクラスのヘッダファイル
- MyRegFreeCOMSrv_i.c C++用に翻訳されたCOMインターフェイス、COMクラスのソースファイル(CLSID, LIBID, IIDなどの定数定義)
- Debug/MyRegFreeCOMSrv.tlb 型情報を格納したタイプライブラリ
- MyRegFreeCOMSrv_p.c プロキシ・スタブ用のルーチン(今回不要)
- dlldata.c マーシャリングを行うためのルーチン(今回不要)

なお、これらのヘッダやソースはコンパイル時にも必要となるが、
idlからの生成物であるのでgitignoreには既定で(だいたい)除外設定されている。(*_h.hは入ってないので自分で追加する)

生成された「MyRegFreeCOMSrv_h.h」と「MyRegFreeCOMSrv_i.c」はプロジェクトに加えてビルドされるようにする。

追加する際には、プロジェクトに「生成されたファイル」というフィルタを追加して、そこに登録しておくと良い。

また、「MyRegFreeCOMSrv_i.c」はstdafx.hを使わないので、プリコンパイルヘッダの使用をOffにしておく。


### 手順5: タイプライブラリのリソースへの埋め込み

つぎに生成されたタイプライブラリをwin32リソースに含めるようにする。

タイプライブラリ(*.tlb)は、スクリプトからCOMを利用する際に型情報を取得するために、よく使われている。

タイプライブラリは、タイプライブラリの保存場所をレジストリに登録して、レジストリを経由してロードする方法があるが、
自分のDLL内にリソースとして格納し、レジストリを介さず、自分自身のリソースからタイプラブラリをロードする方法もある。

今回はレジストリフリーでやりたいので、リソース内におく必要がある。

Visual Studio Express 2017 for Windows Desktopでは、リソースエディタは存在しないが、リソース自身はプロジェクトに追加できるし、コンパイルもできる。

リソースを追加したら、リソースエディタは開かないが、**テキストエディタとして開くことは可能**である。

末尾にある「#ifndef APSTUDIO_INVOKED」のブロックの内に、以下のようにタイプライブラリをリソースとして取り込むように指定する。

(リソース的には、どこでもよいのだが、このNOT APSTUDIO_INVOKEDの範囲であれば、Visual Studioのリソースエディタが勝手に書き換えない場所となる。)

```
#ifndef APSTUDIO_INVOKED
/////////////////////////////////////////////////////////////////////////////
//
// TEXTINCLUDE 3 リソースから生成されました。
//

1 TYPELIB "MyRegFreeCOMSrv.tlb"

/////////////////////////////////////////////////////////////////////////////
#endif    // APSTUDIO_INVOKED でない場合
```

なお、タイプライブラリのリソースIDは1としておく。

これは```LoadTypeLib``` が、このDLLからタイプライブラリをロードするとき、デフォルトでは、最初に見つかったタイプライブラリを使用するためである。

また、ファイル名として「MyRegFreeCOMSrv.tlb」を指定しているが、タイプライブラリはビルド時の中間フォルダに出力されているので、実際にはDebug, Release、あるいはx64フォルダの下にいる。(タイプライブラリはx64/x86で生成される内容が異なるので。)
なので、ビルド時にリソースの参照パスを設定する必要がある。

プロジェクトのプロパティページからリソースを選択して「追加のインクルードディレクトリ」に**$(IntDir)**を追加する。

これで、中間ディレクトリ内の"MyRegFreeCOMSrv.tlb"ファイルが参照できるようになる。

(ここまででビルドが通ることを確認しておくと良い。)

## 手順6: MyRegFreeCOMSrvクラスの実装

つぎに、いよいよ、MyRegFreeCOMSrvクラスの実装を行う。

今回は、先ほどDLLのエクスポート関数を定義した、```MyRegFreeCOMSrv.cpp```ファイル内に実装することにする。

以下のようにクラスを書き始める。

```cpp
#include "stdafx.h"

#include "MyRegFreeCOMSrv_h.h"

class MyRegFreeCOMSrv
	: public IMyRegFreeCOMSrv
{
};

... このうしろに前述のDLLエクスポート関数 ...
```

インターフェイス名、クラス名は、IDLで指定したinterface, coclassの名前と一致している必要がある。

これは、ヘッダファイル"MyRegFreeCOMSrv_h.h"を見れば分かるが、ここで以下のように先行宣言されている。

```cpp
class DECLSPEC_UUID("4475E395-FA5A-42A3-901C-F3062802B9B2")
MyRegFreeCOMSrv;
```

クラス名だけ先行定義をしているが、そこに属性として```__declspec(uuid(...))```としてIDLで指定したuuidが付与されている。

なので、このクラス名に対してVC++の拡張である```__uuidof(MyRegFreeCOMSrv)```という関数を呼び出すことで、クラス名から、それに割り当てられているUUID(4475E395-FA5A-42A3-901C-F3062802B9B2)を得られる、という便利な仕組みになっている。


また、今回作成するCOMには```_IMyRegFreeCOMSrvEvents```というイベントソースをもつことになっている。

C++のCOMでイベントソースを定義するのは、ちょっと面倒くさくて、```IConnectionPoint```というインターフェイスを経由して、クライアント側が提供するイベント受け取り側(イベントシンク)と接続する、という方法を採らなければならない。

そして、この```IConnectionPoint```は、それを保持する```IConnectionPointContainer```インターフェイスから取得されなければならないので、結局、最低でも、```MyRegFreeCOMSrv```クラスは、それ自身が実装すべき```IMyRegFreeCOMSrv```インターフェイスの他に、イベント接続のための```IConnectionPointContainer```と```IConnectionPoint```も実装する必要がある。

しかし、更にWSHなどのスクリプト言語からイベントを接続可能にするためには、型情報をスクリプト側に教えてやるための```IProvideClassInfo2```インターフェイスも実装しなければならない。
(IProvideClassInfo2が実装されていないと、WSHの```WScript.ConnectObject()```を実行すると、接続できませんというようなエラーになる。なお、IProvideClassInfo2がなくても、```WScript.CreateObject("progid", "prefix_")```だと接続できる。)

なので、結局、以下のような形になる。

```cpp
class MyRegFreeCOMSrv
	: public IMyRegFreeCOMSrv
	, public IConnectionPointContainer
	, public IConnectionPoint
	, public IProvideClassInfo2
{
};
```

### 手順6-2: IUnknownの実装

まずは、もっとも基本となる、IUnknownで参照カウンタとインターフェイスの問い合わせ部について実装する。

```cpp
class MyRegFreeCOMSrv
	: public IMyRegFreeCOMSrv
	, public IConnectionPointContainer
	, public IConnectionPoint
	, public IProvideClassInfo2
{
private:
	ULONG m_refCount;

	virtual ~MyRegFreeCOMSrv()
	{
	}

public:
	MyRegFreeCOMSrv()
		: m_refCount(0)
	{
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
};
```

```Release``` で参照カウンタが0になったら、```delete this```で自分を破棄するようにしている。

そのため、デストラクタはprivateにおく。(これにより、ヒープ上のみ作成可能なオブジェクトとなる。)

```QueryInterface``` は要求されたインターフェイスに対応するポインタをキャストして返すだけである。

(複数のIUnknown派生インターフェイスを継承しているので、それ単位でポインタが異なる。)

### 手順6-3: IDispatchの実装

つぎに、IDispatchの実装を行う。

IDispatchはスクリプト言語用のインターフェイスであり、文字列で関数名やプロパティ名を指定して、その番号を取得したり、その番号に対してパラメータを渡すことで、さまざまな処理を行ったり、というような動的な処理が行われることになる。

これらの動的処理のコードを自分で実装すればかなり面倒なコードになるのだが、タイプライブラリにある型情報を使うことで、これらを代行してくれる**標準のWindows API**が用意されている。

なので、タイプライブラリをロードして代行するAPIに丸投げするだけでよい。

```cpp
#include <comdef.h>

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

	virtual ~MyRegFreeCOMSrv()
	{
		if (m_pTypeInfo) {
			m_pTypeInfo.Release();
		}
		if (m_pTypeLib) {
			m_pTypeLib.Release();
		}
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

....IUnknownの実装....

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
};
```

```ITypeLibPtr```, ```ITypeInfoPtr``` の使用のために```comdef.h``` ヘッダをインクルードする。

タイプライブラリは先ほど、自分自身(DLL)にリソースとして埋め込んでいるので、DLL自身のパスを指定して```LoadTypeLib``` APIによってロードすることができる。

DLLのパスを得るには、DLLのモジュールハンドルが必要になるので、最初にグローバル変数として保存した```g_hModule```を使うことになる。(別ソースファイルで定義しているので、```extern```で参照している。)

また、タイプライブラリは変更されないものなので、最初に一度読み込めば十分なので、Staticとしている。

```DispGetIDsOfNames``` や```DispInvoke``` が標準のDispatch代行APIである。

```DispInvoke``` ではタイプライブラリからDISPIDに対応するメソッドの呼び出すを行う。


### 手順6-4: IConnectionPointContainerの実装

次に```IConnectionPointContainer``` の実装を行う。

これは単にイベントソースのIIDであるか判定して、自身のコネクションポイントを返せばよいだけである。

```cpp
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
```

### 手順6-5: IConnectionPointの実装

次に```IConnectionPoint``` の実装を行う。

IConnectionPointは、自身がサポートしているコネクションインターフェイスのUUIDの返却、自身を含むIConnectionPointContainerの返却、および、クライアントとの接続と接続解除の処理を実装する。

クライアントはIDispatchを実装したイベントシンクを渡すので、それに対して```DWORD```のCookie値を(適当に)割り当てて返却する。

オブジェクトは保持しているイベントシンクに対して必要に応じて```IDispatch#Invoke```を用いてイベントを通知する。

クライアントから接続解除のためにCookie値を渡されたら、これでイベントシンクを解放する。


cookieと保持しているイベントシンクのIDispatchを保存するためのメンバフィールドと、cookie値を管理するカウンタが必要である。

```cpp
private:
	// クッキーがかぶらないように管理するクッキーの最終値
	ULONG m_lastCookie;

	// コネクションポイントのアドバイスしているイベントシンク
	std::map<ULONG, IDispatch*> m_adviseMap;

public:
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
```

```std::map``` を使っているので、```#include <map>``` でSTLのmapをインクルードしておく。

```FireEvent``` はイベントシンクに対してイベントを送信するためのヘルパ。

また、デストラクタでイベントシンクの解放漏れがあれば、ここで解放しておく。

```cpp
	virtual ~MyRegFreeCOMSrv()
	{
		....省略....

		for (auto ite = m_adviseMap.begin(); ite != m_adviseMap.end(); ++ite) {
			ite->second->Release();
		}
		m_adviseMap.clear();
	}

```

### 手順6-6: IProvideClassInfo2の実装

次に、```IProvideClassInfo2``` の実装を行う。

これは単にタイプライブラリの情報をクライアントに引き渡すだけのものである。

```cpp
#pragma region IProvideClassInfo2
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
```

### 手順6-7: IMyRegFreeCOMSrvの実装

最後に、ようやくオブジェクト固有のインターフェイスの実装ができる。

```cpp
private:
	_bstr_t m_bstrName;

public:
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
		MessageBoxW(NULL, (LPCWSTR)m_bstrName, L"MyRegFreeCOMSrv", MB_ICONINFORMATION | MB_OK);
		return S_OK;
	}
#pragma endregion
```

```put_Name``` では、Nameプロパティを更新する際に、イベントシンクに対してイベントを送信している。

## 手順7: IClassFactoryの実装

これで、MyRegFreeCOMSrvオブジェクトは完成したが、これをDLL外から利用するには、生成するためのファクトリが必要である。
ファクトリは、```IClassFactory``` インターフェイスを実装したCOMオブジェクトであるが、DLLと生存期間が一致するものなので、参照カウンタのような寿命管理のメカニズムは必要ない。

(自分で```delete this```することはなく、staticでインスタンス化して使う。)

しかし、DLLをアンロードして良いか？という```DllCanUnloadNow```関数の問い合わせに必要なロックカウントは保持する必要がある。

また、```DllGetClassObject``` 関数ではCLSIDを指定してファクトリを取得する必要があるので、そのあたりも踏まえたコードは、以下のような感じになる。

```cpp
#pragma region クラスファクトリの実装
// IMyRegFreeCOMSrvのスマートポインタ(IMyRegFreeCOMSrvPtr)の定義
_COM_SMARTPTR_TYPEDEF(IMyRegFreeCOMSrv, __uuidof(IMyRegFreeCOMSrv));

/**
* MyRegFreeCOMSrvのクラスファクトリ。
* 参照カウンタでは寿命を管理しない。
* DLLで一度だけ構築される、staticな生存期間をもつ。
*/
class MyRegFreeCOMSrvFactory
	: public IClassFactory
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
#pragma region クラスファクトリの実装
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

	virtual HRESULT __stdcall LockServer(BOOL fLock)
	{
		LockModule(fLock);
		return S_OK;
	}
#pragma endregion
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

ULONG MyRegFreeCOMSrvFactory::m_lockCount = 0;
#pragma endregion

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

	if (rclsid == CLSID_MyRegFreeCOMSrv) {
		return factory.QueryInterface(riid, ppv);
	}

	return CLASS_E_CLASSNOTAVAILABLE;
}
#pragma endregion

```

クラスファクトリもCOMオブジェクトではあるが、外部からCLSIDを使用してインスタンス化されることはなく、
CLSIDで区別されることもないのでUUID属性の定義は不要である。

以上まで実装すれば、ビルドは可能になっている。

## 手順8: レジストリフリーCOMのためのSide by side assembly manifestの設定

あとはCLSIDやPROGIDのレジストリを登録すればCOMオブジェクトとして起動できるようになっているはずである。

今回はレジストリフリーのCOMオブジェクトとしたいので、レジストリではなくSide By Side Assemblyのマニフェストを作成する。

ソリューションエクスプローラから追加するが、マニフェストファイルのテンプレートは存在しないので、テキストファイルを選択して名前を「MyRegFreeCOMSrv.manifest」として作成する。

追加したファイルのプロパティを確認して、項目の種類が「マニフェストツール」になっていればOKである。

(マニフェストツールになっていれば、マニフェストは暗黙でリソースファイルに埋め込まれる。)


ここにレジストリのかわりにCOMの情報を書き込む。(テキストは初期状態でSJISになっていると思われるので、保存するときは**UTF-8**であることを確認すること。)

```xml
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">

	<assemblyIdentity
        type="win32"
        name="MyRegFreeCOMSrv"
        version="1.0.0.0" />

  <!-- 自分自身がもっているCOMの情報についてマニフェストにアセンブリとして登録する -->
  <file name="MyRegFreeCOMSrv.dll">
    <typelib tlbid="{CE20ECF4-B344-49DE-AD01-CDCF7098108D}"
          version="1.0" helpdir="" flags="HASDISKIMAGE"/>
    <comClass
          clsid="{4475E395-FA5A-42A3-901C-F3062802B9B2}"
          threadingModel="Apartment"
          tlbid="{CE20ECF4-B344-49DE-AD01-CDCF7098108D}"
          progid="MyRegFreeCOMSrv">
      <progid>MyRegFreeCOMSrv.1</progid>
    </comClass>
    
    <!-- カスタムインターフェイスの場合で、スレッドアパートメントが異なる場合はマーシャリングが必要なため、
      DLLがProxy/Stubをマージしている場合は、このDLLを示す<file>要素の子として、comInterfaceProxyStubの定義が必要。
      (Proxy/Stubが別DLLの場合は、*PS.DLLの<file>要素の子として作成する。)
    <comInterfaceProxyStub name="ISxSCustom" iid="{[IID_ISxSCustom]}" />
    -->
    
  </file>

  <!-- comInterfaceExternalProxyStubは、
  IDispatch派生クラスでスレッドアパートメントの互換性がない場合のマーシャリングのために使われる
  -->
  <comInterfaceExternalProxyStub
   name="_IMyRegFreeCOMSrvEvents"
   iid="{8C11D374-E2BF-4DEF-89AB-81756137C1D0}"
   tlbid="{ECEF3763-1C5F-41A9-AD2D-6BBD335F576C}"
   proxyStubClsid32="{00020420-0000-0000-C000-000000000046}"/>

  <comInterfaceExternalProxyStub
   name="IMyRegFreeCOMSrv"
   iid="{E17C4111-F731-44E6-B262-D45D1241DD75}"
   tlbid="{ECEF3763-1C5F-41A9-AD2D-6BBD335F576C}"
   proxyStubClsid32="{00020420-0000-0000-C000-000000000046}"/>

</assembly>
```

アセンブリマニフェストの文法については https://msdn.microsoft.com/en-us/library/windows/desktop/aa374191(v=vs.85).aspx のあたりを見る。

要点は

- ```<assemblyIdentity>```で自分のアセンブリ名を宣言する
- ```<file name="*.dll">...</file>```で自分のDLL名を指定し、このDLLが提供するCOMの情報を列挙する。
  - ```<typelib>```タイプライブラリのIDとバージョン
  - ```<comClass>``` COMオブジェクトのCLSID, 代表ProgIDの宣言と、使用するタイプライブラリの制限。
    - 子要素として```<progid>```を追加で、必要なだけつけられる。(バージョン別ProgIdなどを表現するため。無くても良い。)
- ```comInterfaceExternalProxyStub```はアパートメントが異なる場合にシステムがマーシャリングを行う際に利用されるようで、とりあえずつけておくことが推奨されている。(が、今回はIn-ProcのSTAのみを想定しており、今回の場合には使われることはなく厳密には不要であろう。)
  - (とりあえずインターフェイスのIID、イベントソースのDIIDをつけておく？)


また、DLLにマニフェスト中にUAC有効化の設定があるのはおかしいので、**必ず削除して**おく。

これでビルドすると、COMの情報が埋め込まれた、Side by Side Assemblyとして利用可能なDLLとして生成される。

## 手順9: テスト用のコンソールアプリの作成

テスト用のコンソールアプリのプロジェクトを追加する。

最終的にはWSHから使いたいが、デバッグしたり動作確認するには同じC++で作られた、同一ソリューションのプロジェクトがあったほうが便利である。

ビルド順序を指定して、MyRegFreeCOMSrvが先にビルドされるようにしておく。

とりあえず、SideBySideでCOMが使えることを確認するため、以下のような単純なクライアントコードを記述する。
(イベントまわりは、とりあえず、いまは省略。)

```cpp
#include "stdafx.h"

#include <Windows.h>
#include <comdef.h>

#include "..\MyRegFreeCOMSrv\MyRegFreeCOMSrv_h.h"
#include "..\MyRegFreeCOMSrv\MyRegFreeCOMSrv_i.c"

//IMyRegFreeCOMSrvのスマートポインタ(IMyRegFreeCOMSrvPtr)の定義
_COM_SMARTPTR_TYPEDEF(IMyRegFreeCOMSrv, __uuidof(IMyRegFreeCOMSrv));

int main()
{
	CoInitialize(nullptr);
	{
		IMyRegFreeCOMSrvPtr pSrv;
		HRESULT hr = pSrv.CreateInstance("MyRegFreeCOMSrv.1");
		if (SUCCEEDED(hr)) {
			_bstr_t name(L"PiyoPiyo");
			pSrv->put_Name(name);
			pSrv->ShowHello();
		}
	}
	CoUninitialize();
    return 0;
}
```

COMの定義のヘッダファイルや定数類はサーバ側のプロジェクトから直接参照している。

(本来は、生成されたタイプライブラリを参照して```#import "*.tlb"```したほうが、いろいろ便利である。)

このまま実行すれば、当然、レジストリに登録されていないので実行は失敗する。

## 手順9-2: クライアント用のマニフェストの作成

これに対して、アプリケーションマニフェストを追加して、依存アセンブリを明示する。

(```app.manifest``` 等の適当な名前で良い。追加方法はサーバ側と同様である。マニフェストツールとして認識されていればexeのリソースとして自動的に取り込まれる。)

```xml
<?xml version="1.0" encoding="utf-8" standalone="yes"?> 
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"> 
	<assemblyIdentity
		version="1.0.0.0"
		processorArchitecture="" 
		name="MyRegFreeCOMClient.app"
		type="win32" /> 
	<description>Side By Side RegFree COM Client</description> 

	<dependency>
		<dependentAssembly>
			<assemblyIdentity
				type="win32"
				name="MyRegFreeCOMSrv"
				version="1.0.0.0"
			/>
		</dependentAssembly>
	</dependency>
</assembly>
```

## 手順9-3: 動作確認方法

これで、クライアントプロジェクトをスタートアップに設定してデバッグ実行すれば、
CoCreateInstanceが成功してCOM呼び出しができていることが確認できる。

失敗するケースとしては、

- ```Side by Side Assembly Manifest```の設定に不備がある場合は、そもそもEXEが起動できない。
  - バージョンも含めてアセンブリ名が一致しているか？
  - DLLのマニフェストにUACの設定が入っているとSxS起動は失敗する。
  - Side by Sideまわりの起動失敗は、```SxsTrace```というツールを使うとトレースログがとれる。

> WinSxs WinSxs トレース ユーティリティ。
>
> 使用法: SxsTrace [オプション]
>
> オプション:
>
>    Trace -logfile:ファイル名 [-nostop]
>
>        sxs のトレースを有効にします。
>        トレース ログはファイル名に保存されます。
>        -nostop を指定すると、トレースを停止するかどうかを確認しません。
>
>    Parse -logfile:ファイル名 -outfile:解析ファイル  [-filter:AppName]
>
>        未処理のトレース ファイルを人が読める形式に変換し、結果を解析ファイルに
>        保存します。
>        出力をフィルターするには、-filter オプションを使用します。
>
>    Stoptrace
>
>        まだ停止していない場合は、トレースを停止します。
>
>    例:
>
>        SxsTrace Trace -logfile:SxsTrace.etl
>        SxsTrace Parse -logfile:SxsTrace.etl -outfile:SxsTrace.txt

上記のクライアントで動作することを確認する。

## 手順10: WSHからレジストリフリーで呼び出す

今度はWSHから試してみる。

以下のようなVBScriptを定義する。

```vb
Option Explicit

Dim actctx
Set actctx = CreateObject("Microsoft.Windows.ActCtx")
actctx.manifest = "client.manifest"

Dim obj
Set obj = actctx.CreateObject("MyRegFreeCOMSrv.1")
WScript.ConnectObject obj, "obj_"
obj.Name = "PiyoPiyo"
obj.ShowHello()

Sub obj_NamePropertyChanging(name, byref cancel)
	WScript.Echo("name changing: " & name)
End Sub

Sub obj_NamePropertyChanged(name)
	WScript.Echo("name changed: " & name)
End Sub
```

スクリプトと同じフォルダに以下のclient.manifestファイルを保存する。

```
<?xml version="1.0" encoding="utf-8" standalone="yes"?> 
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"> 
	<dependency>
		<dependentAssembly>
			<assemblyIdentity
				type="win32"
				name="MyRegFreeCOMSrv"
				version="1.0.0.0"
			/>
		</dependentAssembly>
	</dependency>
</assembly>
```

これは先に実験したclient用のマニフェストファイルから、依存アセンブリの部分だけにしたものである。

ここに、コンパイルされた```MyRegFreeCOMSrv.dll```ファイルも一緒に置いておく。

(もし、使用するWSHが32ビット版であれば、このMyRegFreeCOMSrv.dllも32ビットでビルドしておく。WSHが64ビット版であれば、MyRegFreeCOMSrv.dllも64ビットでビルドしなければならない。)

これで実行すると、プロパティの変更通知に対してメッセージが表示され、最終的にメッセージボックスが表示されるはずである。

### 手順10-2: Microsoft.Windows.ActCtxについて

ここで使われている、```Microsoft.Windows.ActCtx```というコントロールは、```アクティベーションコンテキスト```を明示的に開始してプログラム実行中にSide By Side Assemblyの効果を与えるActiveXである。

(これは、Windows2003以降に搭載されているようである。Window7, Windows10での動作を確認した。)

これにより、VBSの中で```client.manifest```を後からロードして、レジストリフリーでCOMオブジェクトを起動することができている。

また、取得したオブジェクトは```WScript.ConnectObject```によって接続することで、このCOMからのイベント通知も受け取ることができている。

(ただし、VBSのイベントハンドラではByRef引数を受け取っても、それを変更して返却することができないようである。このサンプルでは、本来はNamePropertyChangingで引数cancelに対してTrueを返すと変更をキャンセルできるはずだが、VBSではうまくゆかない。)

以上、VS Express 2017でATLを使わずレジストリフリーなCOMを作成することができた。


## Excel/Word等のVBAから利用する場合

WSHから利用可能であるように、VBAからも同様にレジストリフリーで利用可能である。

こちらは、WSHの場合と異なり、ByRefのイベントでcancel値を返却できる。

```vb
Private WithEvents obj As MyRegFreeCOMSrv

Sub obj_NamePropertyChanging(ByVal name As String, ByRef cancel As Boolean)
    Dim rslt As VbMsgBoxResult
    rslt = MsgBox("名前の変更: " & name, Title:="名前の変更を許可しますか？", Buttons:=vbYesNo)
    If rslt <> vbYes Then
        cancel = True ' 変更を許可しない
    End If
End Sub

Sub obj_NamePropertyChanged(ByVal name As String)
    MsgBox "name changed: " & name
End Sub

Public Sub sxstest()
    Dim dir As String
    dir = ThisWorkbook.Path
    Dim actctx As Object
    Set actctx = CreateObject("Microsoft.Windows.ActCtx")
    actctx.manifest = dir & "\client.manifest"
    
    Set obj = actctx.CreateObject("MyRegFreeCOMSrv.1")
    obj.name = "PiyoPiyo"
    Call obj.ShowHello
    Set obj = Nothing
    Set actctx = Nothing
End Sub
```
ただし、これが利用可能になるためには、**VBAの参照設定が必要**である。

生成されたDLLにはタイプライブラリがリソースとして含まれているので、このDLLを参照設定で直接指定すれば良い。

(Officeが32ビット版であれば、DLLも32ビットでビルドしていなければならない。通常、Officeは32ビット版が推奨されているが、もし64ビット版のOfficeを使っている場合は、64ビットでビルドしたDLLを使用すること。)


VBAでのイベントハンドリングでは

```vb
Private WithEvents obj As MyRegFreeCOMSrv
```

のように、WithEventsと同時に型の指定も必須であるため、コンパイル時点で型情報が必要である。


このように、レジストリフリーではあるが、DLLを参照設定しなければならない、という手間はある。

なお、**イベントを使わない場合は参照設定は不要**である。

※ あと、イベントを使う場合は、最後に```Set obj = Nothing``` で明示的に解放していないと２回目以降で、よくわからんエラーになる。(何か実装に問題があるのも？)


## 注意点、その他

なお、Side by Side Assembly ManifestでレジストリフリーにできるのはCOMのCLSID, ProgID等々のレジストリ参照の仕組みであって、COMオブジェクトが自分のソースの中でレジストリを見ていたりしたら、当然、レジストリフリーにはできない。


ちなみに、ATLで作成する場合でも、```IDispatchImpl```や```IProvideClassInfo2Impl```のタイプライブラリのバージョン指定を```-1, -1```にしておかないと、自分のDLLのタイプライブラリではなく、レジストリ経由のタイプライブラリを見に行く処理が入るため、レジストリフリーにならないので注意が必要である。


なお、In-Proc COM DLLであれば、C#でCOMオブジェクトを作成するほうが簡単なのであるが、DotNETで作成されたCOMオブジェクトはSide by side assemblyとして実行させようとすると、EXEと同じフォルダにDLLを入れておかないと動作しない。

  - (同じフォルダに入れておけばレジストリフリーで動作する。)

  - (原因はよくわからないが、エラーメッセージは「File Not Found」となっている。)

なので、今回のようにWSHで使うには、ちょっと難がある。(wscript.exe, cscript.exeを自分のフォルダにコピーすれば動作はするが...。)


## C++/CLIの対応

C++ではなく、C++/CLIにすることで、C++のコードからDotNETのクラスライブラリを自由に呼び出せるようになる。

それには、以下の変更を行う。

- プロジェクトのプロパティから「共通言語ランタイムのサポート」を「共通言語ランタイム サポート (/clr)」にする。
- IDLが生成した「MyRegFreeCOMSrv_i.c」をC++としてコンパイルするため、プロジェクトのプロパティの「C/C++」の「詳細設定」から「コンパイル言語の選択」を「既定」から「C++ コードとしてコンパイル (/TP)」にする
- いくつかの警告に対応する (なくてもコンパイルは可能)
  - DllMain関数に```#pragma unmanaged``` プラグマをつける。
  - プロジェクトのプロパティの「リンカー」の「全般」から「インクリメンタルリンクを有効にする」を「いいえ (/INCREMENTAL:NO)」
  - プロジェクトのプロパティの「リンカー」の「デバッグ」から「デバッグ情報の生成」を「共有と発行用に最適化されたデバッグ情報の生成 (/DEBUG:FULL)」

こうすれば、たとえば、以下のようなコードを書けるようになる。

```cpp
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
```

以上、メモ終了。
