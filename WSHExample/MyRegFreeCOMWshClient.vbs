Option Explicit

Dim actctx
Set actctx = CreateObject("Microsoft.Windows.ActCtx")
actctx.manifest = "client.manifest"

Dim obj
Set obj = actctx.CreateObject("MyRegFreeCOMSrv.1") ' ネイティブC++のCOMサーバを使う場合
'Set obj = actctx.CreateObject("MyRegFreeCOMDotnetSrv.1") ' DotNETクラスのCOMサーバを使う場合

' イベントの接続
WScript.ConnectObject obj, "obj_"

obj.Name = "PiyoPiyo"
obj.ShowHello()

Sub obj_NamePropertyChanging(ByVal name, ByRef cancel)
	WScript.Echo("name changing: " & name)
	cancel = True ' VBSからは効果ない
End Sub

Sub obj_NamePropertyChanged(ByVal name)
	WScript.Echo("name changed: " & name)
End Sub
