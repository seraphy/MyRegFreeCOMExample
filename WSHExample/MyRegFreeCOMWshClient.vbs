Option Explicit

Dim actctx
Set actctx = CreateObject("Microsoft.Windows.ActCtx")
actctx.manifest = "client.manifest"

Dim obj
Set obj = actctx.CreateObject("MyRegFreeCOMSrv.1")
WScript.ConnectObject obj, "obj_"

' DotNETクラスのCOMサーバを使う場合 (ただし、ConnectObjectで接続できない。)
'Set obj = actctx.CreateObject("MyRegFreeCOMDotnetSrv.1")

obj.Name = "PiyoPiyo"
obj.ShowHello()

Sub obj_NamePropertyChanging(ByVal name, ByRef cancel)
	WScript.Echo("name changing: " & name)
	cancel = True ' VBSからは効果ない
End Sub

Sub obj_NamePropertyChanged(ByVal name)
	WScript.Echo("name changed: " & name)
End Sub
