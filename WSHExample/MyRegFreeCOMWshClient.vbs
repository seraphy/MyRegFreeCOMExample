Option Explicit

Dim actctx
Set actctx = CreateObject("Microsoft.Windows.ActCtx")
actctx.manifest = "client.manifest"

Dim obj
Set obj = actctx.CreateObject("MyRegFreeCOMSrv.1")
WScript.ConnectObject obj, "obj_"

' DotNET�N���X��COM�T�[�o���g���ꍇ (�������AConnectObject�Őڑ��ł��Ȃ��B)
'Set obj = actctx.CreateObject("MyRegFreeCOMDotnetSrv.1")

obj.Name = "PiyoPiyo"
obj.ShowHello()

Sub obj_NamePropertyChanging(ByVal name, ByRef cancel)
	WScript.Echo("name changing: " & name)
	cancel = True ' VBS����͌��ʂȂ�
End Sub

Sub obj_NamePropertyChanged(ByVal name)
	WScript.Echo("name changed: " & name)
End Sub
