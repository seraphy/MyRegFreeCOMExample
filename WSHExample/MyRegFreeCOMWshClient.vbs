Option Explicit

Dim actctx
Set actctx = CreateObject("Microsoft.Windows.ActCtx")
actctx.manifest = "client.manifest"

Dim obj
Set obj = actctx.CreateObject("MyRegFreeCOMSrv.1") ' �l�C�e�B�uC++��COM�T�[�o���g���ꍇ
'Set obj = actctx.CreateObject("MyRegFreeCOMDotnetSrv.1") ' DotNET�N���X��COM�T�[�o���g���ꍇ

' �C�x���g�̐ڑ�
WScript.ConnectObject obj, "obj_"

obj.Name = "PiyoPiyo"
obj.ShowHello()

Sub obj_NamePropertyChanging(ByVal name, ByRef cancel)
	WScript.Echo("name changing: " & name)
	cancel = True ' VBS����͌��ʂȂ�
End Sub

Sub obj_NamePropertyChanged(ByVal name)
	WScript.Echo("name changed: " & name)
End Sub
