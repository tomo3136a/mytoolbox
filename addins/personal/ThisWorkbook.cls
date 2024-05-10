VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'PERSONAL.XLSB の ThisWorkbook に内容をコピーする。
'コピー後 EXCEL を再起動する。

Option Explicit

'========================================
'デバッグ用変数
'========================================

'デバッグアドイン名
Private Const addin_name_default As String = "works"
Private addin_name As String


'========================================
'アドインON/OFF設定
'========================================

Public Sub Workbook_Open()
    Application.OnKey "^w"
    Application.OnKey "^w", "ThisWorkbook.SetupAddin"
End Sub

Public Sub SetupAddin()
    If addin_name = "" Then addin_name = addin_name_default
    If AddIns(addin_name) Is Nothing Then
        AddIns.Add (ThisWorkbook.path & "\" & addin_name & ".xlsm")
    End If
    AddIns(addin_name).Installed = Not AddIns(addin_name).Installed
End Sub

