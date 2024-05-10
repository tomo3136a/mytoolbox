Attribute VB_Name = "Api"
Option Explicit
'Option Private Module

'ユーザアドインフォルダ取得
Function AddinsPath() As String
    AddinsPath = ThisWorkbook.path
End Function

'アドイン名取得
Function AddinName(Optional name As String) As String
    If name = "" Then name = ThisWorkbook.name
    AddinName = Replace(name, ".xlam", "")
End Function

Sub AddinMode()
    MsgBox "aa"
End Sub

Private Sub AddinMode2()

End Sub

