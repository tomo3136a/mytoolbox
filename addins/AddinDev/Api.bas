Attribute VB_Name = "Api"
Option Explicit
'Option Private Module

'���[�U�A�h�C���t�H���_�擾
Function AddinsPath() As String
    AddinsPath = ThisWorkbook.path
End Function

'�A�h�C�����擾
Function AddinName(Optional name As String) As String
    If name = "" Then name = ThisWorkbook.name
    AddinName = Replace(name, ".xlam", "")
End Function

Sub AddinMode()
    MsgBox "aa"
End Sub

Private Sub AddinMode2()

End Sub

