Attribute VB_Name = "Others"
Option Explicit
Option Private Module

'�V�[�g�����l�[���_�C�A���O
Private Sub SheetRenameDialog()
    CommandBars.ExecuteMso "SheetRename"
End Sub


'----------------------------------------
'���O�I���_�C�A���O
'----------------------------------------

Private Function SelectJump() As Range
    Application.Dialogs(63).Show
End Function

'---------------------------------------------
'�ُ��ԉ���
'---------------------------------------------
Private Sub DeleteErrName()
    Dim v As name
    On Error Resume Next
    For Each v In ActiveWorkbook.Names
        If v.Value Like "*[#]REF!*" Then
            v.Delete
        End If
    Next v
    On Error GoTo 0
End Sub

