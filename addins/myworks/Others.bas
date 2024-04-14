Attribute VB_Name = "Others"
Option Explicit
Option Private Module

'シート名リネームダイアログ
Private Sub SheetRenameDialog()
    CommandBars.ExecuteMso "SheetRename"
End Sub


'----------------------------------------
'名前選択ダイアログ
'----------------------------------------

Private Function SelectJump() As Range
    Application.Dialogs(63).Show
End Function

'---------------------------------------------
'異常状態解消
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

