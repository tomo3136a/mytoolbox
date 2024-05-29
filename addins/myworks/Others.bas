Attribute VB_Name = "Others"
Option Explicit
Option Private Module

'---------------------------------------------
'シート名リネームダイアログ
'---------------------------------------------
Private Sub SheetRenameDialog()
    CommandBars.ExecuteMso "SheetRename"
End Sub


'---------------------------------------------
'名前選択ダイアログ
'---------------------------------------------
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

'---------------------------------------------
'定型式挿入
'---------------------------------------------
Sub InsertFormula(ra As Range, mode As Integer)
    Dim v1 As Integer, v2 As Integer
    Dim r0 As Range, r1 As Range, r2 As Range
    Select Case mode
    Case 1
        '文字列分割(英字・数値)
        ra.Offset(, 1).Formula2R1C1 = _
            "=LET(v,RC[-1],LEFT(v,MIN(FIND({1,2,3,4,5,6,7,8,9,0},v&""1234567890""))-1))"
        ra.Offset(, 2).FormulaR1C1 = "=MID(RC[-2],LEN(RC[-1])+1,LEN(RC[-2]))"
    Case 2
        '文字列分割(数値・英字・数値)
        ra.Offset(, 1).FormulaR1C1 = "=IFERROR(VALUE(LEFT(RC[-1],2)),IFERROR(VALUE(LEFT(RC[-1],1)),""""))"
        ra.Offset(, 2).Formula2R1C1 = _
            "=LET(v,MID(RC[-2],LEN(RC[-1])+1,LEN(RC[-2])),LEFT(v,MIN(FIND({1,2,3,4,5,6,7,8,9,0},v&""1234567890""))-1))"
        ra.Offset(, 3).FormulaR1C1 = "=MID(RC[-3],LEN(RC[-2]&RC[-1])+1,LEN(RC[-3]))"
    Case 3
        '差分マーカー
        Set r0 = ra.Cells(1, 1)
        On Error Resume Next
        Set r1 = Application.InputBox("比較元位置1", "差分マーカ―", r0.Offset(0, 1).Address, Type:=8)
        On Error GoTo 0
        If r1 Is Nothing Then Exit Sub
        On Error Resume Next
        Set r2 = Application.InputBox("比較元位置2", "差分マーカ―", r1.Offset(0, 1).Address, Type:=8)
        On Error GoTo 0
        If r2 Is Nothing Then Exit Sub
        If r2.Address = r1.Address Then Exit Sub
        v1 = r1.Column - ra.Column
        v2 = r2.Column - ra.Column
        '
        ra.Formula2R1C1 = _
            "=IF(OFFSET(RC,0," & v1 & ")=OFFSET(RC,0," & v2 & "),""〇"","""")"
    End Select
End Sub


