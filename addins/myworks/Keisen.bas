Attribute VB_Name = "Keisen"
'==================================
'Œrü˜g‘€ì
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'‹@”\ŒÄ‚Ño‚µ
'----------------------------------------

'ƒe[ƒuƒ‹‘I‘ğ
Sub SelectTable(mode As Integer, ra As Range)
    On Error Resume Next
    Select Case mode
    Case 0: TableLeftTop(ra).Select
    Case 1: LeftBottom(TableRange(TableHeaderRange(TableLeftTop(ra)))).Offset(1).Select
    Case 2: Intersect(TableRange(TableHeaderRange(TableLeftTop(ra))), ra.EntireRow).Select
    Case 3: Intersect(TableRange(TableHeaderRange(TableLeftTop(ra))), ra.EntireColumn).Select
    Case 4: TableRange(TableHeaderRange(TableLeftTop(ra))).Select
    Case 5: TableHeaderRange(ra).Select
    End Select
    On Error GoTo 0
End Sub

'Œrü˜g
Sub TableWaku(mode As Integer, ra As Range)
    Select Case mode
    Case 0: Waku TableLeftTop(ra), fit:=True
    Case 1: Waku ra
    Case 3: HeaderFilter ra
    Case 4: HeaderAutoFit ra
    Case 5: HeaderFixed ra
    Case 6: HeaderColor ra
    Case 7: WakuClear TableLeftTop(ra)
    Case 8: TableRange(TableHeaderRange(TableLeftTop(ra)).Offset(1)).Clear
    Case 9: TableRange(TableHeaderRange(TableLeftTop(ra))).Clear
    End Select
End Sub

'—ñ’Ç‰Á
Sub AddColumn(mode As Integer, ra As Range)
    Select Case mode
    Case 1: AddLineNo ra
    End Select
End Sub


'----------------------------------------
'API
'----------------------------------------

'ˆÍ‚¢
Sub Waku(ra As Range, _
        Optional filter As Boolean, _
        Optional fit As Boolean, _
        Optional fixed As Boolean, _
        Optional color As Integer = 15 _
    )
    Dim rh As Range
    Set rh = TableHeaderRange(ra)
    If rh.Cells(1, 1).Value = "" Then Exit Sub
    Dim rb As Range
    Set rb = TableRange(rh)
    '
    If GetHeaderColor = 0 Then
        rh.Interior.ColorIndex = color
    Else
        rh.Interior.color = GetHeaderColor
    End If
    Call WakuBorder(rb)
    '
    If filter Then HeaderFilter rh
    If fit Then rb.Columns.AutoFit
End Sub

'----------------------------------------
'”Ô†—ñ’Ç‰Á
'----------------------------------------

Private Function IsLineNo(ra As Range) As Boolean
    Dim s As String
    s = ra.Cells(1, 1).Value
    If s = "No." Or s = "#" Or s = "”Ô†" Then IsLineNo = True
End Function

Sub AddLineNo(ra As Range)
    Dim rs As Range
    Set rs = TableRange(ra)
    rs.EntireColumn.Insert shift:=xlShiftToRight
    Set rs = rs.Offset(0, -1)
    rs(1, 1).Value = "No."
    Set rs = Intersect(rs, rs.Offset(1))
    
    Dim i As Integer
    Dim ce As Range
    For Each ce In rs
        i = i + 1
        ce.Value = i
    Next ce
    rs.EntireColumn.Columns.AutoFit
    Call Waku(ra)
End Sub

