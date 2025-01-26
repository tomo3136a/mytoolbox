Attribute VB_Name = "Keisen"
'==================================
'årê¸ògëÄçÏ
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'ã@î\åƒÇ—èoÇµ
'----------------------------------------

'ÉeÅ[ÉuÉãëIë
Sub SelectTable(mode As Integer, ra As Range)
    Select Case mode
    Case 1
        TableRow(ra).Select
    Case 2
        Intersect(TableRange(TableHeaderRange(TableLeftTop(ra))), TableColumn(ra)).Select
    Case 3
        TableRange(TableHeaderRange(TableLeftTop(ra))).Select
    Case Else
         TableLeftTop(ra).Select
    End Select
End Sub

'årê¸òg
Sub TableWaku(mode As Integer, ra As Range)
    Select Case mode
    Case 1
        Waku ra
    Case 2
        AddLineNo ra
    Case 3
        HeaderFilter ra
    Case 4
        HeaderAutoFit ra
    Case 5
        HeaderFixed ra
    Case 6
        HeaderColor ra
    Case 7
        WakuClear TableLeftTop(ra)
    Case 8
        TableRange(TableHeaderRange(TableLeftTop(ra)).Offset(1)).Clear
    Case 9
        TableRange(TableHeaderRange(TableLeftTop(ra))).Clear
    Case Else
        Waku TableLeftTop(ra), fit:=True
    End Select
End Sub

'----------------------------------------
'API
'----------------------------------------

'àÕÇ¢
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
'î‘çÜóÒí«â¡
'----------------------------------------

Private Function IsLineNo(ra As Range) As Boolean
    Dim s As String
    s = ra.Cells(1, 1).Value
    If s = "No." Or s = "#" Or s = "î‘çÜ" Then IsLineNo = True
End Function

Sub AddLineNo(ra As Range)
    Dim rs As Range
    Set rs = FarLeft(ra)
    If IsLineNo(rs) Then
        Set rs = rs.Offset(1, 1)
        Set rs = Range(rs, FarBottom(rs))
    Else
        Set rs = TableRange(rs)
        rs.EntireColumn.Insert shift:=xlShiftToRight
        rs(1, 1).Offset(0, -1).Value = "No."
        Set rs = rs.Offset(1)
        Set rs = Range(rs, FarBottom(rs))
    End If
    '
    Dim i As Integer
    Dim ce As Range
    For Each ce In rs
        If ce.Value <> "" Then
            i = i + 1
            ce.Offset(0, -1).Value = i
        Else
            ce.Offset(0, -1).Value = ""
        End If
    Next ce
    Call Waku(ra)
End Sub

