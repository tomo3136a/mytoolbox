Attribute VB_Name = "Keisen"
'==================================
'årê¸ògëÄçÏ
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'menu
'----------------------------------------

'ÉeÅ[ÉuÉãëIë
Sub SelectTable(mode As Integer, ra As Range)
    Select Case mode
    Case 1
        TableRow(ra).Select
    Case 2
        TableColumn(ra).Select
    Case 3
        TableRange(HeaderRange(ra)).Select
    Case 11
        HeaderRange(ra).Select
    Case 12
        TableRange(ra).Select
    Case Else
         FarLeftTop ra
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
        WakuClear ra
    Case 8
        TableRange(HeaderRange(ra)).Clear
    Case 9
        FarLeftTop ra
        TableRange(HeaderRange(ra).Offset(1)).Clear
    Case Else
        FarLeftTop ra
        Waku ra, fit:=True
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
    Set rh = HeaderRange(ra)
    If rh.Cells(1, 1).Value = "" Then Exit Sub
    Dim rb As Range
    Set rb = TableRange(rh)
    '
    If g_header_color = 0 Then
        rh.Interior.ColorIndex = color
    Else
        rh.Interior.color = g_header_color
    End If
    Call WakuBorder(rb)
    '
    If filter Then HeaderFilter rh
    If fit Then rb.Columns.AutoFit
End Sub

'ògê¸
Sub WakuBorder(ra As Range)
    ra.Borders.LineStyle = xlContinuous
    Dim c As Integer
    Dim r As Integer
    If g_columns_margin > 1 Then
        r = ra.Rows.Count
        For c = 1 To ra.Columns.Count
            If ra.Cells(1, c).Value = "" Then
                Dim rc As Range
                Set rc = Range(ra.Cells(1, c), ra.Cells(r, c))
                rc.Borders(xlEdgeLeft).LineStyle = xlNone
            End If
        Next c
    End If
    If g_rows_margin > 1 Then
        c = ra.Columns.Count
        For r = 1 To ra.Rows.Count
            If ra.Cells(r, 1).Value = "" Then
                Dim rr As Range
                Set rr = Range(ra.Cells(r, 1), ra.Cells(r, c))
                rr.Borders(xlEdgeTop).LineStyle = xlNone
            End If
        Next r
    End If
End Sub

'ÉtÉBÉãÉ^
Sub HeaderFilter(ra As Range)
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    Else
        HeaderRange(ra).AutoFilter
    End If
End Sub

'ïùí≤êÆ
Sub HeaderAutoFit(ra As Range)
    TableRange(HeaderRange(ra)).Columns.AutoFit
End Sub

'ògå≈íË
Sub HeaderFixed(ra As Range)
    Application.ScreenUpdating = False
    If ActiveWindow.FreezePanes Then
        'Application.ScreenUpdating = True
        ActiveWindow.FreezePanes = False
        Exit Sub
    End If
    '
    Dim old As Range: Set old = Selection
    '
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    If ce.Column = 1 Then
        ce.Offset(1).EntireRow.Select
    ElseIf ce.Offset(0, -1).Value = "" Then
        ce.Offset(1).EntireRow.Select
    Else
        ce.Offset(1).Select
    End If
    Application.ScreenUpdating = True
    ActiveWindow.FreezePanes = True
    '
    old.Select
End Sub

'ÉwÉbÉ_êFê›íË
Sub HeaderColor(ra As Range)
    Dim old As Range
    Set old = Selection
    '
    HeaderRange(ra).Select
    Application.ScreenUpdating = True
    If Application.Dialogs(xlDialogPatterns).Show Then
        g_header_color = Selection.Interior.color
    End If
    '
    old.Select
End Sub

'àÕÇ¢ÉNÉäÉA
Sub WakuClear(ra As Range)
    Dim rb As Range
    Set rb = TableRange(HeaderRange(ra))
    '
    rb.Interior.ColorIndex = xlColorIndexNone
    rb.Borders.LineStyle = xlNone
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
    ActiveWindow.FreezePanes = False
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

