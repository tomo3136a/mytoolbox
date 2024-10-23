Attribute VB_Name = "CommonTable"
'==================================
'テーブル共通
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'変数
'----------------------------------------

Public g_columns_margin As Integer
Public g_rows_margin As Integer
Public g_header_color As Long

'----------------------------------------
'表操作マージン
'----------------------------------------

Sub SetTableMargin(Optional mode As Integer = xlColumns, Optional v As Integer)
    If v < 1 Then
        Dim s As String
        If mode = xlRows Then v = g_rows_margin: s = "行"
        If mode = xlColumns Then v = g_columns_margin: s = "列"
        If v < 1 Then v = 1
        s = s + "マージンを入力してください(1～9)"
        v = Application.InputBox(s, Type:=1, Default:=v)
    End If
    If v < 1 Or v > 9 Then
    Else
        If mode = xlRows Then g_rows_margin = v
        If mode = xlColumns Then g_columns_margin = v
    End If
End Sub

Function GetTableMargin(Optional mode As Integer = xlColumns) As Integer
    Dim v As Integer
    If g_rows_margin < 1 Then g_rows_margin = 1
    If g_columns_margin < 1 Then g_columns_margin = 1
    If mode = xlRows Then v = g_rows_margin
    If mode = xlColumns Then v = g_columns_margin
    If v < 1 Or v > 9 Then v = 1
    GetTableMargin = v
End Function

'----------------------------------------
'セル操作
'----------------------------------------

'領域の角取得
Function LeftTop(ra As Range) As Range
    Set LeftTop = ra.Parent.Cells(ra.Row, ra.Column)
End Function

Function RightTop(ra As Range) As Range
    Set RightTop = ra.Parent.Cells(ra.Row, ra.Column + ra.Columns.Count - 1)
End Function

Function LeftBottom(ra As Range) As Range
    Set LeftBottom = ra.Parent.Cells(ra.Row + ra.Rows.Count - 1, ra.Column)
End Function

Function RightBottom(ra As Range) As Range
    Set RightBottom = ra.Parent.Cells(ra.Row + ra.Rows.Count - 1, ra.Column + ra.Columns.Count - 1)
End Function

'文字列が一致するcellを探す
Function FindCell(ra As Range, s As String) As Range
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Dim r As Long
    Dim c As Integer
    For r = ra.Row To ws.UsedRange.Rows.Count
        For c = ra.Column To ws.UsedRange.Columns.Count
            Dim ce As Range
            Set ce = ws.Cells(r, c)
            If ce.Value = s Then
                Set FindCell = ce
                Exit Function
            End If
        Next c
    Next r
End Function

'ブランクをスキップする
Public Function SkipBlankCell(ra As Range) As Range
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Dim r As Long
    Dim c As Integer
    For r = ra.Row To ws.UsedRange.Rows.Count
        For c = ra.Column To ws.UsedRange.Columns.Count
            Dim ce As Range
            Set ce = ws.Cells(r, c)
            If ce.Value <> "" Then
                Set SkipBlankCell = ce
                Exit Function
            End If
        Next c
    Next r
End Function

'----------------------------------------
'表の範囲取得
'----------------------------------------

'テーブル先頭取得
Function FarLeftTop(ra As Range) As Range
    Dim rs As Range, ce As Range
    Set rs = FarLeft(ra)
    Set rs = FarTop(rs)
    Set rs = FarLeft(rs)
    Set ce = rs.Cells(1, 1)
    Set FarLeftTop = ce
    '
    Dim i As Integer
    For i = 1 To 5
        If ce = "" Then
            Set ce = ce.Offset(1)
        ElseIf ce.Offset(0, 1) = "" Then
            Set ce = ce.Offset(1)
        End If
    Next i
    'If ce <> "" Then ce.Select
    If ce <> "" Then Set FarLeftTop = ce
End Function

'上端取得
Function FarTop(ra As Range) As Range
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Row
    Dim cnt As Integer
    Dim rs As Range
    Set rs = ce
    If g_rows_margin < 1 Then g_rows_margin = 1
    Do While ce.Row > p And cnt < g_rows_margin
        If ce.Offset(-1).Value = "" Then
            Set ce = ce.Offset(-1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlUp)
            Set rs = ce
            cnt = 0
        End If
    Loop
    Set FarTop = rs
End Function

'左端取得
Function FarLeft(ra As Range) As Range
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Column
    Dim cnt As Integer
    Dim rs As Range
    Set rs = ce
    Do While ce.Column > p And cnt < g_columns_margin
        If ce.Offset(0, -1).Value = "" Then
            Set ce = ce.Offset(0, -1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlToLeft)
            Set rs = ce
            cnt = 0
        End If
    Loop
    Set FarLeft = rs
End Function

'右端取得
Function FarRight(ra As Range) As Range
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Column
    p = p + ra.Worksheet.UsedRange.Columns.Count
    Dim cnt As Integer
    Dim re As Range
    Set re = ce
     If g_columns_margin < 1 Then g_columns_margin = 1
   Do While ce.Column < p And cnt < g_columns_margin
        If ce.Offset(0, 1).Value = "" Then
            Set ce = ce.Offset(0, 1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlToRight)
            Set re = ce
            cnt = 0
        End If
    Loop
    Set FarRight = re
End Function

'下端取得
Function FarBottom(ra As Range) As Range
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Row
    p = p + ra.Worksheet.UsedRange.Rows.Count
    Dim cnt As Integer
    Dim re As Range
    Set re = ce
    If g_rows_margin < 1 Then g_rows_margin = 1
    Do While ce.Row < p And cnt < g_rows_margin
        'Set ce = ce.Offset(1)
        If ce.Offset(1).Value = "" Then
            Set ce = ce.Offset(1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlDown)
            Set re = ce
            cnt = 0
        End If
    Loop
    Set FarBottom = re
End Function

'テーブル行取得
Function TableRow(ra As Range) As Range
    Dim rs As Range
    Set rs = FarLeft(ra)
    Dim re As Range
    Set re = FarRight(ra)
    Set TableRow = ra.Worksheet.Range(rs, re.Offset(ra.Rows.Count - 1))
End Function

'テーブル列取得
Function TableColumn(ra As Range) As Range
    Dim rs As Range
    Set rs = FarTop(ra)
    Dim re As Range
    Set re = FarBottom(ra)
    Set TableColumn = ra.Worksheet.Range(rs, re.Offset(, ra.Columns.Count - 1))
End Function

'列見出し取得
Function HeaderRange(ra As Range) As Range
    If ra.Columns.Count <> 1 Then
        Set HeaderRange = ra.Rows(1)
    ElseIf ra.Cells(1, 1).Value = "" Then
        Set HeaderRange = ra.Cells(1, 1)
    Else
        Set HeaderRange = TableRow(ra)
    End If
End Function

'表取得
Function TableRange(ra As Range) As Range
    Dim rh As Range
    Set rh = ra
    '
    Dim rc As Long
    rc = ra.Rows.Count - 1
    If rc = 0 Then
        Dim rs As Range
        Set rs = FarLeft(ra)
        rc = FarBottom(rs).Row - rs.Row
    End If
    '
    Set TableRange = ra.Worksheet.Range(rh, rh.Offset(rc))
End Function

'----------------------------------------
'テーブル読み込み
'----------------------------------------

'テキストファイル読み込み
Function AddTextSheet(path As String) As Worksheet
    Dim ws_old As Worksheet
    Set ws_old = ActiveSheet
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.name = UniqueSheetName(ws.Parent, fso.GetFileName(path))
    '
    Call ReadText(ws.Cells(1, 1), path)
    '
    Set AddTextSheet = ws
End Function

'テキストファイル読み込み
Public Sub ReadText( _
        ra As Range, _
        path As String, _
        Optional space As Boolean = True, _
        Optional comma As Boolean, _
        Optional utf8 As Boolean)
    If Not fso.FileExists(path) Then Exit Sub
    '
    Dim arrDataType(255) As Long
    Dim i As Long
    For i = 0 To 255
        arrDataType(i) = xlTextFormat
    Next i
    '
    Dim enc As Integer
    enc = 932
    If utf8 Then enc = 65001
    '
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    With ws.QueryTables.Add( _
            Connection:="TEXT;" + path, _
            Destination:=ra)
        .TextFileParseType = xlDelimited
        .TextFileSpaceDelimiter = space
        .TextFileCommaDelimiter = comma
        .TextFilePlatform = enc
        .TextFileStartRow = 1
        .TextFileColumnDataTypes = arrDataType
        .Refresh BackgroundQuery:=False
        .name = "tmp"
        .Delete
    End With
    '
    Dim na As Variant
    For Each na In ws.Parent.Names
        If na.name = ws.name & "!" & "tmp" Then na.Delete
    Next na
End Sub
