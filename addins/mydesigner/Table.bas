Attribute VB_Name = "Table"
Option Explicit
Option Private Module

'==================================
'テーブル共通
'==================================

'----------------------------------------
'変数
'----------------------------------------

Public g_columns_margin As Integer
Public g_rows_margin As Integer
Public g_header_color As Long

'----------------------------------------
'マージン
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
    If g_rows_margin < 2 Then g_rows_margin = 1
    If g_columns_margin < 2 Then g_columns_margin = 1
    If mode = xlRows Then v = g_rows_margin
    If mode = xlColumns Then v = g_columns_margin
    If v < 1 Or v > 9 Then v = 1
    GetTableMargin = v
End Function


'----------------------------------------
'範囲取得
'----------------------------------------

'列見出し取得
Public Function HeaderRange(ra As Range) As Range
    If ra.Columns.Count <> 1 Then
        Set HeaderRange = ra.Rows(1)
        Exit Function
    ElseIf ra.Cells(1, 1).Value = "" Then
        Set HeaderRange = ra.Cells(1, 1)
        Exit Function
    End If
    '
    Dim rs As Range
    Set rs = FarLeft(ra, g_columns_margin)
    Dim re As Range
    Set re = FarRight(ra, g_rows_margin)
    Set HeaderRange = ra.Worksheet.Range(rs, re)
End Function

'表取得
Public Function TableRange(ra As Range) As Range
    Dim rh As Range
    Set rh = ra
    '
    Dim rc As Long
    rc = ra.Rows.Count - 1
    If rc = 0 Then
        Dim rs As Range
        Set rs = FarLeft(ra, g_columns_margin)
        rc = FarBottom(rs, g_rows_margin).Row - rs.Row
    End If
    '
    Set TableRange = ra.Worksheet.Range(rh, rh.Offset(rc))
End Function


'----------------------------------------
'テーブル読み込み
'----------------------------------------

'テキストファイル読み込み(シート追加)
Public Function AddTextSheet( _
        path As String, _
        Optional space As Boolean = True, _
        Optional comma As Boolean, _
        Optional utf8 As Boolean) As Worksheet
    If Not fso.FileExists(path) Then Exit Function
    '
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.name = UniqueSheetName(ws.Parent, fso.GetFileName(path))
    '
    Call ReadText(ws.Cells(1, 1), path, space, comma, utf8)
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

