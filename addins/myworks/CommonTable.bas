Attribute VB_Name = "CommonTable"
'==================================
'テーブル共通
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'変数
'----------------------------------------

Private g_columns_margin As Integer
Private g_rows_margin As Integer
Private g_header_color As Long

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
'セル範囲
'----------------------------------------

'領域の角取得
Function LeftTop(ra As Range) As Range
    Set LeftTop = ra(1, 1)
End Function

Function RightTop(ra As Range) As Range
    Set RightTop = ra(1, ra.Columns.Count)
End Function

Function LeftBottom(ra As Range) As Range
    Set LeftBottom = ra(ra.Rows.Count, 1)
End Function

Function RightBottom(ra As Range) As Range
    Set RightBottom = ra(ra.Rows.Count, ra.Columns.Count)
End Function

'----------------------------------------
'テーブル範囲
'----------------------------------------

'テーブル先頭取得
Function TableLeftTop(ByVal ra As Range, Optional n As Long = 0) As Range
    Dim ce As Range
    Set ra = ra.Cells(1, 1)
    Do
        Set ce = ra
        Set ra = FarTop(FarLeft(ra))
    Loop While ce.Address <> ra.Address
    Set TableLeftTop = ce
    '
    Dim i As Long
    For i = 1 To n
        If ce = "" Then
            Set ce = ce.Offset(1)
        ElseIf ce.Offset(0, 1) = "" Then
            Set ce = ce.Offset(1)
        End If
    Next i
    If ce <> "" Then Set TableLeftTop = ce
End Function

'上端取得
Function FarTop(ByVal ra As Range) As Range
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
Function FarLeft(ByVal ra As Range) As Range
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Column
    Dim cnt As Integer
    Dim rs As Range
    Set rs = ce
    If g_columns_margin < 1 Then g_columns_margin = 1
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
Function FarRight(ByVal ra As Range) As Range
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

'テーブル見出し取得
Function TableHeaderRange(ra As Range) As Range
    If ra.Columns.Count <> 1 Then
        Set TableHeaderRange = ra.Rows(1)
    ElseIf ra.Cells(1, 1).Value = "" Then
        Set TableHeaderRange = ra.Cells(1, 1)
    Else
        Set TableHeaderRange = TableRow(ra)
    End If
End Function

'テーブルのデータ領域取得
Function TableDataRange(ra As Range) As Range
    Set TableDataRange = TableRange(TableHeaderRange(ra).Offset(1))
End Function

'テーブル領域取得
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
'セル操作
'----------------------------------------

'文字列が一致するcellを探す
Function FindCell(s As String, Optional ByVal ra As Range) As Range
    If ra Is Nothing Then Set ra = ActiveCell
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    
    Dim ce As Range
    Set ce = ws.UsedRange
    Dim r As Long, c As Long
    r = ra.Row
    c = ra.Column
    If r < ce.Row Then r = ce.Row
    If c < ce.Column Then c = ce.Column
    Set ce = ws.Range(ws.Cells(r, c), ce(ce.Rows.Count, ce.Columns.Count))
    If ce.Rows.Count = 1 And ce.Columns.Count = 1 Then
        If ce.Value = s Then Set FindCell = ce
        Exit Function
    End If
    
    Dim arr As Variant
    arr = ce.Value
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If arr(r, c) = s Then
                Set FindCell = ce(r, c)
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
'テーブルヘッダ
'----------------------------------------

'テーブルフィルタ
Sub HeaderFilter(ra As Range)
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    Else
        TableHeaderRange(ra).AutoFilter
    End If
End Sub

'テーブル幅調整
Sub HeaderAutoFit(ra As Range)
    TableRange(TableHeaderRange(ra)).Columns.AutoFit
End Sub

'テーブル枠固定
Sub HeaderFixed(ra As Range)
    Application.ScreenUpdating = False
    If ActiveWindow.FreezePanes Then
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

'テーブルヘッダ色設定
Sub HeaderColor(ra As Range)
    Dim old As Range
    If TypeName(Selection) = "Range" Then Set old = Selection
    '
    TableHeaderRange(ra).Select
    Application.ScreenUpdating = True
    If Application.Dialogs(xlDialogPatterns).Show Then
        g_header_color = Selection.Interior.color
    End If
    '
    If TypeName(old) = "Range" Then old.Select
End Sub

Sub SetHeaderColor(ra As Range)
    If g_header_color = 0 Then HeaderColor ra
    TableHeaderRange(ra).Interior.color = g_header_color
End Sub

Function GetHeaderColor() As Long
    GetHeaderColor = g_header_color
End Function

'ヘッダ配列取得
Function GetHeaderArray(ce As Range, dic As Dictionary) As String()
    Dim hdr() As String
    ReDim hdr(dic.Count)
    Dim c As Long
    For c = 0 To UBound(hdr)
        Dim k As String
        k = ce.Cells(1, 1 + c).Value
        If Not dic.Exists(k) Then Exit For
        hdr(c) = dic(k)(0)
    Next c
    If c = 0 Then Exit Function
    ReDim Preserve hdr(c - 1)
    GetHeaderArray = hdr
End Function


'----------------------------------------
'テーブル操作拡張
'----------------------------------------

'枠線
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

'囲いクリア
Sub WakuClear(ra As Range)
    Dim rb As Range
    Set rb = TableRange(TableHeaderRange(ra))
    '
    rb.Interior.ColorIndex = xlColorIndexNone
    rb.Borders.LineStyle = xlNone
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
    ActiveWindow.FreezePanes = False
End Sub

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
