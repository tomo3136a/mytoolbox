Attribute VB_Name = "MIDF"
Option Explicit
Option Private Module

Private g_idf_path As String

'----------------------------------------
'IDF作図機能
'----------------------------------------
'IDFファイル読み込み
Public Sub ImportIDF( _
        Optional ce As Boolean, _
        Optional enc As Integer = 932)
    '読み込みファイル選択
    Dim path As String
    path = g_idf_path
    If path = "" Then path = ActiveWorkbook.path
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "IDFファイル", "*.emn"
        .Filters.Add "ライブラリファイル", "*.emp"
        .InitialFileName = path & "\"
        .AllowMultiSelect = False
        If .Show = True Then
            path = .SelectedItems(1)
        End If
    End With
    g_idf_path = fso.GetParentFolderName(path)
    If Not fso.FileExists(path) Then Exit Sub
    '
    '画面チラつき防止処置
    ScreenUpdateOff
    '
    'ワークシート作成
    Dim ws_old As Worksheet
    Set ws_old = ActiveSheet
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.name = UniqueSheetName(ws.Parent, fso.GetFileName(path))
    '
    'テキストファイル読み込み前処理
    Dim arrDataType(255) As Long
    Dim i As Long
    For i = 0 To 255
        arrDataType(i) = xlTextFormat
    Next i
    '
    'テキストファイル読み込み
    With ws.QueryTables.Add( _
            Connection:="TEXT;" + path, _
            Destination:=ws.Cells(1, 1))
        .TextFilePlatform = enc
        .TextFileStartRow = 1
        .TextFileColumnDataTypes = arrDataType
        .Refresh BackgroundQuery:=False
        .name = "tmp"
        .Delete
    End With
    '
    'テキストファイル読み込み後処理
    Dim na As Variant
    For Each na In ws.Parent.Names
        If na.name = ws.name & "!" & "tmp" Then na.Delete
    Next na
    '
    'テキストを空白で区切る
    ws.Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, comma:=False, space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)), _
        TrailingMinusNumbers:=True
    ws.Columns("A:A").EntireColumn.AutoFit
    '
    '編集前のワークシートを表示
    ws_old.Activate
    Set ws_old = Nothing
    '
    If ce Then
        ActiveCell.Value = ws.name
        ActiveCell.Offset(1).Select
    End If
    Set ws = Nothing
    '
    '画面チラつき防止処置解除
    ScreenUpdateOn
End Sub

'IDFファイル書き出し
Public Sub ExportIDF(ws As Worksheet)
    '書き出しファイル選択
    Dim path As String
    path = g_idf_path
    If path = "" Then path = ActiveWorkbook.path
    Dim name As String
    name = re_replace(ws.name, "\s*\(\d+\)$", "")
    path = fso.BuildPath(path, name)
    Dim idx As Integer
    If LCase(Right(path, 4)) = ".emn" Then idx = 1
    If LCase(Right(path, 4)) = ".emp" Then idx = 2
    Dim flt As String
    flt = "IDFファイル,*.emn,ライブラリファイル,*.emp"
    path = Application.GetSaveAsFilename(path, flt, idx)
    If path = "False" Then Exit Sub
    g_idf_path = fso.GetParentFolderName(path)
    '
    Dim ra As Range
    Set ra = ws.UsedRange
    Dim r As Long
    Dim c As Long
    Dim n As Long
    Dim sect As String
    Dim line As String
    Open path For Output As #1
    For r = 1 To ra.Rows.Count
        line = ""
        Dim s0 As String
        Dim s1 As String
        s0 = Trim(ra(r, 1).Value)
        If Left(s0, 1) = "." Then
            sect = s0
            n = 0
        End If
        If s0 = "" Then line = "  "
        For c = 1 To ra.Columns.Count
            s1 = Trim(ra(r, c).Value)
            If InStr(s1, " ") Then s1 = Chr(34) & s1 & Chr(34)
            If Trim(s1) = "" Then s1 = "   "
            line = line + s1 + "  "
        Next c
        Print #1, RTrim(line)
        n = n + 1
    Next r
    Close #1
End Sub

'IDFファイル描画
Public Function DrawIDF( _
        ws As Worksheet, x As Double, y As Double, _
        Optional path As String, Optional g As Double, _
        Optional sheet_load As Boolean) As String
    '
    Dim idf As CIDF
    Set idf = New CIDF
    If sheet_load Then
        If Not idf.LoadSheet(path) Then Exit Function
    Else
        If Not idf.LoadFile(path) Then Exit Function
    End If
    If idf.Count = 0 Then Exit Function
    '
    '画面チラつき防止処置
    ScreenUpdateOff
    '
    If g = 0 Then g = GetDrawParam(2)
    Dim w As Double
    Dim h As Double
    w = g * idf.Width
    h = g * idf.Height
    If IsDrawParam(4) Then
        Dim sh As Object
        Set sh = ws.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
        sh.Fill.ForeColor.RGB = RGB(128, 255, 128)
    End If
    '
    Dim x0 As Double
    Dim y0 As Double
    x0 = x - g * idf.Left
    y0 = y + g * idf.Bottom
    '
    Dim s As String
    s = idf.Draw(ws, x0, y0, 0#, g)
    If IsDrawParam(5) Then Call DrawAxis(ws, x0, y0, w, h)
    '
    '画面チラつき防止処置解除
    ScreenUpdateOn
End Function

'IDF部品ファイル読み込み
Public Function DrawIDF2( _
        ws As Worksheet, x As Double, y As Double, _
        Optional path As String, Optional g As Double) As String
    '
    Dim idf As CIDF
    Set idf = New CIDF
    If Not idf.LoadSheet() Then Exit Function
    If idf.Count = 0 Then Exit Function
    '
    '画面チラつき防止処置
    ScreenUpdateOff
    '
    If g = 0 Then g = GetDrawParam(2)
    Dim w As Double
    Dim h As Double
    w = g * idf.Width
    h = g * idf.Height
    If IsDrawParam(4) Then
        Dim sh As Object
        Set sh = ws.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
        sh.Fill.ForeColor.RGB = RGB(128, 255, 128)
    End If
    '
    Dim x0 As Double
    Dim y0 As Double
    x0 = x - g * idf.Left
    y0 = y + g * idf.Bottom
    '
    Dim s As String
    s = idf.Draw(ws, x0, y0, 0#, g)
    If IsDrawParam(5) Then Call DrawAxis(ws, x0, y0, w, h)
    '
    '画面チラつき防止処置解除
    ScreenUpdateOn
End Function

