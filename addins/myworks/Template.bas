Attribute VB_Name = "Template"
'==================================
'テンプレート機能
'==================================

Option Explicit
Option Private Module

' mode=1: シート追加
'      2: シート登録
'      3: シート削除
'      4: テーブル作成
'      5: テーブル読み込み
'      6: テーブル更新
'      7: ヘッダーシート取得
'      8: addins: toggle visible/hidden sheets
Sub TemplateProc(mode As Long, Optional id As Long)
    Select Case mode
    Case 1: CopyAddinSheet
    Case 2: UpdateAddinSheet ActiveSheet
    Case 3: RemoveAddinSheet
    Case 4: AddTable
    Case 5: LoadCsvTable
    Case 6: 'BuildAddin
    Case 7: CopyHeaderAddinSheet
    Case 8: ToggleAddinBook
    End Select
End Sub

'----------------------------------
'機能
'----------------------------------

'アドインブックからテンプレートシートを複製
Private Sub CopyAddinSheet(Optional src As String, Optional dst As String)
    Dim ws As Worksheet
    If src = "" Then
        Set ws = SelectSheet(ThisWorkbook, "^[^#]")
    Else
        Set ws = SearchName(ThisWorkbook.Sheets, src)
    End If
    If ws Is Nothing Then Exit Sub
    '
    Dim s As String
    s = dst
    If s = "" Then
        Dim msg As String
        msg = "作成するシート名を入れてください。"
        s = InputBox(msg, app_name, ws.name)
        If StrPtr(s) = 0 Then Exit Sub
        If s = "" Then s = ws.name
    End If
    '
    s = UniqueSheetName(ActiveWorkbook, s)
    ws.Copy After:=ActiveSheet
    ActiveSheet.name = s
End Sub

'アドインブックへテンプレートシート更新
Private Sub UpdateAddinSheet(ws As Worksheet)
    Dim ws2 As Worksheet
    For Each ws2 In ThisWorkbook.Sheets
        If ws2.name = ws.name Then Exit For
    Next ws2
    '
    '名前が登録されていなければ新規に追加
    If ws2 Is Nothing Then
        ScreenUpdateOff
        ThisWorkbook.IsAddin = False
        ws.Copy After:=ThisWorkbook.Sheets(1)
        ThisWorkbook.IsAddin = True
        ScreenUpdateOn
        Exit Sub
    End If
    '
    '上書き確認
    Dim msg As String
    msg = "同名の登録があります。" & vbLf & ws.name
    msg = msg & vbLf & "上書きしますか。"
    Dim res As VbMsgBoxResult
    res = MsgBox(msg, vbYesNo Or vbDefaultButton2, app_name)
    If res = vbNo Then Exit Sub
    '
    '名前が登録されていれば上書きコピー
    Dim old As Range
    Set old = Selection
    ws.Cells.Copy
    With ThisWorkbook.Sheets(ws.name)
        .Paste .Cells(1, 1)
    End With
    Application.CutCopyMode = False
    old.Select
End Sub

'アドインブックからテンプレートシートを削除
Private Sub RemoveAddinSheet()
    Dim ws As Worksheet
    Set ws = SelectSheet(ThisWorkbook)
    If ws Is Nothing Then Exit Sub
    '
    Dim msg As String
    msg = "テンプレートを削除しますか。" & vbLf & ws.name
    Dim res As VbMsgBoxResult
    res = MsgBox(msg, vbYesNo Or vbDefaultButton2, app_name)
    If res = vbNo Then Exit Sub
    '
    Dim f As Boolean
    f = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = f
End Sub

'テーブル作成機能
Private Sub AddTable()
    'ヘッダ定義シート取得
    Dim ws As Worksheet
    Set ws = SearchName(ActiveWorkbook.Sheets, "#header")
    If ws Is Nothing Then Set ws = SearchName(ThisWorkbook.Sheets, "#header")
    If ws Is Nothing Then Exit Sub
    
    'データ範囲取得
    Dim ra As Range
    Set ra = SectionRange(ws.UsedRange.Columns(1))
    If ra Is Nothing Then Exit Sub
    'Set ra = ws.UsedRange
    'Set ra = ra.Columns(1)
    'Set ra = ws.Range(ra.Cells(1, 1), ra.Cells(ra.Rows.Count, 1))
    'Set ra = SectionRange(ra)
    Set ra = SelectCell(ra)
    If ra Is Nothing Then Exit Sub
    If ra.Count <> 1 Then Exit Sub
    Set ra = ra.Offset(1, 1)
    '
    'ScreenUpdateOff
    '
    Dim cm As Long, c As Long
    cm = ws.UsedRange.Column + ws.UsedRange.Columns.Count
    c = ra.Column
    '
    Dim rm As Long, r As Long
    rm = ra.Row
    If ra.Offset(1).Value <> "" Then rm = ra.End(xlDown).Row
    '
    Dim r2 As Long, c2 As Long
    r2 = ra.Row
    For r = r2 To rm
        If Not IsNumeric(ws.Cells(r, 2).Value) Then Exit For
        If ws.Cells(r, 2).Value < 2 Then r2 = r
        c2 = ws.Cells(r, cm).End(xlToLeft).Column
        If c2 > c Then c = c2
    Next r
    r = r - 1
    Set ra = ra.Offset(0, 1)
    '
    Dim tbl As Range
    Set tbl = ws.Range(ra, ws.Cells(r, c))
    tbl.Copy Destination:=Selection.Cells(1, 1)
    Selection.Offset(r2 - ra.Row + 1).Select
    '
    ScreenUpdateOn
End Sub

'テーブル読み込み機能
Private Sub LoadCsvTable(Optional path As String, Optional utf8 As Boolean)
    If path = "" Then path = SelectCsvFile()
    If path = "" Then Exit Sub
    '
    Dim asu As Boolean
    asu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim ra As Range
    Set ra = ActiveCell
    '
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    '
    Dim arrDataType(255) As Long
    Dim i As Integer
    For i = 0 To 255
        arrDataType(i) = xlTextFormat
    Next i
    '
    With ws.QueryTables.Add( _
            Connection:="TEXT;" + path, _
            Destination:=ra)
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        If utf8 Then
            .TextFilePlatform = 65001 'UTF-8
        Else
            .TextFilePlatform = 932   'SJIS
        End If
        .TextFileStartRow = 1
        .TextFileColumnDataTypes = arrDataType
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    '
    Application.ScreenUpdating = asu
End Sub

'ヘッダーシート取得
Private Sub CopyHeaderAddinSheet()
    Dim s As String
    s = "#header"
    '
    Dim ws As Worksheet
    Set ws = SearchName(ActiveWorkbook.Sheets, s)
    If Not ws Is Nothing Then
        ws.Activate
        Exit Sub
    End If
    '
    Set ws = SearchName(ThisWorkbook.Sheets, s)
    If Not ws Is Nothing Then
        CopyAddinSheet s, s
        Exit Sub
    End If
    '
    ActiveWorkbook.Sheets.Add
    ActiveSheet.name = s
End Sub

'アドインブック表示トグル
Private Sub ToggleAddinBook()
    If ThisWorkbook.IsAddin Then
        ThisWorkbook.IsAddin = False
        ThisWorkbook.Activate
    Else
        ThisWorkbook.IsAddin = True
        ThisWorkbook.Save
    End If
End Sub

'----------------------------------
'その他
'----------------------------------

Private Sub LoadTable()
    Dim Title As String
    Title = "データ"
    '
    Dim arrDataType(255) As Long
    Dim i As Integer
    For i = 0 To 255
        arrDataType(i) = xlTextFormat
    Next i
    '
    Dim asu As Boolean
    asu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim v As Variant
    For Each v In SelectFiles(, , "CSVファイル")
        Dim ws As Worksheet
        Set ws = Sheets.Add
        ws.Cells(1, 1) = Title
        ws.Cells(2, 1) = fso.GetBaseName(v)
        ws.name = UniqueSheetName(ws.Parent, fso.GetBaseName(v))
        With ws.QueryTables.Add( _
                Connection:="TEXT;" + v, _
                Destination:=ws.Cells(3, 1))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            '.TextFilePlatform = 932   'SJIS
            .TextFilePlatform = 65001 'UTF-8
            .TextFileStartRow = 1
            .TextFileColumnDataTypes = arrDataType
            .Refresh BackgroundQuery:=False
            '.name = "tmp_tbl"
            .Delete
        End With
        'Dim n As name
        'For Each n In ActiveWorkbook.Names
        '    If n.name = .name & "!" & "tmp_tbl" Then
        '        n.Delete
        '    End If
        'Next n
    Next v
    '
    Application.ScreenUpdating = asu
End Sub

