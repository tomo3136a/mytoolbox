Attribute VB_Name = "TableTemplate"
'==================================
'テンプレート機能
'==================================

Option Explicit
Option Private Module

Sub TemplateProc(id As Long, Optional id2 As Long)
    Select Case id
    Case 1: CopyAddinSheet                  'シート複製
    Case 2: UpdateAddinSheet ActiveSheet    'シート更新
    Case 3: RemoveAddinSheet                'シート削除
    Case 4:
        Select Case id2
        Case 1:
        Case 2:
        Case Else: AddTable                 'テーブル作成
        End Select
    Case 5: LoadCsvTable                    'テーブル読み込み
    Case 6: 'BuildAddin
    Case 7: CopyHeaderAddinSheet            'ヘッダーシート取得
    Case 8: ToggleAddin                     'addins: toggle visible/hidden sheets
    End Select
End Sub

'----------------------------------
'機能
'----------------------------------

'アドインブックからテンプレートシートを複製
Private Sub CopyAddinSheet()
    Dim ws As Worksheet
    Set ws = SelectSheet(ThisWorkbook, "^[^#]")
    If ws Is Nothing Then Exit Sub
    
    Dim msg As String
    msg = "作成するシート名を入れてください。"
    Dim s As String
    s = InputBox(msg, app_name, ws.name)
    If StrPtr(s) = 0 Then Exit Sub
    
    If s = "" Then s = ws.name
    ws.Copy After:=ActiveSheet
    ActiveSheet.name = UniqueSheetName(ActiveWorkbook, s)
End Sub

Private Sub CopyHeaderAddinSheet()
    Dim s As String
    s = "#header"
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(s)
    On Error GoTo 0
    If ws Is Nothing Then
        ActiveWorkbook.Sheets.Add
        ActiveSheet.name = s
        Exit Sub
    End If
    
    ws.Copy After:=ActiveSheet
    'ActiveSheet.name = UniqueSheetName(ActiveWorkbook, s)
End Sub

'アドインブックへテンプレートシート更新
Private Sub UpdateAddinSheet(ws As Worksheet)
    '名前が登録されているか確認
    Dim ws2 As Worksheet
    For Each ws2 In ThisWorkbook.Sheets
        If ws2.name = ws.name Then Exit For
    Next ws2
    
    '名前が登録されていなければ新規に追加
    If ws2 Is Nothing Then
        ScreenUpdateOff
        ThisWorkbook.IsAddin = False
        ws.Copy After:=ThisWorkbook.Sheets(1)
        ThisWorkbook.IsAddin = True
        ScreenUpdateOn
        Exit Sub
    End If
        
    '名前が登録されていれば上書きコピー
    Dim msg As String
    msg = "同名の登録があります。" & vbLf & ws.name
    msg = msg & vbLf & "上書きしますか。"
    Dim res As VbMsgBoxResult
    res = MsgBox(msg, vbYesNo Or vbDefaultButton2, app_name)
    If res = vbNo Then Exit Sub
    
    Dim old As Range
    Set old = Selection
    ws.Cells.Select
    Selection.Copy
    Set ws2 = ThisWorkbook.Sheets(ws.name)
    ws2.Paste ws2.Cells(1, 1)
    Application.CutCopyMode = False
    old.Select
End Sub

'アドインブックからテンプレートシートを削除
Private Sub RemoveAddinSheet()
    Dim ws As Worksheet
    Set ws = SelectSheet(ThisWorkbook)
    If ws Is Nothing Then Exit Sub
    
    Dim msg As String
    msg = "テンプレートを削除しますか。" & vbLf & ws.name
    Dim res As VbMsgBoxResult
    res = MsgBox(msg, vbYesNo Or vbDefaultButton2, app_name)
    If res = vbNo Then Exit Sub
    
    Dim f As Boolean
    f = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = f
End Sub

'アドインブック表示トグル
Private Sub ToggleAddin()
    If ThisWorkbook.IsAddin Then
        ThisWorkbook.IsAddin = False
        ThisWorkbook.Activate
    Else
        ThisWorkbook.IsAddin = True
        ThisWorkbook.Save
    End If
End Sub

'テーブル作成機能
Private Sub AddTable()
    'ヘッダ定義シート取得
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets.Item("#header")
    On Error GoTo 0
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets.Item("#header")
        On Error GoTo 0
    End If
    If ws Is Nothing Then Exit Sub
    
    'データ範囲取得
    Dim ra As Range
    Set ra = ws.UsedRange
    Set ra = ws.Range(ra.Cells(1, 1), ra.Cells(ra.Rows.Count, 1))
    Set ra = SectionRange(ra)
    If ra Is Nothing Then Exit Sub
    Set ra = SelectCell(ra)
    If ra Is Nothing Then Exit Sub
    If ra.Count <> 1 Then Exit Sub
    Set ra = ra.Offset(0, 1)
    '
    'ScreenUpdateOff
    '
    Dim cm As Long, c As Long
    c = ra.Column
    cm = ws.UsedRange.Column + ws.UsedRange.Columns.Count
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
Sub LoadCsvTable(Optional path As String, Optional utf8 As Boolean)
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

'----------------------------------
'その他
'----------------------------------

Sub LoadTable()
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

