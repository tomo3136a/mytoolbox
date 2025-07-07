Attribute VB_Name = "Template"
'==================================
'テンプレート機能
'==================================

Option Explicit
Option Private Module

'テンプレート機能
' mode=1: シート追加
'      2: シート登録
'      3: シート削除
'      4: テーブル追加
'      4: テーブル登録
'      6: テーブル削除
'      7: ヘッダーシート取得
'      8: テーブル読み込み
'      9: addins: toggle visible/hidden sheets
Sub TemplateProc(mode As Long, Optional id As Long)
    Select Case mode
    Case 1: CopyTemplateSheet
    Case 2: UpdateTemplateSheet ActiveSheet
    Case 3: RemoveTemplateSheet
    
    Case 4: CopyTemplateTable
    Case 5: UpdateTemplateTable Selection
    Case 6: RemoveTemplateTable
    Case 7: CopyHeaderSheet
    
    Case 8: LoadCsvTable
    Case 9: ToggleAddinBook
    End Select
End Sub

'----------------------------------
'テンプレートシート機能
'----------------------------------

'テンプレートシート複製
Private Sub CopyTemplateSheet( _
    Optional src As String, Optional dst As String)
    Dim ws As Worksheet
    If src <> "" Then Set ws = TakeByName(ThisWorkbook.Sheets, src)
    If ws Is Nothing Then Set ws = SelectSheet(ThisWorkbook, "^[^#]")
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

'テンプレートシート更新
Private Sub UpdateTemplateSheet(ws As Worksheet)
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

'テンプレートシート削除
Private Sub RemoveTemplateSheet( _
    Optional sname As String, Optional bforce As Boolean)
    Dim ws As Worksheet
    Set ws = TakeByName(ThisWorkbook.Worksheets, sname)
    If ws Is Nothing Then Set ws = SelectSheet(ThisWorkbook)
    If ws Is Nothing Then Exit Sub
    '
    '削除確認
    Dim res As VbMsgBoxResult
    Dim msg As String
    If Not bforce Then
        msg = "テンプレートを削除しますか。" & vbLf & ws.name
        res = MsgBox(msg, vbYesNo Or vbDefaultButton2, app_name)
        If res = vbNo Then Exit Sub
    End If
    '
    'シート削除
    Dim f As Boolean
    f = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = f
End Sub

'----------------------------------
'テンプレートテーブル機能
'----------------------------------

'テンプレートテーブル複製
Private Sub CopyTemplateTable( _
        Optional tname As String, _
        Optional bforce As Boolean)
    '
    'テンプレートテーブル取得
    Dim ra As Range, rb As Range
    Set ra = GetTemplateTable(tname)
    Set rb = GetTemplateTableRange(ra, eol:=True)
    If rb Is Nothing Then Exit Sub
    '
    '出力先取得
    Dim rc As Range
    Set rc = ActiveCell
    '
    'テーブルコピー
    ScreenUpdateOff
    rb.Copy Destination:=rc
    If Left(LeftBottom(rb).Offset(1, -2), 1) <> "#" Then
        ScreenUpdateOn
        Exit Sub
    End If
    '
    'テーブルサイズ取得
    Dim rm As Long, cm As Long
    rm = rb.Rows.Count
    cm = rb.Columns.Count
    '
    '操作抽出
    Dim cmd As Variant, arr As Variant
    Set rb = rb(1, 1).Offset(rm)
    Set rb = rb.Resize(SectionRowCount(rb.Offset(0, -2)))
    cmd = rb.Offset(0, -2).Resize(, 1).Value
    arr = rb.Value
    '
    '操作処理
    Dim rd As Range
    Dim r As Long, c As Long, i As Long
    For r = 1 To UBound(cmd, 1)
        Select Case LCase(Trim(cmd(r, 1)))
        Case "#continue"
            i = CLng(arr(r, 1)) + 1
            If i > 1 Then
                Set rb = rc(rm, 1).Resize(1, cm)
                Set rd = rc(rm, 1).Resize(i, cm)
                rb.AutoFill Destination:=rd, Type:=xlFillDefault
                rm = rm + i
            End If
        Case "#hide"
            For c = cm To 1 Step -1
                If arr(r, c) Then
                    rd(1, c).EntireColumn.Hidden = True
                End If
            Next c
        Case "#delete"
            For c = cm To 1 Step -1
                If arr(r, c) Then
                    Set rc = rd(1, c).Resize(rm, 1)
                    rc.Delete Shift:=xlToLeft
                    cm = cm - 1
                End If
            Next c
        End Select
    Next r
    '
    ScreenUpdateOn
End Sub

'テンプレートテーブル更新
Private Sub UpdateTemplateTable( _
        ra As Range, _
        Optional ByVal tname As String)
    If ra Is Nothing Then Exit Sub
    '
    '名前取得
    Dim rb As Range
    If tname = "" Then
        tname = InputBox("名前を入力してください。", app_name)
        tname = Trim(Replace(Replace(tname, "[", ""), "]", ""))
        If tname = "" Then
            Set rb = GetTemplateTable
            If rb Is Nothing Then Exit Sub
            tname = rb.Offset(0, -2)
            tname = Mid(tname, 2, Len(tname) - 2)
        End If
    End If
    '
    'テンプレートテーブル取得
    Set rb = GetTemplateTable(tname)
    '
    'テーブル登録
    Dim ws As Worksheet
    Set ws = ConfigSheet("#table", True)
    Dim rc As Range
    Set rc = ws.UsedRange
    If rc Is Nothing Then Set rc = ws.Cells(1, 1)
    Set rc = LeftBottom(rc).Offset(2)
    rc.Value = "[" & tname & "]"
    Set rc = rc.Offset(, 2)
    ra.Copy Destination:=rc
    Set rc = rc.Offset(ra.Rows.Count)
    '
    'データ行複製登録
    Dim s As String
    s = InputBox("繰り返し数があれば入力してください。", app_name)
    Dim i As Long
    i = CLng("0" & s)
    If i > 0 Then
        rc.Offset(0, -2) = "#continue"
        rc.Offset(0, 0) = i
        Set rc = rc.Offset(1)
    End If
    '
    'テーブル削除
    If rb Is Nothing Then Exit Sub
    GetTemplateTableRange(rb).EntireRow.Delete
End Sub

'テンプレートテーブル削除
Private Sub RemoveTemplateTable( _
    Optional ByVal tname As String, Optional bforce As Boolean)
    '
    'テンプレートテーブル取得
    Dim ra As Range, rb As Range
    Set ra = GetTemplateTable(tname)
    Set rb = GetTemplateTableRange(ra)
    If rb Is Nothing Then Exit Sub
    '
    '削除確認
    Dim res As VbMsgBoxResult
    Dim msg As String
    If Not bforce Then
        msg = "テンプレートを削除しますか。" & vbLf & tname
        res = MsgBox(msg, vbYesNo Or vbDefaultButton2, app_name)
        If res = vbNo Then Exit Sub
    End If
    '
    'テーブル削除
    rb.EntireRow.Delete
End Sub

'テンプレートテーブルの開始位置を取得
Private Function GetTemplateTable( _
    Optional tname As String) As Range
    'ヘッダ定義シート取得
    Dim ws As Worksheet
    Set ws = ConfigSheet("#table")
    If ws Is Nothing Then Exit Function
    '
    'テンプレート開始位置取得
    Dim ra As Range
    Set ra = SectionTags(ws.UsedRange.Columns(1))
    Set ra = SectionCell(ra, tname)
    If ra Is Nothing Then Exit Function
    '
    Set GetTemplateTable = ra.Offset(0, 2)
End Function

'テンプレートテーブル範囲を取得
Private Function GetTemplateTableRange( _
    ra As Range, Optional eol As Boolean) As Range
    If ra Is Nothing Then Exit Function
    '
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    '
    'テンプレート行列数取得
    Dim t As String
    Dim i As Long
    Dim r As Long, rm As Long
    Dim c As Long, cm As Long
    rm = SectionRowCount(ra.Offset(0, -2), eol)
    cm = ws.UsedRange.Columns.Count + ws.UsedRange.Column - ra.Column
    For r = 1 To rm
        i = ra(r, cm + 1).End(xlToLeft).Column - ra.Column + 1
        If c < i Then c = i
    Next r
    If r < rm Then rm = r
    cm = c
    If rm < 1 Or cm < 1 Then Exit Function
    Set GetTemplateTableRange = ra.Resize(rm, cm)
    Exit Function
End Function

'ヘッダーシート取得
Private Sub CopyHeaderSheet()
    ActivateConfigSheet "#table"
End Sub

'----------------------------------
'テンプレートでファイル読込
'----------------------------------

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

'テスト用
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

'----------------------------------
'その他
'----------------------------------

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

