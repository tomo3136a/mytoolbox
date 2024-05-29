Attribute VB_Name = "DataList"
'==================================
'データリスト取得
'==================================

Option Explicit
Option Private Module

Private g_sheet As Boolean

'----------------------------------
'共通機能
'----------------------------------

Sub SetDataListParam(id As Integer, ByVal val As String)
    Select Case id
    Case 1
        g_sheet = val
    End Select
End Sub

Function GetDataListParam(id As Integer) As String
    Dim val As String
    Select Case id
    Case 1
        val = g_sheet
    End Select
    GetDataListParam = val
End Function

'---------------------------------------------
'共通
'---------------------------------------------

'プロパティ
Private Sub SetInfoSheet(Optional ws As Worksheet, Optional v As String = "")
    If ws Is Nothing Then Set ws = ActiveSheet
    ws.CustomProperties.Add "info", v
End Sub

Private Function TestInfoSheet(ws As Worksheet) As Boolean
    If SheetPropertyIndex(ws, "info") > 0 Then TestInfoSheet = True
End Function

'---------------------------------------------
'情報リストアップ
'---------------------------------------------

Sub AddInfoSheet(mode As Integer)
    Application.ScreenUpdating = False
    '
    Dim Title As String
    Title = GetInfoTitle(mode)
    '
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim ra As Range
    If g_sheet Then
        Dim ws As Worksheet
        Set ws = wb.Worksheets.Add
        Call SetInfoSheet(ws, Title)
        If Title <> "" Then ws.name = UniqueSheetName(wb, Title)
        Set ra = ws.Cells(2, 2)
    Else
        Set ra = ActiveCell
    End If
    '
    Call AddInfoTable(mode, ra, Title, wb)
    '
    Application.ScreenUpdating = True
End Sub

Private Function GetInfoTitle(mode As Integer) As String
    Dim s As String
    Select Case mode
    Case 1
        s = "目次"
    Case 2
        s = "シート一覧"
    Case 3
        s = "名前一覧"
    Case 4
        s = "リンク一覧"
    Case 5
        s = "プロパティ一覧"
    Case 6
        s = "コメント一覧"
    Case 7
        s = "ファイル一覧"
    End Select
    GetInfoTitle = s
End Function

Private Sub AddInfoTable(mode As Integer, ByRef ra As Range, _
        Optional Title As String, Optional wb As Workbook)
    Dim hd As Range
    Set hd = ra.Offset(0)
    If Title <> "" Then Set hd = hd.Offset(1)
    Dim ce As Range
    Set ce = hd.Offset(1)
    '
    Dim r As Long
    r = ce.Row
    '
    Dim h As Variant
    Select Case mode
    Case 1
        h = Array("番号", "名前", "リンク", "説明")
        Call IndexList(ce, wb)
    Case 2
        h = Array("番号", "シート名", "状態", "使用範囲", _
            "テーブル数", "グラフ数", "図形数", "名前数", _
            "リンク数", "コメント数", "プロパティ数", _
            "高さ(標準)", "幅(標準)")
        Call SheetList(ce, wb)
    Case 3
        h = Array("番号", "名前", "状態", "参照範囲", _
            "値", "種類", "範囲", "コメント")
        Call NameList(ce, wb)
    Case 4
        h = Array("番号", "シート", "状態", "リンク元", _
            "表示文字列", "リンク先", "ヒント")
        Call LinkList(ce, wb)
    Case 5
        h = Array("番号", "名前", "状態", "参照範囲", _
            "値", "種類", "範囲", "コメント")
        Call PropList(ce, wb)
    Case 6
        h = Array("番号", "シート", "状態", _
            "参照範囲", "値", "表示")
        Call CommentList(ce, wb)
    Case 7
        Dim path As String
        path = ActiveWorkbook.path
        path = SelectFolder(path)
        If path = "" Then Exit Sub
        h = Array("番号", "名前", "状態", "種別", "サイズ", "編集日")
        Set hd = hd.Offset(1)
        Set ce = ce.Offset(1)
        Call FileList(ce, path)
        If r < ce.Row Then hd.Offset(-1).Value = GetShortPath(path)
    End Select
    '
    If r = ce.Row Then Exit Sub
    If Title <> "" Then ra.Value = Title
    Range(hd, hd.Offset(0, UBound(h))) = h
    Call Waku(hd, fit:=True)
End Sub

'---------------------------------------------
'テーブル作成
'---------------------------------------------

'目次一覧
Private Sub IndexList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If TestInfoSheet(ws) Then
        ElseIf ws.Visible Then
            no = no + 1
            ra.Value = no
            ra.Offset(0, 1).Value = ws.name
            ra.Worksheet.Hyperlinks.Add _
                Anchor:=ra.Offset(0, 2), _
                Address:="", _
                SubAddress:=(ws.name & "!A1"), _
                TextToDisplay:="シート", _
                ScreenTip:=ws.name
            Set ra = ra.Offset(1)
        End If
    Next ws
End Sub

'シート一覧
Private Sub SheetList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    On Error Resume Next
    '
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.name = ActiveSheet.name Then
        Else
            no = no + 1
            ra.Value = no
            ra.Worksheet.Hyperlinks.Add _
                Anchor:=ra.Offset(0, 1), _
                Address:=wb.path & "\" & wb.name, _
                SubAddress:=ws.name & "!A1", _
                TextToDisplay:=ws.name
                ', _
                'ScreenTip:=ws.name
            If Not ws.Visible Then ra.Offset(0, 2).Value = "非表示"
            ra.Offset(0, 3).Value = ws.UsedRange.Address(False, False)
            ra.Offset(0, 4).Value = ws.Shapes.Count
            ra.Offset(0, 5).Value = ws.ChartObjects.Count
            ra.Offset(0, 6).Value = ws.Shapes.Count
            ra.Offset(0, 7).Value = ws.Names.Count
            ra.Offset(0, 8).Value = ws.Hyperlinks.Count
            ra.Offset(0, 9).Value = ws.Comments.Count
            ra.Offset(0, 10).Value = ws.CustomProperties.Count
            ra.Offset(0, 11).Value = ws.StandardHeight
            ra.Offset(0, 12).Value = ws.StandardWidth
            Set ra = ra.Offset(1)
        End If
    Next ws
    '
    On Error GoTo 0
End Sub

'名前一覧
Private Sub NameList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    On Error Resume Next
    '
    Dim nm As name
    For Each nm In wb.Names
        Dim sts As String
        sts = ""
        no = no + 1
        ra.Value = no
        ra.Offset(0, 1).Value = nm.name
        If nm.Visible = False Then sts = "非表示"
        ra.Offset(0, 3).Value = "'" & nm.Value
        ra.Offset(0, 4).Value = StrRange(nm.Value)
        If err Then
            ra.Offset(0, 4).Value = "#REF!"
            sts = "エラー"
            err.Clear
        End If
        ra.Offset(0, 5).Value = TypeName(nm.Parent)
        ra.Offset(0, 6).Value = nm.Parent.name
        ra.Offset(0, 7).Value = nm.Comment
        If sts <> "" Then ra.Offset(0, 2).Value = sts
        Set ra = ra.Offset(1)
    Next nm
    '
    On Error GoTo 0
End Sub

'リンク一覧
Private Sub LinkList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If TestInfoSheet(ws) Then
        Else
            Dim lnk As Hyperlink
            For Each lnk In ws.Hyperlinks
                no = no + 1
                ra.Value = no
                Call HyperlinkInfo(ra.Offset(0, 1), ws, lnk)
                Set ra = ra.Offset(1)
            Next lnk
        End If
    Next ws
End Sub

Private Sub HyperlinkInfo(ra As Range, ws As Worksheet, lnk As Hyperlink)
    Dim sts As String
    Dim src_addr As String
    Dim src_disp As String
    If lnk.Type = 0 Then
        src_addr = lnk.Range.Address
        src_disp = lnk.TextToDisplay
    Else
        src_addr = lnk.Shape.TopLeftCell.Address
        src_disp = "[" & lnk.Shape.name & "]"
    End If
    Dim src As String
    src = "'" & ws.name & "'!" & src_addr
    '
    Dim dst_addr As String
    Dim dst_disp As String
    '
    On Error Resume Next
    ra.Value = ws.name
    ActiveSheet.Hyperlinks.Add Anchor:=ra.Offset(0, 2), _
        Address:="", SubAddress:=src, TextToDisplay:=src_addr
    If err Then sts = "Error": err.Clear
    ra.Offset(0, 3).Value = src_disp
    '
    If lnk.SubAddress = "" Then
        ActiveSheet.Hyperlinks.Add Anchor:=ra.Offset(0, 4), _
            Address:=lnk.Address, TextToDisplay:=lnk.Address
    ElseIf lnk.Address = "" Then
        ActiveSheet.Hyperlinks.Add Anchor:=ra.Offset(0, 4), _
            Address:="", SubAddress:=lnk.SubAddress, TextToDisplay:=lnk.SubAddress
    Else
        ActiveSheet.Hyperlinks.Add Anchor:=ra.Offset(0, 4), _
            Address:=lnk.Address, SubAddress:=lnk.SubAddress, _
            TextToDisplay:="'" & lnk.Address & "'!" & lnk.SubAddress
    End If
    If err Then sts = "Error": err.Clear
    ra.Offset(0, 5).Value = lnk.ScreenTip
    '
    If lnk.Address <> "" Then
        Dim wb As Workbook: Set wb = ws.Parent
        Dim path As String: path = wb.path & "\" & lnk.Address
        If Dir(path) = "" Then sts = "リンク切れ(" & lnk.Address & ")"
        If err Then sts = "不明": err.Clear
    End If
    '
    If sts Then ra.Offset(0, 1).Value = sts
    On Error GoTo 0
End Sub

'プロパティ一覧
Private Sub PropList(ByRef ra As Range, wb As Workbook)
    Dim ts() As Variant
    ts = Array("", "整数", "論理値", "日付", "文字列", "実数")
    '
    Dim no As Integer
    Dim sts As String
    Dim dp As DocumentProperty
    On Error Resume Next
    '
    'BuiltinDocumentProperties
    For Each dp In wb.BuiltinDocumentProperties
        If dp Is Nothing Then
        Else
            sts = ""
            no = no + 1
            ra.Value = no
            ra.Offset(0, 1).Value = dp.name
            ra.Offset(0, 3).Value = ts(dp.Type)
            ra.Offset(0, 4).Value = dp.Value
            ra.Offset(0, 5).Value = "組み込み"
            If err Then sts = "エラー": err.Clear
            ra.Offset(0, 2).Value = sts
            Set ra = ra.Offset(1)
        End If
    Next dp
    '
    'customDocumentPropertoes
    For Each dp In wb.customDocumentPropertoes
        If dp Is Nothing Then
        Else
            sts = ""
            no = no + 1
            ra.Value = no
            ra.Offset(0, 1).Value = dp.name
            ra.Offset(0, 3).Value = ts(dp.Type)
            ra.Offset(0, 4).Value = dp.Value
            ra.Offset(0, 5).Value = "カスタム"
            If err Then sts = "エラー": err.Clear
            ra.Offset(0, 2).Value = sts
            Set ra = ra.Offset(1)
        End If
    Next dp
    '
    'CustomProperty
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If TestInfoSheet(ws) Then
        Else
            Dim cp As CustomProperty
            For Each cp In ws.CustomProperties
                sts = ""
                no = no + 1
                ra.Value = no
                ra.Offset(0, 1).Value = cp.name
                ra.Offset(0, 3).Value = TypeName(cp.Value)
                ra.Offset(0, 4).Value = cp.Value
                ra.Offset(0, 5).Value = ws.name
                If err Then sts = "エラー": err.Clear
                ra.Offset(0, 2).Value = sts
                Set ra = ra.Offset(1)
            Next cp
        End If
    Next ws
    '
    On Error GoTo 0
End Sub

'コメント一覧
Private Sub CommentList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    Dim cm As Comment
    On Error Resume Next
    '
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        For Each cm In ws.Comments
            no = no + 1
            ra.Value = no
            ra.Offset(0, 1).Value = ws.name
            If TypeName(cm.Parent) = "Range" Then
                ra.Offset(0, 3).Value = cm.Parent.Address
            End If
            ra.Offset(0, 4).Value = cm.text
            If cm.Visible Then
                ra.Offset(0, 5).Value = "表示"
            End If
            Set ra = ra.Offset(1)
        Next cm
    Next ws
    '
    On Error GoTo 0
End Sub

'ファイル一覧
Private Sub FileList(ByRef ra As Range, path As String)
    If path = "" Then Exit Sub
    Dim no As Integer
    On Error Resume Next
    '
    Dim obj As Variant
    For Each obj In fso.GetFolder(path).SubFolders
        no = no + 1
        ra.Value = no
        With obj
            ra.Offset(0, 1).Value = .name
            ra.Offset(0, 2).Value = ""
            ra.Offset(0, 3).Value = "フォルダ"
            ra.Offset(0, 4).Value = .Size
            ra.Offset(0, 4).Style = "Comma [0]"
            ra.Offset(0, 5).Value = .DateLastModified
            ra.Offset(0, 5).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss;@"
        End With
        Set ra = ra.Offset(1)
    Next obj
    '
    For Each obj In fso.GetFolder(path).Files
        no = no + 1
        ra.Value = no
        With obj
            ra.Offset(0, 1).Value = .name
            ra.Offset(0, 2).Value = ""
            ra.Offset(0, 3).Value = "ファイル"
            ra.Offset(0, 4).Value = .Size
            ra.Offset(0, 4).Style = "Comma [0]"
            ra.Offset(0, 5).Value = .DateLastModified
            ra.Offset(0, 5).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss;@"
        End With
        Set ra = ra.Offset(1)
    Next obj
    '
    On Error GoTo 0
End Sub

