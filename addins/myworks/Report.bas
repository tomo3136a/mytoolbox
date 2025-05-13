Attribute VB_Name = "Report"
'==================================
'レポート編集機能
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'レポートにサイン
'----------------------------------------

Sub ReportSign(ra As Range)
    Dim s As String
    Dim c As Long
    Dim r As Long
    s = Date & " " & Application.UserName
    c = ra.Column
    r = ra.Row
    '
    '書き込み先の空白探索
    Do Until Cells(r, c).Value = ""
        r = r + 1
    Loop
    
    '更新なら「更新」を明記
    If r > 1 Then
        If Cells(r - 1, c).Value <> "" Then
            s = "更新 " & s
        End If
    End If
    
    '右詰めにしてサイン書き込み
    With Cells(r, c)
        .HorizontalAlignment = xlRight
        .Value = s
    End With
End Sub

'----------------------------------------
'ページフォーマット設定
'----------------------------------------

Sub PagePreview()
    ScreenUpdateOff
    '
    Dim ws As Worksheet
    For Each ws In ActiveWindow.SelectedSheets
        '最終行が空白行でないなら右下に空白追加
        Dim ra As Range
        Set ra = ws.UsedRange
        Set ra = ra(ra.Rows.Count, ra.Columns.Count)
        If ra <> " " Then ra.Offset(1, 0) = " "
        
        '印刷範囲表示に設定        '
        Dim wnd As Window
        Set wnd = ws.Application.ActiveWindow
        wnd.Zoom = 100
        ws.PageSetup.PrintArea = ""
        Application.PrintCommunication = False
        With ws.PageSetup
            .Orientation = xlPortrait
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .CenterHorizontally = True
            .CenterVertically = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 0
        End With
        Application.PrintCommunication = True
        wnd.ActiveSheet.Cells(1, 1).Select
        wnd.ScrollColumn = 1
        wnd.ScrollRow = 1
        wnd.View = xlPageBreakPreview
        wnd.Zoom = 100
    Next ws
    '
    ScreenUpdateOn
End Sub

'----------------------------------
'テキスト変換
' mode=1: トリム(冗長なスペース削除)
'      2: シングルライン(冗長なスペース削除かつ1行化)
'      3: スペース削除
'      4: 文字列変更(大文字に変換)
'      5: 文字列変更(小文字に変換)
'      6: 文字列変更(各単語の先頭の文字を大文字に変換)
'      7: 文字列変更(半角文字を全角文字に変換)
'      8: 文字列変更(全角文字を半角文字に変換)
'      9: 文字列変更(ASCII文字のみ半角化)
'      *: 文字列変更(ASCII文字のみ半角化、冗長なスペース削除)
'----------------------------------

Sub TextConvProc(ra As Range, mode As Integer)
    ScreenUpdateOff
    '
    Dim rb As Range
    For Each rb In ra.Areas
        Set rb = Intersect(rb, ra.Parent.UsedRange)
        If rb Is Nothing Then Exit For
        '
        Dim va As Variant
        va = RangeToFormula2(rb)
        '
        Select Case mode
        Case 1: Call Cells_RemoveSpace(va)
        Case 2: Call Cells_RemoveSpace(va, SingleLine:=True)
        Case 3: Call Cells_RemoveSpace(va, sep:="")
        Case 4: Call Cells_StrConv(va, vbUpperCase)
        Case 5: Call Cells_StrConv(va, vbLowerCase)
        Case 6: Call Cells_StrConv(va, vbProperCase)
        Case 7: Call Cells_StrConv(va, vbWide)
        Case 8: Call Cells_StrConv(va, vbNarrow)
        Case 9: Call Cells_StrConvNarrow(va)
        Case Else
            Call Cells_StrConvNarrow(va)
            Call Cells_RemoveSpace(va)
        End Select
        '
        rb.Value = va
    Next rb
    '
    ScreenUpdateOn
End Sub

'スペース削除
Private Sub Cells_RemoveSpace( _
        va As Variant, _
        Optional SingleLine As Boolean = False, _
        Optional sep As String = " ")
    Dim re1 As Object, re2 As Object, re3 As Object
    Set re1 = regex("[\s\u00A0\u3000]+")
    Set re2 = regex("[ \t\v\f\u00A0\u3000]+")
    Set re3 = regex(sep & "(\r?\n)")
    '
    Dim r As Long, c As Long
    For r = LBound(va, 1) To UBound(va, 1)
        For c = LBound(va, 2) To UBound(va, 2)
            Dim s As String
            s = va(r, c)
            If s <> "" Then
                If SingleLine Then
                    s = re1.Replace(s, sep)
                Else
                    s = re2.Replace(s, sep)
                    s = re3.Replace(s, "$1")
                End If
                va(r, c) = Trim(s)
            End If
        Next c
    Next r
End Sub

'文字列変更
' vbUpperCase   1   大文字に変換
' vbLowerCase   2   小文字に変換
' vbProperCase  3   各単語の先頭の文字を大文字に変換
' vbWide        4   半角文字を全角文字に変換
' vbNarrow      8   全角文字を半角文字に変換
' vbKatakana    16  ひらがなをカタカナに変換
' vbHiragana    32  カタカナをひらがなに変換
Private Sub Cells_StrConv(va As Variant, mode As Integer)
    Dim s As String
    Dim r As Long, c As Long
    For r = LBound(va, 1) To UBound(va, 1)
        For c = LBound(va, 2) To UBound(va, 2)
            s = va(r, c)
            If s <> "" And Left(s, 1) <> "=" Then
                va(r, c) = StrConv(s, mode)
            End If
        Next c
    Next r
End Sub

'文字列変更(ASCII文字のみ半角化)
Private Sub Cells_StrConvNarrow(va As Variant)
    Dim re As Object: Set re = regex("[！-～]+")
    '
    Dim s As String
    Dim r As Long, c As Long
    For r = LBound(va, 1) To UBound(va, 1)
        For c = LBound(va, 2) To UBound(va, 2)
            s = va(r, c)
            If s <> "" And Left(s, 1) <> "=" Then
                Dim m As Object
                For Each m In re.Execute(s)
                    s = Replace(s, m.Value, StrConv(m.Value, vbNarrow))
                Next m
                va(r, c) = s
            End If
        Next c
    Next r
End Sub

Private Function RangeToFormula2(ra As Range) As Variant
    Dim va As Variant
    va = ra.Formula2
    If ra.Count = 1 Then
        ReDim va(1 To 1, 1 To 1)
        va(1, 1) = ra.Formula2
    End If
    RangeToFormula2 = va
End Function

'---------------------------------------------
'表示/非表示操作
'mode=1: 非表示行削除・非表示列削除
'     2: 非表示列削除
'     3: 非表示行削除
'     4: 非表示シート削除
'     8: 非表示シート表示
'     9: 非表示名前表示
'---------------------------------------------

Sub ShowHide(mode As Integer)
    ScreenUpdateOff
    '
    Select Case mode
    Case 1: DeleteHideColumn: DeleteHideRow
    Case 2: DeleteHideColumn
    Case 3: DeleteHideRow
    Case 4: DeleteHideSheet
    Case 8: VisibleHideSheet
    Case 9: VisibleHideName
    End Select
    '
    ScreenUpdateOn
End Sub

'非表示行削除
Private Sub DeleteHideRow()
    Dim ra As Range
    Set ra = Selection
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Dim e As Long
    e = ws.UsedRange.Row
    e = e + ws.UsedRange.Rows.Count
    Dim i As Long
    For i = e + 1 To 1 Step -1
        If ws.Rows(i).Hidden Then ws.Rows(i).Delete
    Next i
    On Error Resume Next
    ra.Select
    On Error GoTo 0
End Sub

'非表示列削除
Private Sub DeleteHideColumn()
    Dim ra As Range
    Set ra = Selection
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Dim e As Long
    e = ws.UsedRange.Column
    e = e + ws.UsedRange.Columns.Count
    Dim i As Long
    For i = e + 1 To 1 Step -1
        If ws.Columns(i).Hidden Then ws.Columns(i).Delete
    Next i
    On Error Resume Next
    ra.Select
    On Error GoTo 0
End Sub

'非表示シート削除
Private Sub DeleteHideSheet()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    '
    Dim cnt As Integer
    Dim msg As String
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = 0 Then
            Dim s As String
            s = ws.name
            ws.Delete
            cnt = cnt + 1
            msg = msg & vbCrLf & s
        End If
    Next ws
    '
    If cnt > 0 Then
        MsgBox cnt & "シートを削除しました。" & msg
    End If
End Sub

'非表示シート表示
Private Sub VisibleHideSheet()
    Dim ws As Worksheet
    Dim cnt As Integer
    Dim msg As String
    '
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = 0 Then
            ws.Visible = True
            cnt = cnt + 1
            msg = msg & vbCrLf & ws.name
        End If
    Next ws
    '
    If cnt > 0 Then
        MsgBox cnt & "シートを表示にしました。" & msg
    End If
End Sub

'非表示名前表示
Private Sub VisibleHideName()
    Dim nm As name
    Dim cnt As Integer
    Dim msg As String
    '
    For Each nm In ActiveWorkbook.Names
        If nm.Visible = 0 Then
            If Right(nm.name, 8) <> "Database" Then
                nm.Visible = True
                cnt = cnt + 1
                msg = msg & vbCrLf & nm.name
            End If
        End If
    Next nm
    '
    If cnt > 0 Then
        MsgBox "名前を" & cnt & "件表示にしました。" & msg
    End If
End Sub

'---------------------------------------------
'書式操作
' mode=1: 数式に条件付き書式を追加
'      2: 0に条件付き書式を追加
'      3: 空白に条件付き書式を追加
'      4: 参照に色を付ける
'      8: 参照スタイル削除
'---------------------------------------------

Sub UserFormatProc(ra As Range, mode As Long)
    Dim rb As Range
    Set rb = ra.Worksheet.UsedRange
    Set rb = Intersect(ra, rb)
    If rb Is Nothing Then Set rb = ra
    '
    Select Case mode
    Case 1: Call AddFormulaConditionColor(rb)
    Case 2: Call AddZeroConditionColor(rb)
    Case 3: Call AddBlankConditionColor(rb)
    Case 4: Call MarkRef(rb)
    Case 5: Call ListConditionFormat
    Case 8: Call ClearMarkRef
    End Select
End Sub

Private Sub AddFormulaConditionColor(ra As Range)
    '条件式設定
    Dim s As String
    s = ra.Cells(1, 1).Address(False, False)
    s = "=ISFORMULA(" & s & ")"
    With ra.FormatConditions
        .Add Type:=xlExpression, Formula1:=s
        .Item(.Count).SetFirstPriority
    End With
    
    '条件時背景色選択
    Dim i As Long
    Call Application.Dialogs(xlDialogEditColor).Show(1)
    i = ActiveWorkbook.Colors(1)

    '条件時背景色指定
    With ra.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = i
        .TintAndShade = 0
    End With
    ra.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub AddZeroConditionColor(ra As Range)
    '条件式設定
    Dim s As String
    s = ra.Cells(1, 1).Address(False, False)
    s = "=AND(" & s & "<>""""," & s & "=0)"
    With ra.FormatConditions
        .Add Type:=xlExpression, Formula1:=s
        .Item(.Count).SetFirstPriority
    End With
    
    '条件時背景色選択
    Dim i As Long
    Call Application.Dialogs(xlDialogEditColor).Show(1)
    i = ActiveWorkbook.Colors(1)

    '条件時背景色指定
    With ra.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = i
        .TintAndShade = 0
    End With
    ra.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub AddBlankConditionColor(ra As Range)
    '条件式設定
    Dim s As String
    s = ra.Cells(1, 1).Address(False, False)
    s = "=TRIM(" & s & ")="""""
    With ra.FormatConditions
        .Add Type:=xlExpression, Formula1:=s
        .Item(.Count).SetFirstPriority
    End With
    
    '条件時背景色選択
    Dim i As Long
    Call Application.Dialogs(xlDialogEditColor).Show(1)
    i = ActiveWorkbook.Colors(1)

    '条件時背景色指定
    With ra.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = i
        .TintAndShade = 0
    End With
    ra.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub MarkRef(ra As Range)
    Dim wb As Workbook
    Set wb = ra.Worksheet.Parent
    
    Dim rb As Range
    On Error Resume Next
    Set rb = ra.DirectPrecedents
    On Error GoTo 0
    If rb Is Nothing Then Exit Sub
    Set rb = rb.DirectDependents
    Set rb = Intersect(ra, rb)
    
    Dim s As String
    s = "参照"
    On Error Resume Next
    If wb.Styles(s) Is Nothing Then
        With wb.Styles.Add(s)
            .IncludeNumber = False
            .IncludeFont = False
            .IncludeAlignment = False
            .IncludeBorder = False
            .IncludePatterns = True
            .IncludeProtection = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = 0
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.8
                .PatternTintAndShade = 0
            End With
        End With
    End If
    On Error GoTo 0
    rb.Style = s
End Sub

Private Sub ClearMarkRef()
    Dim s As String
    s = "参照"
    On Error Resume Next
    ActiveWorkbook.Styles(s).Delete
    On Error GoTo 0
End Sub

Private Sub ListConditionFormat()
    Dim ra As Range
    Set ra = ActiveCell
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim fc As FormatCondition
    For Each fc In ws.Cells.FormatConditions
        ra.Value = "'" & fc.Formula1
        ra.Offset(, 1).Value = fc.PTCondition
        ra.Offset(, 2).Value = fc.AppliesTo.Address
        ra.Offset(, 3).Value = fc.Type
        Set ra = ra.Offset(1)
    Next fc
End Sub


Private Sub NewStyle(ra As Range, name As String)
    If ra Is Nothing Then Exit Sub
    If name = "" Then Exit Sub
    
    Dim wb As Workbook
    Set wb = ra.Parent.Parent
    On Error Resume Next
    If ra.Parent.Parent.Styles(name) Is Nothing Then
        With ra.Parent.Parent.Styles.Add(name)
            .IncludeNumber = False
            .IncludeFont = False
            .IncludeAlignment = False
            .IncludeBorder = False
            .IncludePatterns = True
            .IncludeProtection = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = 0
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.8
                .PatternTintAndShade = 0
            End With
        End With
    End If
    On Error GoTo 0
    ra.Style = name
End Sub

'---------------------------------------------
'定型式挿入
'mode=1: 文字列分割(英字・数値)
'     2: 文字列分割(数値・英字・数値)
'     3: 差分マーカー
'---------------------------------------------

Sub UserFormulaProc(ra As Range, mode As Integer)
    Dim v1 As Integer, v2 As Integer, v3 As Integer
    Dim r0 As Range, r1 As Range, r2 As Range, r3 As Range
    Select Case mode
    Case 1
        '文字列分割(英字・数値)
        Set r0 = ra.Cells(1, 1)
        On Error Resume Next
        Set r1 = Application.InputBox("記号位置", "文字列分割", r0.Offset(0, 1).Address, Type:=8)
        On Error GoTo 0
        If r1 Is Nothing Then Exit Sub
        v1 = ra.Column - r1.Column
        ra.Offset(, -v1).Formula2R1C1 = "=LET(v,RC[" & v1 & "],LEFT(v,MIN(FIND({1,2,3,4,5,6,7,8,9,0},v&""1234567890""))-1))"
        
        On Error Resume Next
        Set r2 = Application.InputBox("数値位置", "文字列分割", r1.Offset(0, 1).Address, Type:=8)
        On Error GoTo 0
        If r2 Is Nothing Then Exit Sub
        If r2.Address = r1.Address Then Exit Sub
        v2 = ra.Column - r2.Column
        ra.Offset(, -v2).FormulaR1C1 = "=MID(RC[" & v2 & "],LEN(RC[" & (v2 - v1) & "])+1,LEN(RC[" & v2 & "]))"
    Case 2
        '文字列分割(数値・英字・数値)
        If True Then
            Set r0 = ra.Cells(1, 1)
            On Error Resume Next
            Set r1 = Application.InputBox("先頭数値位置", "文字列分割", r0.Offset(0, 1).Address, Type:=8)
            On Error GoTo 0
            If r1 Is Nothing Then Exit Sub
            v1 = ra.Column - r1.Column
            ra.Offset(, -v1).FormulaR1C1 = "=IFERROR(VALUE(LEFT(RC[" & v1 & "],2)),IFERROR(VALUE(LEFT(RC[" & v1 & "],1)),""""))"
            
            On Error Resume Next
            Set r2 = Application.InputBox("記号位置", "文字列分割", r1.Offset(0, 1).Address, Type:=8)
            On Error GoTo 0
            If r2 Is Nothing Then Exit Sub
            If r2.Address = r1.Address Then Exit Sub
            v2 = ra.Column - r2.Column
            ra.Offset(, -v2).Formula2R1C1 = "=LET(v,MID(RC[" & v2 & "],LEN(RC[" & (v2 - v1) & "])+1,LEN(RC[" & v2 & "])),LEFT(v,MIN(FIND({1,2,3,4,5,6,7,8,9,0},v&""1234567890""))-1))"
            
            On Error Resume Next
            Set r3 = Application.InputBox("数値位置", "文字列分割", r2.Offset(0, 1).Address, Type:=8)
            On Error GoTo 0
            If r3 Is Nothing Then Exit Sub
            If r3.Address = r1.Address Then Exit Sub
            If r3.Address = r2.Address Then Exit Sub
            v3 = ra.Column - r3.Column
            ra.Offset(, -v3).FormulaR1C1 = "=MID(RC[" & v3 & "],LEN(RC[" & (v3 - v1) & "]&RC[" & (v3 - v2) & "])+1,LEN(RC[" & v3 & "]))"
        Else
            ra.Offset(, 1).FormulaR1C1 = "=IFERROR(VALUE(LEFT(RC[" & -1 & "],2)),IFERROR(VALUE(LEFT(RC[" & -1 & "],1)),""""))"
            ra.Offset(, 2).Formula2R1C1 = _
                "=LET(v,MID(RC[-2],LEN(RC[-1])+1,LEN(RC[-2])),LEFT(v,MIN(FIND({1,2,3,4,5,6,7,8,9,0},v&""1234567890""))-1))"
            ra.Offset(, 3).FormulaR1C1 = "=MID(RC[-3],LEN(RC[-2]&RC[-1])+1,LEN(RC[-3]))"
        End If
    Case 3
        '差分マーカー
        Set r0 = ra.Cells(1, 1)
        On Error Resume Next
        Set r1 = Application.InputBox("比較元位置1", "差分マーカ―", r0.Offset(0, 1).Address, Type:=8)
        On Error GoTo 0
        If r1 Is Nothing Then Exit Sub
        On Error Resume Next
        Set r2 = Application.InputBox("比較元位置2", "差分マーカ―", r1.Offset(0, 1).Address, Type:=8)
        On Error GoTo 0
        If r2 Is Nothing Then Exit Sub
        If r2.Address = r1.Address Then Exit Sub
        v1 = r1.Column - ra.Column
        v2 = r2.Column - ra.Column
        '
        ra.Formula2R1C1 = _
            "=IF(OFFSET(RC,0," & v1 & ")=OFFSET(RC,0," & v2 & "),""〇"","""")"
    End Select
End Sub

