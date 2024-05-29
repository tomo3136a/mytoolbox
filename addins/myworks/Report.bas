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
    Do Until Cells(r, c).Value = ""
        r = r + 1
    Loop
    If r > 1 Then
        If Cells(r - 1, c).Value <> "" Then
            s = "更新 " & s
        End If
    End If
    '
    With Cells(r, c)
        .HorizontalAlignment = xlRight
        .Value = s
    End With
End Sub

'----------------------------------------
'ページフォーマット設定
'----------------------------------------

Sub PagePreview()
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim ws As Worksheet
    For Each ws In ActiveWindow.SelectedSheets
        Call SheetPagePreview(ws)
    Next ws
    '
    Application.ScreenUpdating = fsu
End Sub

Private Sub SheetPagePreview(ws As Worksheet)
    Dim wnd As Window
    Set wnd = ws.Application.ActiveWindow
    wnd.Zoom = 100
    ws.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ws.PageSetup
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
End Sub

'----------------------------------
'テキスト変換
'----------------------------------

Sub Cells_Conv(ra As Range, Optional mode As Integer = 0)
    Select Case mode
    Case 1
        'トリム(冗長なスペース削除)
        Call Cells_RemoveSpace(ra)
    Case 2
        'シングルライン(冗長なスペース削除かつ1行化)
        Call Cells_RemoveSpace(ra, SingleLine:=True)
    Case 3
        'スペース削除
        Call Cells_RemoveSpace(ra, sep:="")
    Case 4
        '文字列変更(大文字に変換)
        Call Cells_StrConv(ra, vbUpperCase)
    Case 5
        '文字列変更(小文字に変換)
        Call Cells_StrConv(ra, vbLowerCase)
    Case 6
        '文字列変更(各単語の先頭の文字を大文字に変換)
        Call Cells_StrConv(ra, vbProperCase)
    Case 7
        '文字列変更(半角文字を全角文字に変換)
        Call Cells_StrConv(ra, vbWide)
    Case 8
        '文字列変更(全角文字を半角文字に変換)
        Call Cells_StrConv(ra, vbNarrow)
    Case 9
        '文字列変更(ASCII文字のみ半角化)
        Call Cells_StrConvNarrow(ra)
    Case Else
        '文字列変更(ASCII文字のみ半角化、冗長なスペース削除)
        Call Cells_StrConvNarrow(ra)
        Call Cells_RemoveSpace(ra)
    End Select
End Sub

'スペース削除
Private Sub Cells_RemoveSpace( _
        ra As Range, _
        Optional SingleLine As Boolean = False, _
        Optional sep As String = " ")
    Dim re1 As Object: Set re1 = regex("\s+")
    Dim re2 As Object: Set re2 = regex("[ 　\t]+")
    Dim re3 As Object: Set re3 = regex(" (\r?\n)")
    '
    Dim ce As Range
    For Each ce In ra.Cells
        If ce.Value <> "" Then
            If Not ce.HasFormula Then
                Dim s As String
                If SingleLine Then
                    s = re1.Replace(ce.Value, sep)
                Else
                    s = re2.Replace(ce.Value, sep)
                    s = re3.Replace(s, "$1")
                End If
                ce.Value = Trim(s)
            End If
        End If
    Next ce
End Sub

'文字列変更
' vbUpperCase   1   大文字に変換
' vbLowerCase   2   小文字に変換
' vbProperCase  3   各単語の先頭の文字を大文字に変換
' vbWide        4   半角文字を全角文字に変換
' vbNarrow      8   全角文字を半角文字に変換
' vbKatakana    16  ひらがなをカタカナに変換
' vbHiragana    32  カタカナをひらがなに変換
Private Sub Cells_StrConv(ra As Range, mode As Integer)
    Dim ce As Range
    For Each ce In ra.Cells
        If ce.Value <> "" Then
            If Not ce.HasFormula Then
                ce.Value = StrConv(ce.Value, mode)
            End If
        End If
    Next ce
End Sub

'文字列変更(ASCII文字のみ半角化)
Private Sub Cells_StrConvNarrow(ra As Range)
    Dim re As Object: Set re = regex("[！-～]+")
    '
    Dim ce As Range
    For Each ce In ra.Cells
        If ce.Value <> "" Then
            If Not ce.HasFormula Then
                Dim s As String
                s = ce.Value
                Dim m As Object
                For Each m In re.Execute(s)
                    s = Replace(s, m.Value, StrConv(m.Value, vbNarrow))
                Next m
                ce.Value = s
            End If
        End If
    Next ce
End Sub

'---------------------------------------------
'表示/非表示操作
'---------------------------------------------

Sub ShowHide(mode As Integer)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    '
    Select Case mode
    Case 1
        '非表示行削除・非表示列削除
        DeleteHideColumn
        DeleteHideRow
    Case 2
        '非表示列削除
        DeleteHideColumn
    Case 3
        '非表示行削除
        DeleteHideRow
    Case 4
        '非表示シート削除
        Call DeleteHideSheet
    Case 8
        '非表示シート表示
        Call VisibleHideSheet
    Case 9
        '非表示名前表示
        Call VisibleHideName
    Case Else
    End Select
    '
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'非表示行削除
Private Sub DeleteHideRow()
    Dim ra As Range
    Set ra = Selection
    Dim e As Long
    'e = Cells.SpecialCells(xlCellTypeLastCell).Row
    'e = FarBottom(ra).Row
    e = ra.Worksheet.UsedRange.Row
    e = e + ra.Worksheet.UsedRange.Rows.Count
    Dim i As Long
    For i = e + 1 To 1 Step -1
        If Rows(i).Hidden Then Rows(i).Delete
    Next i
    On Error Resume Next
    ra.Select
    On Error GoTo 0
End Sub

'非表示列削除
Private Sub DeleteHideColumn()
    Dim ra As Range
    Set ra = Selection
    Dim i As Long
    Dim e As Long
    'e = Cells.SpecialCells(xlCellTypeLastCell).column
    'e = FarBottom(ra).Row
    e = ra.Worksheet.UsedRange.Row
    e = e + ra.Worksheet.UsedRange.Rows.Count
    For i = Cells.SpecialCells(xlCellTypeLastCell).Column + 1 To 1 Step -1
        If Columns(i).Hidden Then Columns(i).Delete
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

Sub AddUserFormat(ra As Range, mode As Integer)
    Select Case mode
    Case 1
        '数式に条件付き書式を追加
        Call AddFormulaConditionColor(ra)
    Case 2
        '参照に色を付ける
        Call MarkRef(ra)
    Case 4
        '空白に条件付き書式を追加
        Call AddBlankConditionColor(ra)
    Case 9
        Call SelectRef(ra)
    Case 3
        '参照スタイル削除
        Call ClearMarkRef
    End Select
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

Private Sub AddFormulaConditionColor(ra As Range)
    Dim s As String
    s = ra.Cells(1, 1).Address(False, False)
    ra.FormatConditions.Add Type:=xlExpression, Formula1:="=ISFORMULA(" & s & ")"
    ra.FormatConditions(ra.FormatConditions.Count).SetFirstPriority
    
    Dim i As Long
    Call Application.Dialogs(xlDialogEditColor).Show(1)
    i = ActiveWorkbook.Colors(1)

    With ra.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .color = i
        .TintAndShade = 0
    End With
    ra.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub AddBlankConditionColor(ra As Range)
    Dim s As String
    s = ra.Cells(1, 1).Address(False, False)
    ra.FormatConditions.Add Type:=xlExpression, Formula1:="=TRIM(" & s & ")="""""
    ra.FormatConditions(ra.FormatConditions.Count).SetFirstPriority
    
    Dim i As Long
    'Call Application.Dialogs(xlDialogEditColor).Show(1)
    'i = ActiveWorkbook.Colors(1)
    i = RGB(127, 127, 127)

    With ra.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .color = i
        .TintAndShade = 0
    End With
    ra.FormatConditions(1).StopIfTrue = False
End Sub


Private Sub SelectRef(ra As Range)
    Dim wb As Workbook
    Set wb = ra.Parent.Parent
    
    Dim ra2 As Range
    On Error Resume Next
    Set ra2 = ra.DirectPrecedents
    On Error GoTo 0
    If ra2 Is Nothing Then Exit Sub
    Set ra2 = ra.Find("!", LookIn:=xlFormulas, LookAt:=xlPart, SearchFormat:=False)
    'Set ra2 = ra2.DirectDependents
    'Set ra2 = Intersect(ra, ra2)
    
    ra2.Select
End Sub

Private Sub MarkRef(ra As Range)
    Dim wb As Workbook
    Set wb = ra.Parent.Parent
    
    Dim ra2 As Range
    On Error Resume Next
    Set ra2 = ra.DirectPrecedents
    On Error GoTo 0
    If ra2 Is Nothing Then Exit Sub
    Set ra2 = ra2.DirectDependents
    Set ra2 = Intersect(ra, ra2)
    
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
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        End With
    End If
    On Error GoTo 0
    ra2.Style = s
End Sub

Private Sub ClearMarkRef()
    Dim s As String
    s = "参照"
    ActiveWorkbook.Styles(s).Delete
End Sub

Private Sub InsertFunction(ra As Range)
    Dim s As String
    s = "=LEFT(A1,MIN(FIND({1,2,3,4,5,6,7,8,9,0},ASC(A1)&1234567890))-1)"
End Sub

