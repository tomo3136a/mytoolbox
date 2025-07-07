Attribute VB_Name = "Keisen"
'==================================
'セル操作
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'機能呼び出し
'----------------------------------------

'選択
' mode=1,2: 先頭へ移動
'      3: 末尾へ移動
'      4: 行選択
'      5: 列選択
'      6: ヘッダ行選択
'      7: テーブル選択
'
Sub SelectProc(ra As Range, Optional mode As Integer)
    Dim rb As Range
    Set rb = CurrentTableRegion(ra)
    Select Case mode
    Case 1, 2: rb(1, 1).Select
    Case 3: rb(rb.Rows.Count + 1, 1).Select
    Case 4: Intersect(rb, ra.EntireRow).Select
    Case 5: Intersect(rb, ra.EntireColumn).Select
    Case 6: rb.Rows(1).Select
    Case 7: rb.Select
    End Select
End Sub

'罫線枠
' mode=1: 罫線枠(枠,幅合わせ)
'      2: 罫線枠(標準)
'      3: 罫線枠(階層構造)
'      4: ヘッダフィルタ
'      5: ヘッダ幅合わせ
'      6: ヘッダ固定
'      7: ヘッダ色
'      8: 枠クリア
'      9: 値クリア
'      10: テーブルクリア
'
Sub WakuProc(ra As Range, Optional mode As Integer)
    Select Case mode
    Case 1: Waku ra, fit:=True
    Case 2: Waku ra
    Case 3: WakuLayered ra
    Case 4: HeaderFilter ra
    Case 5: HeaderAutoFit ra
    Case 6: HeaderFixed ra
    Case 7: HeaderColor ra
    Case 8: WakuClear ra: ra.FormatConditions.Delete
    Case 9: TableRange(TableHeaderRange(TableLeftTop(ra)).Offset(1)).Clear
    Case 10: TableRange(TableHeaderRange(TableLeftTop(ra))).Clear
    End Select
End Sub

'列追加
' mode=1: 番号列追加
Sub AddColumn(ra As Range, mode As Integer)
    'テーブルの1列取得
    Dim rb As Range
    Set rb = CurrentTableRegion(ra)
    Set rb = Intersect(rb, ra.EntireColumn)
    If ra.Rows.Count > 1 Then Set rb = ra
    Set rb = rb.Columns(1)
    rb.EntireColumn.Insert Shift:=xlShiftToRight
    
    '1列追加し書式を複製
    Dim rc As Range
    Set rc = rb.Offset(0, -1)
    rb.Copy
    rc.PasteSpecial Paste:=xlPasteFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = 0
    
    '複数行選択している場合はヘッダ行なしとする
    Dim bhdr As Boolean
    If ra.CurrentRegion.Row = rb.Row Then bhdr = True
    
    '列に値を入れる
    Select Case mode
    Case 1: AddNoColumn rc, bhdr
    'Case 2: AddRankColumn rc, bhdr
    End Select
    
    '選択範囲を戻す
    ra.Select
End Sub


'----------------------------------------
'API
'----------------------------------------

'囲い
Sub Waku(ByVal ra As Range, _
        Optional filter As Boolean, _
        Optional fit As Boolean, _
        Optional fixed As Boolean, _
        Optional icolor As Integer = 15 _
    )
    If ra.Count = 1 Then Set ra = CurrentTableRegion(ra)
    ra.Borders.LineStyle = xlContinuous
    '
    Dim rh As Range
    Set rh = ra.Rows(1)
    If GetHeaderColor = 0 Then
        rh.Interior.ColorIndex = icolor
    Else
        rh.Interior.Color = GetHeaderColor
    End If
    If filter Then HeaderFilter rh
    '
    Set ra = TableRange(TableHeaderRange(ra))
    If fit Then ra.Columns.AutoFit

End Sub

'囲い(階層構造)
Private Sub WakuLayered(ByVal ra As Range)
    If ra.Rows.Count = 1 And ra.Count > 1 Then
        Set ra = Intersect(ra.CurrentRegion, ra.EntireColumn)
    End If
    If ra.Count = 1 Then Set ra = CurrentTableRegion(ra)
    
    Waku ra, fit:=True
    ra.FormatConditions.Delete
    
    Dim s As String
    s = ra(1, 1).Address(False, False)
    s = "=LET(a," & s & ",b,OFFSET(a,-1,0),OR(""""&a="""",AND(""""&a=""""&b,SUBTOTAL(3,b)>0)))"
    ra.FormatConditions.Add Type:=xlExpression, Formula1:=s
    ra.FormatConditions(ra.FormatConditions.Count).SetFirstPriority
    ra.FormatConditions(1).NumberFormat = ";;;"
    ra.FormatConditions(1).Borders(xlTop).LineStyle = xlNone
    ra.FormatConditions(1).StopIfTrue = False
End Sub

'囲いクリア
Private Sub WakuClear(ByVal ra As Range)
    If ra.Rows.Count = 1 And ra.Count > 1 Then
        Set ra = Intersect(ra.CurrentRegion, ra.EntireColumn)
    End If
    If ra.Count = 1 Then Set ra = ra.CurrentRegion
    
    ra.FormatConditions.Delete
    ra.Interior.ColorIndex = xlColorIndexNone
    ra.Borders.LineStyle = xlNone
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
    ActiveWindow.FreezePanes = False
End Sub

'----------------------------------------
'番号列追加
'----------------------------------------

'番号列追加
Sub AddNoColumn(ra As Range, bhdr As Boolean, Optional shdr As String = "No.")
    Dim arr() As Variant
    arr = ra.Value
    Dim i As Long, j As Long
    If Not bhdr Or shdr = "" Then j = 1
    For i = 1 To ra.Rows.Count
        arr(i, 1) = j
        j = j + 1
    Next i
    If bhdr And shdr <> "" Then arr(1, 1) = shdr
    ra.Value = arr
    ra.EntireColumn.Columns.AutoFit
End Sub

'ランク列追加
Sub AddRankColumn(ra As Range, bhdr As Boolean, Optional shdr As String = "No.")
    Dim arr() As Variant
    ReDim arr(0 To ra.Rows.Count, 1 To 1)
    
    arr(0, 1) = shdr
    Dim i As Long, j As Long
    If shdr <> "" Then j = 1
    For i = 1 To ra.Rows.Count
        arr(j, 1) = i
        j = j + 1
    Next i
    ra.Value = arr
    ra.EntireColumn.Columns.AutoFit
End Sub

'----------------------------------------
'値追加
'----------------------------------------

Sub Cells_GenerateValue(ra As Range, mode As Long)
    ScreenUpdateOff
    '
    Dim idx As Long
    idx = 1
    Dim rb As Range
    For Each rb In ra.Areas
        If rb Is Nothing Then Exit For
        '
        Dim va As Variant
        va = RangeToFormula2(rb)
        '
        Select Case mode
        Case Else
            Cells_GenerateIndex va, idx
        End Select
        '
        rb.Value = va
    Next rb
    '
    ScreenUpdateOn
End Sub

Private Sub Cells_GenerateIndex(va As Variant, v0 As Long)
    Dim i As Long
    Dim r As Long, c As Long
    'header
    r = LBound(va, 1)
    For c = LBound(va, 2) To UBound(va, 2)
        va(r, c) = ColumnName(c)
        i = i + 1
    Next c
    'data
    i = v0
    For r = LBound(va, 1) + 1 To UBound(va, 1)
        For c = LBound(va, 2) To UBound(va, 2)
            va(r, c) = i
            i = i + 1
        Next c
    Next r
    v0 = i
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

