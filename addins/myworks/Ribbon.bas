Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI

'----------------------------------------
'ribbon helper
'----------------------------------------

'リボンを更新
Private Sub RefreshRibbon(Optional id As String)
    If Not g_ribbon Is Nothing Then
        If id = "" Then
            g_ribbon.Invalidate
        Else
            g_ribbon.InvalidateControl id
        End If
    End If
    DoEvents
End Sub

'リボンID番号取得
Private Function RibbonID(control As IRibbonControl, Optional n As Long) As Long
    Dim vs As Variant
    vs = Split(control.id, ".")
    If UBound(vs) > n Then RibbonID = Val("0" & vs(n + 1))
End Function

'----------------------------------------
'イベント
'----------------------------------------

'起動時実行
Private Sub works_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    '
    'ショートカットキー設定
    On Error Resume Next
    Application.OnKey "{F1}"
    Application.OnKey "{F1}", "works_ShortcutKey1"
    '
    Application.OnKey "+{F1}"
    Application.OnKey "+{F1}", "works_ShortcutKey2"
    On Error GoTo 0
    '
    '初期化
    SetRtBool "path.1", True        'リンクあり
    SetRtBool "path.2", True        'フォルダあり
    SetRtBool "path.3", True        '再帰あり
    SetRtBool "info.sheet", True    'シート追加
    StrNum "mark.color", 0       'マーカカラーは黄色
End Sub

Private Sub works_ShortcutKey1()
    'worksタブに移動
    If g_ribbon Is Nothing Then Exit Sub
    g_ribbon.ActivateTab "TabWorks"
End Sub

Private Sub works_ShortcutKey2()
    'homeタブに移動
    SendKeys "%"
    SendKeys "H"
    SendKeys "%"
End Sub

'----------------------------------------
'■機能グループ1
'レポート機能
'----------------------------------------

Private Sub works1_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control)
    Case 1: 'レポートサイン
        If TypeName(Selection) <> "Range" Then Exit Sub
        ReportSign Selection
    Case 2: 'ページフォーマット
        Select Case RibbonID(control, 1)
        Case 1: AddLastRow
        Case 2: AddLastColumn
        Case 3: ResetCellPos
        Case Else: PagePreview
        End Select
    Case 3: 'テキスト変換
        If RibbonID(control, 1) = 13 Then
            WriteBookKeys
            Exit Sub
        End If
        If TypeName(Selection) <> "Range" Then Exit Sub
        TextConvProc Selection, RibbonID(control, 1)
    Case 4: '拡張書式
        If TypeName(Selection) <> "Range" Then Exit Sub
        UserFormatProc Selection, RibbonID(control, 1)
    Case 5: '定型式挿入
        If TypeName(Selection) <> "Range" Then Exit Sub
        UserFormulaProc Selection, RibbonID(control, 1)
    Case 6: '表示・非表示
        ShowHide RibbonID(control, 1)
    Case 7: 'パス名
        If TypeName(Selection) <> "Range" Then Exit Sub
        PathProc Selection, RibbonID(control, 1)
    Case 8: '情報取得
        AddInfoTable RibbonID(control, 1)
    Case 9: 'エクスポート
        If TypeName(Selection) <> "Range" Then Exit Sub
        ExportProc Selection, RibbonID(control, 1)
    End Select
End Sub

Private Sub works1_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Select Case RibbonID(control)
    Case 3: SetBookBool "page." & RibbonID(control, 1), pressed   'テキスト変換
    Case 7: SetBookBool "path." & RibbonID(control, 1), pressed   'パス名
    Case 8: SetBookBool "info." & control.Tag, pressed            '情報取得
    Case 9: SetBookBool "export." & RibbonID(control, 1), pressed 'エクスポート
    End Select
End Sub

Private Sub works1_getPressed(control As IRibbonControl, ByRef returnedVal)
    Select Case RibbonID(control)
    Case 3: returnedVal = GetBookBool("page." & RibbonID(control, 1))    'テキスト変換
    Case 7: returnedVal = GetBookBool("path." & RibbonID(control, 1))    'パス名
    Case 8: returnedVal = GetBookBool("info." & control.Tag)             '情報取得
    Case 9: returnedVal = GetBookBool("export." & RibbonID(control, 1))  'エクスポート
    End Select
End Sub

'----------------------------------------
''■機能グループ2
'罫線枠
'----------------------------------------

'枠設定
Private Sub works2_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    ScreenUpdateOff
    Select Case RibbonID(control)
    Case 1: Call WakuProc(Selection, RibbonID(control, 1))
    Case 2: Call SelectProc(Selection, RibbonID(control, 1))
    Case 3: Call AddColumn(Selection, RibbonID(control, 1))
    Case 7:
        Select Case RibbonID(control, 1)
        Case 1: Call WakuProc(Selection, 7)    '囲いクリア
        Case 2: Call WakuProc(Selection, 8)    'データクリア
        Case 3: Call WakuProc(Selection, 9)    '表クリア
        Case Else                               '囲い・データクリア
            Call WakuProc(Selection, 7)
            Call WakuProc(Selection, 9)
        End Select
    Case 8:
        SetTableMargin xlRows
        SetTableMargin xlColumns
        RefreshRibbon control.id
    End Select
    ScreenUpdateOn
End Sub

Private Sub works2_getLabel(ByRef control As IRibbonControl, ByRef label As Variant)
   label = "行: " & GetTableMargin(xlRows) & ", 列: " & GetTableMargin(xlColumns)
End Sub

'----------------------------------------
'■機能グループ3
'テンプレート機能
'----------------------------------------

Private Sub works3_onAction(ByVal control As IRibbonControl)
    Call TemplateProc(RibbonID(control), RibbonID(control, 1))
End Sub

Private Sub works3_getEnabled(control As IRibbonControl, ByRef enable As Variant)
    enable = ThisWorkbook.IsAddin
End Sub

'----------------------------------------
'■機能グループ9
'アドイン機能
'----------------------------------------

Private Sub works9_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control)
    Case 1: RefreshRibbon
    End Select
End Sub

Private Sub works9_getLabel(ByRef control As IRibbonControl, ByRef label As Variant)
    If ThisWorkbook.IsAddin Then label = "ブック開く" Else label = "ブック閉じる"
    RefreshRibbon
    DoEvents
End Sub

Private Sub works9_getEnabled(control As IRibbonControl, ByRef enable As Variant)
    enable = ThisWorkbook.IsAddin
End Sub

'----------------------------------------
'■機能グループ4
'marker
'----------------------------------------

Private Sub works4_getLabel(control As IRibbonControl, ByRef label As Variant)
    Dim ss() As String
    ss = Split("黄色,赤,青,薄緑,灰色,橙,青緑,淡い橙,紫,緑", ",")
    label = ss(Val(GetRtStr("mark.color")))
End Sub

Private Sub works4_onGetImage(control As IRibbonControl, ByRef bitmap As Variant)
    Dim id As Integer
    id = ((Val(GetRtStr("mark.color")) + 9) Mod 10) + 1
    bitmap = "AppointmentColor" & id
End Sub

Private Sub works4_onAction(control As IRibbonControl)
    Select Case RibbonID(control)
    Case 1
        If TypeName(Selection) <> "Range" Then Exit Sub
        Call AddMarker(Selection, Val(GetRtStr("mark.color")))
    Case 3
        Call ListMarker
    Case 4
        If TypeName(Selection) <> "Range" Then Exit Sub
        Call DelMarker(Selection)
    End Select
End Sub

Private Sub works4_onSelected(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    Call SetRtStr("mark.color", "" & selectedIndex)
    RefreshRibbon
    DoEvents
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call AddMarker(Selection, Val(GetRtStr("mark.color")))
End Sub

'----------------------------------------
'■機能グループ5
'revision mark
'----------------------------------------

Private Sub works5_getLabel(control As IRibbonControl, ByRef label As Variant)
    Call GetRevMark(label)
End Sub

Private Sub works5_onAction(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim res As Long
    Call RevProc(Selection, RibbonID(control), res)
    RefreshRibbon
    If Not res Then Exit Sub
    Call RevProc(Selection, 1)
End Sub

'----------------------------------------
'■機能グループ6
'test
'----------------------------------------

Private Sub works6_onAction(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    'ScreenUpdateOff
    Select Case RibbonID(control)
    Case 1: Call Cells_GenerateValue(Selection, 1)
    Case 2: Call TestProc(Selection, RibbonID(control, 1))
    Case 3: Call TestProc(Selection, RibbonID(control, 1))
    Case 4: Call TestProc(Selection, RibbonID(control, 1))
    Case 5: Call TestProc(Selection, RibbonID(control, 1))
    End Select
    'ScreenUpdateOn
End Sub

Private Sub TestProc(ra As Range, id As Long)
    Select Case id
    Case 1: Call Cells_GenerateValue(ra, 1)
    End Select
End Sub

