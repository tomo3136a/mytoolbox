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
    vs = Split(re_replace(control.id, "[^0-9.]", ""), ".")
    If UBound(vs) >= n Then RibbonID = Val("0" & vs(n))
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
    SetRtParam "path.1", True       'リンクあり
    SetRtParam "path.2", True       'フォルダあり
    SetRtParam "path.3", True       '再帰あり
    SetRtParam "info.sheet", True   'シート追加
    SetRtParam "mark.color", 0      'マーカカラーは黄色
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
    Select Case RibbonID(control, 1)
    Case 1:
        'レポートサイン
        If TypeName(Selection) <> "Range" Then Exit Sub
        ReportSign Selection
    Case 2:
        'ページフォーマット
        PagePreview
    Case 3:
        'テキスト変換
        If TypeName(Selection) <> "Range" Then Exit Sub
        TextConvProc Selection, RibbonID(control, 2)
    Case 4:
        '拡張書式
        If TypeName(Selection) <> "Range" Then Exit Sub
        UserFormatProc Selection, RibbonID(control, 2)
    Case 5:
        '定型式挿入
        If TypeName(Selection) <> "Range" Then Exit Sub
        UserFormulaProc Selection, RibbonID(control, 2)
    Case 6:
        '表示・非表示
        ShowHide RibbonID(control, 2)
    Case 7:
        'パス名
        If TypeName(Selection) <> "Range" Then Exit Sub
        PathProc Selection, RibbonID(control, 2)
    Case 8:
        '情報取得
        AddInfoTable RibbonID(control, 2)
    Case 9:
        'エクスポート
        If TypeName(Selection) <> "Range" Then Exit Sub
        ExportProc Selection, RibbonID(control, 2)
    End Select
End Sub

Private Sub works1_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Select Case RibbonID(control, 1)
    Case 7: SetRtParam "path." & RibbonID(control, 2), CStr(pressed)    'パス名
    Case 8: SetRtParam "info." & control.Tag, CStr(pressed)             '情報取得
    Case 9: SetRtParam "export." & RibbonID(control, 2), CStr(pressed)  'エクスポート
    End Select
End Sub

Private Sub works1_getPressed(control As IRibbonControl, ByRef returnedVal)
    Select Case RibbonID(control, 1)
    Case 7: returnedVal = GetRtParamBool("path." & RibbonID(control, 2))    'パス名
    Case 8: returnedVal = GetRtParamBool("info." & control.Tag)             '情報取得
    Case 9: returnedVal = GetRtParamBool("export." & RibbonID(control, 2))  'エクスポート
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
    Select Case RibbonID(control, 1)
    Case 1: Call WakuProc(Selection, RibbonID(control, 2))
    Case 2: Call SelectProc(Selection, RibbonID(control, 2))
    Case 3: Call AddColumn(Selection, RibbonID(control, 2))
    Case 7
        Select Case RibbonID(control, 2)
        Case 1: Call WakuProc(Selection, 7)    '囲いクリア
        Case 2: Call WakuProc(Selection, 8)    'データクリア
        Case 3: Call WakuProc(Selection, 9)    '表クリア
        Case Else                               '囲い・データクリア
            Call WakuProc(Selection, 7)
            Call WakuProc(Selection, 9)
        End Select
    Case 8
        SetTableMargin xlRows
        SetTableMargin xlColumns
        RefreshRibbon control.id
    End Select
    ScreenUpdateOn
End Sub

Private Sub works2_getLabel(ByRef control As Office.IRibbonControl, ByRef label As Variant)
   label = "行: " & GetTableMargin(xlRows) & ", 列: " & GetTableMargin(xlColumns)
End Sub

'----------------------------------------
'■機能グループ3
'テンプレート機能
'----------------------------------------

Private Sub works3_onAction(ByVal control As IRibbonControl)
    Call TemplateMenu(RibbonID(control, 1))
    Select Case RibbonID(control, 1)
    'Case 8 '更新
    '    RBTable_Init
    '    g_ribbon.InvalidateControl "b4.1"
    '    g_ribbon.InvalidateControl "b4.2"
    '    g_ribbon.InvalidateControl "b4.3"
    '    g_ribbon.InvalidateControl "b4.4"
    '    g_ribbon.InvalidateControl "b4.5"
    '    g_ribbon.InvalidateControl "b4.6"
    '    g_ribbon.InvalidateControl "b4.7"
    '    g_ribbon.InvalidateControl "b4.8"
    '    g_ribbon.InvalidateControl "b4.9"
    Case 9 '開発
        g_ribbon.InvalidateControl "b3.9"
    End Select
End Sub

Private Sub works3_getLabel(ByRef control As IRibbonControl, ByRef label As Variant)
    If ThisWorkbook.IsAddin Then label = "ブック開く" Else label = "ブック閉じる"
    RefreshRibbon
    DoEvents
End Sub

Private Sub works3_getEnabled(control As IRibbonControl, ByRef enable As Variant)
    enable = ThisWorkbook.IsAddin
End Sub

'----------------------------------------
'■機能グループ4
'marker
'----------------------------------------

Private Sub works4_getLabel(control As IRibbonControl, ByRef label As Variant)
    Dim ss() As String
    ss = Split("黄色,赤,青,薄緑,灰色,橙,青緑,淡い橙,紫,緑", ",")
    label = ss(Val(GetRtParam("mark.color")))
End Sub

Private Sub works4_onGetImage(control As IRibbonControl, ByRef bitmap As Variant)
    Dim id As Integer
    id = ((Val(GetRtParam("mark.color")) + 9) Mod 10) + 1
    bitmap = "AppointmentColor" & id
End Sub

Private Sub works4_onAction(control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 1
        If TypeName(Selection) <> "Range" Then Exit Sub
        Call AddMarker(Selection, Val(GetRtParam("mark.color")))
    Case 3
        Call ListMarker
    Case 4
        If TypeName(Selection) <> "Range" Then Exit Sub
        Call DelMarker(Selection)
    End Select
End Sub

Private Sub works4_onSelected(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    Call SetRtParam("mark.color", "" & selectedIndex)
    RefreshRibbon
    DoEvents
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call AddMarker(Selection, Val(GetRtParam("mark.color")))
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
    Call RevProc(Selection, RibbonID(control, 1), res)
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
    Select Case RibbonID(control, 1)
    Case 1: Call TestProc(Selection, RibbonID(control, 1))
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

