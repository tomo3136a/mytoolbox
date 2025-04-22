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
    If UBound(vs) >= n Then RibbonID = Val("0" & vs(UBound(vs) - n))
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
    SetRtParam "path", 1, True          'リンクあり
    SetRtParam "path", 2, True          'フォルダあり
    SetRtParam "path", 3, True          '再帰あり
    SetRtParam "info", "sheet", True    'シート追加
    SetRtParam "mark", "color", 0       'マーカカラーは黄色
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

'レポートサイン
Private Sub works11_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    ReportSign Selection
End Sub

'ページフォーマット
Private Sub works12_onAction(ByVal control As IRibbonControl)
    PagePreview
End Sub

'テキスト変換
Private Sub works13_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    MenuTextConv Selection, RibbonID(control)
End Sub

'拡張書式
Private Sub works14_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    MenuUserFormat Selection, RibbonID(control)
End Sub

'定型式挿入
Private Sub works15_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    MenuUserFormula Selection, RibbonID(control)
End Sub

'表示・非表示
Private Sub works16_onAction(ByVal control As IRibbonControl)
    ShowHide RibbonID(control)
End Sub

'パス名
Private Sub works17_onAction(ByVal control As IRibbonControl)
    PathMenu Selection, RibbonID(control)
End Sub

Private Sub works17_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetRtParam "path", RibbonID(control), CStr(pressed)
End Sub

Private Sub works17_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetRtParamBool("path", RibbonID(control))
End Sub

'情報取得
Private Sub works18_onAction(ByVal control As IRibbonControl)
    AddInfoTable RibbonID(control)
End Sub

Private Sub works18_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetRtParam "info", control.Tag, CStr(pressed)
End Sub

Private Sub works18_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetRtParamBool("info", control.Tag)
End Sub

'エクスポート
Private Sub works19_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    MenuExport Selection, RibbonID(control)
End Sub

Private Sub works19_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetRtParam "export", RibbonID(control), CStr(pressed)
End Sub

Private Sub works19_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetRtParamBool("export", RibbonID(control))
End Sub

'----------------------------------------
''■機能グループ2
'罫線枠
'----------------------------------------

'移動・選択
Private Sub works21_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    TableSelect Selection, RibbonID(control)
End Sub

'枠設定
Private Sub works22_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Call TableWaku(Selection, RibbonID(control))
    Application.ScreenUpdating = True
End Sub

'列追加
Private Sub works23_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Call AddColumn(Selection, RibbonID(control))
    Application.ScreenUpdating = True
End Sub

'囲いクリア
Private Sub works27_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Select Case RibbonID(control)
    Case 1
        '囲いクリア
        Call TableWaku(Selection, 7)
    Case 2
        'データクリア
        Call TableWaku(Selection, 8)
    Case 3
        '表クリア
        Call TableWaku(Selection, 9)
    Case Else
        '囲い・データクリア
        Call TableWaku(Selection, 7)
        Call TableWaku(Selection, 9)
    End Select
    Application.ScreenUpdating = True
End Sub

'マージン表示・設定
Private Sub works28_onAction(ByVal control As IRibbonControl)
    SetTableMargin xlRows
    SetTableMargin xlColumns
    RefreshRibbon control.id
End Sub

Private Sub works28_getLabel(ByRef control As Office.IRibbonControl, ByRef label As Variant)
   label = "行: " & GetTableMargin(xlRows) & ", 列: " & GetTableMargin(xlColumns)
End Sub

'----------------------------------------
'■機能グループ3
'テンプレート機能
'----------------------------------------

Private Sub works3_onAction(ByVal control As IRibbonControl)
    Call TemplateMenu(RibbonID(control))
    Select Case RibbonID(control)
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
    g_ribbon.Invalidate
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
    Dim name() As String
    name = Split("黄色,赤,青,薄緑,灰色,橙,青緑,淡い橙,紫,緑", ",")
    label = name(Val(GetRtParam("mark", "color")))
End Sub

Private Sub works4_onGetImage(control As IRibbonControl, ByRef bitmap As Variant)
    Dim id As Integer
    id = ((Val(GetRtParam("mark", "color")) + 9) Mod 10) + 1
    bitmap = "AppointmentColor" & id
End Sub

Private Sub works4_onAction(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Select Case RibbonID(control)
    Case 1
        Call AddMarker(Selection, Val(GetRtParam("mark", "color")))
    Case 2
        Call ListMarker(Selection)
    Case 3
        ScreenUpdateOff
        Dim ce As Range
        For Each ce In Selection.Cells
            Call DelMarker(ce.Value)
            ce.Clear
        Next ce
        ScreenUpdateOn
    End Select
End Sub

Private Sub works41_onAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    Call SetRtParam("mark", "color", Mid(selectedId, 1 + InStr(1, selectedId, ".")))
    If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
    DoEvents
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call AddMarker(Selection, Val(GetRtParam("mark", "color")))
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
    If Not res Then Exit Sub
    RefreshRibbon
    Call RevProc(Selection, 1)
End Sub

