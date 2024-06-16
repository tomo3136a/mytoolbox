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

'ID番号取得
Private Function RB_ID(control As IRibbonControl) As Integer
    RB_ID = val(Right(control.id, 1))
End Function

'TAG番号取得
Private Function RB_TAG(control As IRibbonControl) As Integer
    RB_TAG = val(control.Tag)
End Function

'リボンID番号取得
Private Function RibbonID(control As IRibbonControl) As Integer
    Dim s As String
    s = control.Tag
    If s = "" Then s = control.id
    Dim vs As Variant
    vs = Split(s, ".")
    If UBound(vs) >= 0 Then
        RibbonID = val(vs(UBound(vs)))
        Exit Function
    End If
    vs = Split(s, "_")
    If UBound(vs) >= 0 Then
        RibbonID = val(vs(UBound(vs)))
        Exit Function
    End If
    RibbonID = val(s)
End Function

'リボンを更新
Private Sub RefreshRibbon()
    If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
    DoEvents
End Sub

'----------------------------------------
'Initialize
'----------------------------------------

Private Sub works_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    '
    'ショートカットキー設定
    Application.OnKey "{F1}"
    Application.OnKey "{F1}", "works_ShortcutKey"
    '
    Application.OnKey "+{F1}"
    Application.OnKey "+{F1}", "works_ShortcutKey2"
    '
    'Application.OnKey "{F10}"
    'Application.OnKey "{F10}", "works_ShortcutKey3"
    '
    RBTable_Init
    SetParam "path", 1, True
    SetParam "path", 2, True
    SetParam "path", 3, True
    SetParam "info", 1, True
End Sub

Private Sub works_ShortcutKey()
    If Not g_ribbon Is Nothing Then g_ribbon.ActivateTab "TabWorks"
End Sub

Private Sub works_ShortcutKey2()
    'SendKeys "% h"
End Sub

Private Sub works_ShortcutKey3()
    'SendKeys "% h"
End Sub

'----------------------------------------
'1x:レポート機能
'----------------------------------------

'レポートサイン
Private Sub works11_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call ReportSign(Selection)
End Sub

'ページフォーマット
Private Sub works12_onAction(ByVal control As IRibbonControl)
    Call PagePreview
End Sub

'テキスト変換
Private Sub works13_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call MenuTextConv(RibbonID(control), Selection)
End Sub

'拡張書式
Private Sub works14_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call MenuUserFormat(RibbonID(control), Selection)
End Sub

'定型式挿入
Private Sub works15_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call MenuUserFormula(RibbonID(control), Selection)
End Sub

'表示・非表示
Private Sub works16_onAction(ByVal control As IRibbonControl)
    Call ShowHide(RibbonID(control))
End Sub

'パス名
Private Sub works17_onAction(ByVal control As IRibbonControl)
    Call PathMenu(RibbonID(control), Selection)
End Sub

Private Sub works17_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetParam "path", RibbonID(control), pressed
End Sub

Private Sub works17_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetParamBool("path", RibbonID(control))
End Sub

'情報取得
Private Sub works18_onAction(ByVal control As IRibbonControl)
    Call AddInfoSheet(RibbonID(control))
End Sub

Private Sub works18_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetParam "info", RibbonID(control), pressed
End Sub

Private Sub works18_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetParamBool("info", RibbonID(control))
End Sub

'エクスポート
Private Sub works19_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call MenuExport(Selection, RibbonID(control))
End Sub

Private Sub works19_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    'Call SetExportParam(RibbonID(control), pressed)
    SetParam "export", RibbonID(control), pressed
End Sub

Private Sub works19_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetParamBool("export", RibbonID(control))
End Sub

'----------------------------------------
'2x:罫線枠
'----------------------------------------

'移動・選択
Private Sub works21_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    SelectTable RibbonID(control), Selection
    Application.ScreenUpdating = True
End Sub

'枠設定
Private Sub works22_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Call TableWaku(RibbonID(control), Selection)
    Application.ScreenUpdating = True
End Sub

'囲いクリア
Private Sub works27_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Select Case RibbonID(control)
    Case 1
        '囲いクリア
        Call TableWaku(7, Selection)
    Case 2
        'データクリア
        Call TableWaku(8, Selection)
    Case 3
        '表クリア
        Call TableWaku(9, Selection)
    Case Else
        '囲い・データクリア
        Call TableWaku(7, Selection)
        Call TableWaku(9, Selection)
    End Select
    Application.ScreenUpdating = True
End Sub

'マージン表示・設定
Private Sub works28_onAction(ByVal control As IRibbonControl)
    SetTableMargin xlRows
    SetTableMargin xlColumns
    g_ribbon.InvalidateControl control.id
End Sub

Private Sub works28_getLabel(ByRef control As Office.IRibbonControl, ByRef label As Variant)
   label = "行: " & GetTableMargin(xlRows) & ", 列: " & GetTableMargin(xlColumns)
End Sub

'----------------------------------------
'3x:テンプレート機能
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
'marker
'----------------------------------------

Private Sub works4_onAction(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call Marker(RibbonID(control), Selection)
End Sub

'----------------------------------------
'common
'----------------------------------------

Private Sub works5_onAction(control As IRibbonControl)
    Call RBTable_onAction(RibbonID(control))
End Sub

Private Sub works5_getVisible(control As IRibbonControl, ByRef Visible As Variant)
    Call RBTable_getVisible(RibbonID(control), Visible)
End Sub

Private Sub works5_getLabel(control As IRibbonControl, ByRef label As Variant)
    Call RBTable_getLabel(RibbonID(control), label)
End Sub

Private Sub works5_onGetImage(control As IRibbonControl, ByRef bitmap As Variant)
    Call RBTable_onGetImage(RibbonID(control), bitmap)
End Sub

Private Sub works5_getSize(control As IRibbonControl, ByRef Size As Variant)
    Call RBTable_getSize(RibbonID(control), Size)
End Sub

