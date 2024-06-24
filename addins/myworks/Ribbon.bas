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
    RB_ID = Val(Right(control.id, 1))
End Function

'TAG番号取得
Private Function RB_TAG(control As IRibbonControl) As Integer
    RB_TAG = Val(control.Tag)
End Function

'リボンID番号取得
Private Function RibbonID(control As IRibbonControl) As Integer
    Dim s As String
    s = control.Tag
    If s = "" Then s = control.id
    Dim vs As Variant
    vs = Split(s, ".")
    If UBound(vs) >= 0 Then
        RibbonID = Val(vs(UBound(vs)))
        Exit Function
    End If
    vs = Split(s, "_")
    If UBound(vs) >= 0 Then
        RibbonID = Val(vs(UBound(vs)))
        Exit Function
    End If
    RibbonID = Val(s)
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
    Application.OnKey "{F1}", "works_ShortcutKey1"
    '
    Application.OnKey "+{F1}"
    Application.OnKey "+{F1}", "works_ShortcutKey2"
    '
    '初期化
    SetParam "path", 1, True        'リンクあり
    SetParam "path", 2, True        'フォルダあり
    SetParam "path", 3, True        '再帰あり
    SetParam "info", 1, True        'シート追加
    SetParam "mark", "color", 10    'マーカカラーは黄色
End Sub

Private Sub works_ShortcutKey1()
    'worksタブに移動
    If Not g_ribbon Is Nothing Then g_ribbon.ActivateTab "TabWorks"
End Sub

Private Sub works_ShortcutKey2()
    'homeタブに移動
    SendKeys "%"
    SendKeys "H"
    SendKeys "%"
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
'4:marker
'----------------------------------------

Private Sub works4_getLabel(control As IRibbonControl, ByRef label As Variant)
    Dim name() As Variant
    name = Array("-", "赤", "青", "緑", "灰色", "橙", "青緑", "淡い橙", "紫", "緑", "黄色")
    label = name(Val(GetParam("mark", "color")))
End Sub

Private Sub works4_onGetImage(control As IRibbonControl, ByRef bitmap As Variant)
    bitmap = "AppointmentColor" & Val(GetParam("mark", "color"))
End Sub

Private Sub works4_onAction(control As IRibbonControl)
    Select Case RibbonID(control)
    Case 1
        If TypeName(Selection) <> "Range" Then Exit Sub
        Call AddMarker(Selection, Val(GetParam("mark", "color")))
    Case 2
        If TypeName(Selection) <> "Range" Then Exit Sub
        Call ListMarker(Selection)
    Case 3
        If TypeName(Selection) <> "Range" Then Exit Sub
        Dim ce As Range
        For Each ce In Selection.Cells
            Call DelMarker(ce.Value)
            ce.Clear
        Next ce
    Case 4
        Call DelMarkerAll
    End Select
End Sub

Private Sub works41_onAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    Call SetParam("mark", "color", Mid(selectedId, InStr(1, selectedId, ".") + 1))
    If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
    DoEvents
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call AddMarker(Selection, Val(GetParam("mark", "color")))
End Sub

'----------------------------------------
'5:revision mark
'----------------------------------------

Private Sub works5_getLabel(control As IRibbonControl, ByRef label As Variant)
    Call GetRevMark(label)
End Sub

Private Sub works5_onAction(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    '
    Dim rev As String
    Select Case RibbonID(control)
    Case 1
        '版数マーク追加
        Call AddRevMark(Selection)
    Case 2
        '版数設定
        Call GetRevMark(rev)
        rev = InputBox("版数を入力してください。", "版数マーク設定", rev)
        If rev = "" Then Exit Sub
        Call SetRevMark(rev)
        If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
        DoEvents
        Call AddRevMark(Selection)
    Case 3
        '版数リスト作成
        Call GetRevMark(rev)
        rev = InputBox("版数を入力してください。", "版数マークリスト", rev)
        If rev = "" Then Exit Sub
        Call ListRevMark(Selection, rev)
    End Select
End Sub

