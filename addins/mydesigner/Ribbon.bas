Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI

'----------------------------------------
'common
'----------------------------------------

Private Function SelectRange() As Range
    Dim ra As Range
    If TypeName(Selection) = "Range" Then
        Dim s As String
        s = TypeName(Selection)
        Set ra = Selection
    Else
        Set ra = Range(Selection.TopLeftCell, Selection.BottomRightCell)
        ra.Select
    End If
    Set SelectRange = ra
End Function

'----------------------------------------
'ribbon helper
'----------------------------------------

Private Sub RefreshRibbon()
    If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
    DoEvents
End Sub

Private Function RibbonID(control As IRibbonControl) As Integer
    Dim s As String
    s = control.Tag
    If s = "" Then s = control.id
    Dim vs As Variant
    vs = Split(s, ".")
    If UBound(vs) >= 0 Then
        RibbonID = val("0" & vs(UBound(vs)))
        Exit Function
    End If
    vs = Split(s, "_")
    If UBound(vs) >= 0 Then
        RibbonID = val("0" & vs(UBound(vs)))
        Exit Function
    End If
End Function

'----------------------------------------
'設定
'----------------------------------------

Private Sub Designer_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    Call ResetDrawParam
End Sub

Private Sub Designer_onChange(ByRef control As IRibbonControl, ByRef text As String)
    Call SetDrawParam(RibbonID(control), text)
End Sub

Private Sub Designer_getText(ByRef control As IRibbonControl, ByRef text As Variant)
    text = GetDrawParam(RibbonID(control))
End Sub

Private Sub Designer_onAction(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Dim v As Integer
    If pressed Then v = 1
    Call SetDrawParam(RibbonID(control), v)
End Sub

'----------------------------------------
'1x. 図形操作機能
'----------------------------------------

Private Sub Designer1_onAction(ByVal control As IRibbonControl)
    Dim obj As Object
    Set obj = Selection
    '
    Select Case RibbonID(control)
    Case 1
        '図形のリストアップ
        Call ListShape(ActiveCell, ActiveSheet, "")
    Case 2
        '図形の更新
        Call UpdateShape(ActiveCell)
    Case 3
        '図形を全て削除
        Call RemoveSharp(ActiveSheet)
    Case 4
        '図形を絵に変換
        Call ConvToPic
    Case 5
        '図形基本設定
        Call SetShapeStyle
    Case 6
        'オブジェクトの初期化
        Call DefaultShapeSetting
    Case 7
        '装飾②
        Dim v As Variant
        For Each v In ActiveSheet.Shapes
            Dim sh As Shape
            Set sh = v
            Dim u As Variant
            For Each u In sh.GroupItems
                Dim s As String
                s = u.name
            Next u
        Next v
    Case 8
        '左右反転
        obj.ShapeRange.Flip msoFlipHorizontal
    Case 9
        '上下反転
        obj.ShapeRange.Flip msoFlipVertical
    End Select
End Sub

'----------------------------------------
'2x. 作図機能(部品)
'----------------------------------------

'図形アイテム
Private Sub Designer2_onAction(ByVal control As IRibbonControl)
    On Error Resume Next
    Call DrawGraphItem(RibbonID(control), SelectRange)
    On Error GoTo 0
End Sub

'----------------------------------------
'3x. 作図機能(IDF)
'----------------------------------------

Private Sub Designer3_onAction(ByVal control As IRibbonControl)
    Dim ce As Range
    Set ce = ActiveCell
    '
    Select Case RibbonID(control)
    Case 1
        'IDFファイル読み込み
        Call ImportIDF
    Case 2
        'IDFファイル書き出し
        Call ExportIDF(ActiveSheet)
    Case 3
        'IDF作図
        Call DrawIDF(ce.Worksheet, ce.Left, ce.Top)
    Case 4
        'IDF作図
        Call DrawIDF(ce.Worksheet, ce.Left, ce.Top, sheet_load:=True)
    Case 5
        'IDF作図
        Call DrawIDF(ce.Worksheet, ce.Left, ce.Top, sheet_load:=True)
    End Select
End Sub

