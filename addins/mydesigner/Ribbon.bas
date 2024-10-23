Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI
Private g_select As Integer

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

Private Sub RefreshRibbon(Optional id As String)
    If g_ribbon Is Nothing Then Exit Sub
    If id = "" Then
        g_ribbon.Invalidate
    Else
        g_ribbon.InvalidateControl id
    End If
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

'起動時実行
Private Sub Designer_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    Call ResetDrawParam
End Sub

'テキスト入力
Private Sub Designer_onChange(ByRef control As IRibbonControl, ByRef text As String)
    Call SetDrawParam(RibbonID(control), text)
End Sub

Private Sub Designer_getText(ByRef control As IRibbonControl, ByRef text As Variant)
    text = GetDrawParam(RibbonID(control))
End Sub

'チェックボックス
Private Sub Designer_onAction(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Dim v As Integer
    If pressed Then v = 1
    Call SetDrawParam(RibbonID(control), v)
End Sub

'----------------------------------------
'1x. 図形操作機能
'----------------------------------------

Private Sub Designer1_onAction(ByVal control As IRibbonControl)
    Dim id As Integer: id = RibbonID(control)
    
    Select Case id \ 10
    Case 1                                  '図形リスト(1x)
        ListShapeInfo ActiveSheet, id Mod 10
    Case 2                                  '図形の更新(2x)
        Select Case id Mod 10
        Case 0: UpdateShapeInfo ActiveCell  '図形リスト反映
        Case 1: UpdateShapeName ActiveSheet '図形名修理
        End Select
    Case Else
        Select Case id
        Case 3: RemoveSharps                '図形を削除
        Case 4: ConvToPic                   '図形を絵に変換
        Case 5: SetShapeStyle               'テキストボックス基本設定
        Case 6: ToggleVisible 0             '塗りつぶし表示ON/OFF
        Case 7: ToggleVisible 3             '3D表示ON/OFF
        Case 8
            Dim s2 As String
            s2 = GetShapeProperty(Selection.ShapeRange, "zero")
        Case 9                              'オブジェクトの初期化
            'Call DefaultShapeSetting
            MsgBox TypeName(Selection)
        End Select
    End Select
End Sub

'----------------------------------------
'2x. ツール機能
'----------------------------------------

'図形アイテム
Private Sub Designer2_onAction(ByVal control As IRibbonControl)
    On Error Resume Next
    Call DrawGraphItem(RibbonID(control), SelectRange)
    On Error GoTo 0
    RefreshRibbon "c21"
End Sub

'部品選択
Private Sub Designer2_getItemCount(control As IRibbonControl, ByRef returnedVal)
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    returnedVal = ws.Shapes.Count
End Sub

Private Sub Designer2_getItemID(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = index
End Sub

Private Sub Designer2_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    returnedVal = ws.Shapes(1 + index).name
End Sub

Private Sub Designer2_getSelectedItemID(control As IRibbonControl, ByRef returnedVal)
    returnedVal = g_select
End Sub

Private Sub Designer2_onActionDropDown(control As IRibbonControl, id As String, index As Integer)
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    Call SetDrawParam(10, ws.Shapes(index + 1).name)
    g_select = index
    If Not g_ribbon Is Nothing Then g_ribbon.InvalidateControl control.id
End Sub


'ダイナミックメニュー
Private Sub Designer2_getMenuContent(ByVal control As IRibbonControl, ByRef returnedVal)
    Dim xml As String

    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
          "<button id=""but1"" imageMso=""Help"" label=""Test1"" onAction=""Test1Macro""/>" & _
          "<button id=""but2"" imageMso=""Help"" label=""Test2"" onAction=""Test2Macro""/>" & _
          "<button id=""but3"" imageMso=""Help"" label=""Test3"" onAction=""Test3Macro""/>" & _
          "<button id=""but4"" imageMso=""Help"" label=""Test4"" onAction=""Test4Macro""/>" & _
          "<button id=""but5"" imageMso=""Help"" label=""Help"" onAction=""HelpMacro""/>" & _
          "<button id=""but6"" imageMso=""FindDialog"" label=""Find"" onAction=""FindMacro""/>" & _
          "</menu>"

    returnedVal = xml
End Sub

Sub Test1Macro(control As IRibbonControl)
    Dim shra As ShapeRange
    Dim obj As Object
    On Error Resume Next
    Set obj = Selection
    On Error GoTo 0
    If obj Is Nothing Then Exit Sub
    '
    Dim sz As Integer
    Set shra = obj.ShapeRange
    
    sz = shra.TextFrame2.TextRange.Characters.Count
    shra.TextFrame2.TextRange.Characters(sz, 1).Font.Superscript = True
    'obj.TextFrame.Characters.text = "test"
    
    MsgBox "Test1 macro"
End Sub

Sub Test2Macro(control As IRibbonControl)
    MsgBox "Test2 macro"
End Sub

Sub Test3Macro(control As IRibbonControl)
    MsgBox "Test3 macro"
End Sub

Sub Test4Macro(control As IRibbonControl)
    MsgBox "Test4 macro"
End Sub

Sub HelpMacro(control As IRibbonControl)
    RefreshRibbon "A"
    MsgBox "Help macro"
End Sub

Sub FindMacro(control As IRibbonControl)
    RefreshRibbon
    MsgBox "Find macro"
End Sub

'----------------------------------------
'3x. 作図機能(IDF)
'----------------------------------------

Private Sub Designer3_onAction(ByVal control As IRibbonControl)
    Dim ce As Range: Set ce = ActiveCell
    
    Select Case RibbonID(control)
    Case 1: ImportIDF           'IDFファイル読み込み
    Case 2: ExportIDF           'IDFファイル書き出し
    Case 3: DrawIDF ce.Worksheet, ce.Left, ce.Top   'IDF作図
    Case 4: 'ListKeywordIDF      'IDF作図
    Case 5
    End Select
End Sub
