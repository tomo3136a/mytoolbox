Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI
Private g_select As Integer

'----------------------------------------
'ribbon helper
'----------------------------------------

'リボンを更新
Private Sub RefreshRibbon(Optional id As String)
    If g_ribbon Is Nothing Then
    ElseIf id = "" Then g_ribbon.Invalidate
    Else: g_ribbon.InvalidateControl id
    End If
    DoEvents
End Sub

'リボンID番号取得
Private Function RibbonID(control As IRibbonControl, Optional n As Long) As Long
    Dim vs As Variant
    vs = Split(re_replace(control.id, "[^0-9.]", ""), ".")
    If UBound(vs) >= n Then RibbonID = val("0" & vs(UBound(vs) - n))
End Function

'----------------------------------------
'イベント
'----------------------------------------

'起動時実行
Private Sub Designer_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    Draw_ResetParam
    IDF_ResetParam
End Sub

'テキスト入力
Private Sub Designer_onChange(ByRef control As IRibbonControl, ByRef text As String)
    Call Draw_SetParam(RibbonID(control), text)
End Sub

Private Sub Designer_getText(ByRef control As IRibbonControl, ByRef text As Variant)
    text = Draw_GetParam(RibbonID(control))
End Sub

'チェックボックス
Private Sub Designer_onActionPressed(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Call Draw_SetParam(RibbonID(control), IIf(pressed, 1, 0))
End Sub

Private Sub Designer_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Draw_IsParamFlag(RibbonID(control))
End Sub

'----------------------------------------
'■メニュー
'  1.x: 図形操作機能
'  2.x: 図形リスト
'  3.x: 図形の更新
'  4.x: 部品配置機能
'  5.x: 作図機能(IDF)
'----------------------------------------

Private Sub Designer_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 1: Draw1_Menu RibbonID(control)
    Case 2: Draw2_Menu RibbonID(control)
    Case 3: Draw3_Menu RibbonID(control)
    End Select
End Sub

'----------------------------------------
'■機能グループ1
'図形操作機能
'----------------------------------------

Private Sub Designer1_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 1: Draw1_Menu RibbonID(control)
    Case 2: Draw2_Menu RibbonID(control)
    Case 3: Draw3_Menu RibbonID(control)
    End Select
End Sub

Private Sub Draw1_Menu(id As Long, Optional opt As Variant)
    Select Case id
    Case 1: SetDefaultShapeStyle        '標準図形設定
    Case 2:
    Case 3: RemoveSharps                '図形削除
    Case 4: ConvertToPicture            '図形を絵に変換
    Case 5: SetTextBoxStyle             'テキストボックス基本設定
    Case 6: ToggleVisible 0             '塗りつぶし表示ON/OFF
    Case 7: ToggleVisible 3             '3D表示ON/OFF
    Case 8: OriginAlignment             '原点合わせ
    Case 9: UpdateShapeName ActiveSheet '図形名一括更新
    Case 10: FlipShapes                 '表裏反転
    End Select
End Sub

'----------------------------------------
'■機能グループ2
'図形操作機能
'  1x: 図形リスト
'  2x: 図形の更新
'----------------------------------------

Private Sub Designer2_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 2: Draw2_Menu RibbonID(control)
    Case 3: Draw3_Menu RibbonID(control)
    End Select
End Sub

Private Sub Draw2_Menu(id As Long, Optional opt As Variant)
    Select Case id
    Case 1: ListShapeInfo               '一覧取得
    Case 2: AddShapeListName            '名前追加
    Case 3: ApplyShapeInfo ActiveCell   '図形情報適用
    Case 4: SelectShapeName             '図形名選択
    Case 5: UpdateShapeInfo             'データ取得
    End Select
End Sub

'----------------------------------------
''■機能グループ3
'
'----------------------------------------

Private Sub Draw3_Menu(id As Long, Optional opt As Variant)
    AddListShapeHeader ActiveCell, id   'ヘッダ項目追加
End Sub

'----------------------------------------
''■機能グループ4
'部品配置機能
'----------------------------------------

'図形アイテム
Private Sub Designer4_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 4: Draw4_Menu RibbonID(control)
    End Select
    RefreshRibbon "c41"
End Sub

Private Sub Draw4_Menu(id As Long, Optional opt As Variant)
    Select Case id
    Case 1: AddDrawItem             '配置
    Case 2: CopyDrawItem            'コピー
    Case 3: DrawItemEntry           '登録
    Case 4: DrawItemDelete          '削除
    Case 5: DuplicateDrawItemSheet  '設定ローカル化
    Case 6: ImportDrawItemSheet     '設定シート取込
    'Case 7: ToggleAddinBook
    Case Else:
    End Select
End Sub

'部品選択
Private Sub Designer4_getItemCount(control As IRibbonControl, ByRef returnedVal)
    Dim cnt As Long
    DrawItemCount cnt
    If cnt < 1 Then cnt = 1
    returnedVal = cnt
End Sub

Private Sub Designer4_getItemID(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    returnedVal = Index
End Sub

Private Sub Designer4_getItemLabel(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    Dim s As String
    DrawItemName Index, s
    returnedVal = s
End Sub

Private Sub Designer4_getSelectedItemID(control As IRibbonControl, ByRef returnedVal)
    Dim cnt As Long
    DrawItemCount cnt
    If cnt > 0 Then cnt = cnt - 1
    If g_select > cnt Then g_select = cnt
    DrawItemSelect g_select
    returnedVal = g_select
End Sub

Private Sub Designer4_onActionDropDown(control As IRibbonControl, id As String, Index As Integer)
    g_select = Index
    DrawItemSelect g_select
    If Not g_ribbon Is Nothing Then g_ribbon.InvalidateControl control.id
End Sub

'----------------------------------------
''■機能グループ5
'タイミング図
'----------------------------------------

Private Sub Designer5_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 5: Draw5_Menu RibbonID(control)
    End Select
End Sub

Private Sub Draw5_Menu(id As Long, Optional opt As Variant)
    Select Case id
    Case 1: GenerateTimeChart 1     'タイムチャートデータ作成(クロック)
    Case 2: GenerateTimeChart 2     'タイムチャートデータ作成(カウンタ)
    Case 3: GenerateTimeChart 3     'タイムチャートデータ作成(ロジック)
    Case 7: AddDrawingSheet         '方眼紙シート追加
    Case 8: DrawTimeChart 1         '作図
    Case 9: DrawTimeChart 2         '作図(図形)
    
    Case 5: GenerateTimeChart 5     '抽出
    Case 6: GenerateTimeChart 6     '結合

    Case 11: ApplyTimeChart 1       '反転
    Case 12: ApplyTimeChart 2       'クリア
    Case 13: ApplyTimeChart 3       'プリセット
    
    Case 21: GenerateTimeChart 1    'データ作成(クロック)
    Case 22: GenerateTimeChart 2    'データ作成(カウンタ)
    Case 23: GenerateTimeChart 3    'データ作成(ロジック)
    Case 24: GenerateTimeChart 4    'データ作成(抽出)
    Case 25: GenerateTimeChart 5    'データ作成(結合)
    
    Case 31: GenerateTimeChart 11   'データ作成(NOT)
    Case 32: GenerateTimeChart 12   'データ作成(AND)
    Case 33: GenerateTimeChart 13   'データ作成(OR)
    Case 34: GenerateTimeChart 14   'データ作成(XOR)
    Case 35: GenerateTimeChart 15   'データ作成(MUX)
    Case 36: GenerateTimeChart 16   'データ作成(D-FF)
    Case 37: GenerateTimeChart 17   'データ作成(SR-FF)
    Case 38: GenerateTimeChart 18   'データ作成(SYNC)
    Case 39: GenerateTimeChart 19   'データ作成(EDGE)
    Case Else: HelpTimingChart
    End Select
End Sub

'----------------------------------------
''■機能グループ6
'作図機能(IDF)
'----------------------------------------

Private Sub Designer6_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 5: Draw6_Menu RibbonID(control)
    End Select
End Sub

Private Sub Draw6_Menu(id As Long, Optional opt As Variant)
    Select Case id
    Case 1: DrawIDF                     'IDF作図
    Case 2: AddSheetIDF                 'IDFシート追加
    Case 3: MacroIDF                    'IDFマクロ
    Case 4: ImportIDF                   'IDFファイル読み込み
    Case 5: ExportIDF                   'IDFファイル書き出し
    Case 6: AddRecordIDF                'IDF行追加
    Case 7: AddRecordIDF 1              'IDF行追加
    Case 8: AddRecordIDF 2              'IDF行追加
    Case 10: ResetShapeSize             'サイズ修正
    Case 11: ResizeShapeScale           'スケール変更
    End Select
End Sub

'テキスト入力
Private Sub Designer6_onChange(ByRef control As IRibbonControl, ByRef text As String)
    Call IDF_SetParam(RibbonID(control), text)
End Sub

Private Sub Designer6_getText(ByRef control As IRibbonControl, ByRef text As Variant)
    text = IDF_GetParam(RibbonID(control))
End Sub

'チェックボックス
Private Sub Designer6_onActionPressed(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Call IDF_SetFlag(RibbonID(control), pressed)
End Sub

Private Sub Designer6_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = IDF_IsFlag(RibbonID(control))
End Sub

'----------------------------------------
''■機能グループn
'拡張機能
'----------------------------------------

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
    If TypeName(obj) = "Range" Then Exit Sub
    '
    Dim sz As Integer
    Set shra = obj.ShapeRange
    
    sz = shra.TextFrame2.TextRange.Characters.Count
    If sz > 0 Then
        shra.TextFrame2.TextRange.Characters(sz, 1).Font.Superscript = True
    End If
    'obj.TextFrame.Characters.text = "test"
    
    'MsgBox "Test1 macro"
End Sub

Sub Test2Macro(control As IRibbonControl)
    ToggleAddinBook
    RefreshRibbon "c41"
End Sub

Sub Test3Macro(control As IRibbonControl)
    Dim sn As String
    sn = "#shapes"
    Dim ws As Worksheet
    Set ws = GetSheet(sn)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = sn
    End If
    
End Sub

Sub Test4Macro(control As IRibbonControl)
End Sub

Sub HelpMacro(control As IRibbonControl)
End Sub

Sub FindMacro(control As IRibbonControl)
End Sub

