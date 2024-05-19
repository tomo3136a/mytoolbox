Attribute VB_Name = "MDraw"
Option Explicit
Option Private Module

'パラメータ
Private g_ignore As String      '無効検索文字
Private g_scale As Double       'スケール
Private g_axes As Double        '軸間隔
Private g_flag As Integer       'モード(0:,1:,2:,3:)
Private g_part As String        '部品名

Private ptype_col As Variant
Private ptypename As Variant
Private pshapetypename As Variant

'----------------------------------------
'パラメータ制御
'----------------------------------------

Public Sub ResetDrawParam(Optional id As Integer)
    If id = 0 Or id = 1 Then g_ignore = ""
    If id = 0 Or id = 2 Then g_scale = 1
    If id = 0 Or id = 3 Then g_axes = 10
    If id = 0 Or id = 4 Then g_flag = 0
    If id = 0 Or id = 10 Then g_part = ""
End Sub

Public Sub SetDrawParam(id As Integer, ByVal val As String)
    Select Case id
    Case 1
        g_ignore = val
    Case 2
        If val <= 0 Then
            MsgBox "比率の設定が間違っています。(設定値>0)" & Chr(10) _
                & "設定値： " & val
            Exit Sub
        End If
        g_scale = val
    Case 3
        If val <= 0 Then
            MsgBox "目盛りの設定が間違っています。(設定値>0)" & Chr(10) _
                & "設定値： " & val
            Exit Sub
        End If
        g_axes = val
    Case 4
        g_flag = (g_flag And (65535 - 1)) Or (val * 1)
    Case 5
        g_flag = (g_flag And (65535 - 2)) Or (val * 2)
    Case 6
        g_flag = (g_flag And (65535 - 4)) Or (val * 4)
    Case 10
        g_part = val
    End Select
End Sub

Public Function GetDrawParam(id As Integer) As String
    Select Case id
    Case 1
        GetDrawParam = g_ignore
    Case 2
        If g_scale <= 0 Then
            ResetDrawParam id
            MsgBox "比率の設定を初期化しました。(設定値" & g_scale & ")"
        End If
        GetDrawParam = g_scale
    Case 3
        If g_axes <= 0 Then
            MsgBox "目盛りの設定を初期化しました。(設定値" & g_scale & ")"
            ResetDrawParam id
        End If
        GetDrawParam = g_axes
    End Select
End Function

Public Function IsDrawParam(id As Integer) As Boolean
    IsDrawParam = ((g_flag And (65535 - (2 ^ (id - 4)))) <> 0)
End Function

'----------------------------------------
'図形属性制御
'----------------------------------------

Function GetShapeProp(sr As ShapeRange, k As String) As String
    Dim line As Variant
    For Each line In Split(sr.AlternativeText, Chr(10), , vbTextCompare)
        Dim kv As Variant
        kv = Split(line, ":", 2, vbTextCompare)
        If UBound(kv) > 0 Then
            If UCase(k) = UCase(Trim(kv(0))) Then
                GetShapeProp = Trim(kv(1))
                Exit For
            End If
        End If
    Next line
End Function

Sub SetShapeProp(sr As ShapeRange, k As String, v As String)
    Dim lines As Variant
    lines = Split(sr.AlternativeText, Chr(10), , vbTextCompare)
    Dim line As Variant
    Dim i As Integer
    For i = 0 To lines.Count
        line = lines(i)
        Dim kv As Variant
        kv = Split(line, ":", 2, vbTextCompare)
        If UBound(kv) > 0 Then
            If k = Trim(kv(0)) Then
                lines(i) = k & ":" & Trim(v)
                Exit For
            End If
        Else
            If k = Trim(kv(0)) Then
                lines(i) = ""
                Exit For
            End If
        End If
    Next i
    line = Replace(Join(lines, Chr(10)), Chr(10) & Chr(10), Chr(10))
    sr.AlternativeText = line
End Sub



'----------------------------------------

'図形装飾
Public Sub SetShapeStyle(Optional sr As ShapeRange)
    On Error Resume Next
    If sr Is Nothing Then Set sr = Selection.ShapeRange
    Dim ws As Worksheet
    Set ws = sr.Parent
    '
    Dim sh As Object
    Set sh = ws.Shapes(sr.name)
    '
    'テキスト設定
    With sh.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorBottom
        .HorizontalAnchor = msoAnchorNone
    End With
    With sh.TextFrame
        .VerticalOverflow = xlOartVerticalOverflowOverflow
        .HorizontalOverflow = xlOartHorizontalOverflowOverflow
    End With
    '
    sh.LockAspectRatio = msoTrue
    sh.Placement = xlFreeFloating
    Dim s As String
    s = sh.Left & "," & sh.Top & "," & sh.Width & "," & sh.Height
    sr.AlternativeText = s
    sr.Title = s
    '
    On Error GoTo 0
End Sub

Public Sub InvertFillVisible(Optional sr As ShapeRange)
    On Error Resume Next
    If sr Is Nothing Then Set sr = Selection.ShapeRange
    Dim ws As Worksheet
    Set ws = sr.Parent
    '
    Dim sh As Object
    Set sh = ws.Shapes(sr.name)
    With sr.Fill
        If .Visible = msoTrue Then
            .Visible = msoFalse
        Else
            .Visible = msoTrue
        End If
    End With
    '
    On Error GoTo 0
End Sub

'ターゲットシート取得
Private Function TargetSheet(s As String) As Worksheet
    Dim v As Variant
    Dim ws As Worksheet
    For Each v In ActiveWorkbook.Worksheets
        If v.name = s Then Set ws = v
    Next v
    If ws Is Nothing Then
        For Each v In ThisWorkbook.Worksheets
            If v.name = s Then ws = v
        Next v
    End If
    If ws Is Nothing Then Exit Function
    Set TargetSheet = ws
End Function


Public Function DrawParts(ws As Worksheet, x0 As Double, y0 As Double, s As String) As String
    Dim cs As Worksheet
    Set cs = TargetSheet("#shapes")
    If cs Is Nothing Then Exit Function
    Dim sh As Shape
    If g_part <> "" Then
        cs.Shapes(g_part).Copy
        ws.Paste
    End If
End Function

'アイテム作画
Public Function DrawGraphItem(id As Integer, Optional ra As Range) As String
    If ra Is Nothing Then Exit Function
    '画面チラつき防止処置
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Interactive = False
    Application.Cursor = xlWait
    '
    Select Case id
    Case 1
        '方眼紙描画
        DrawGraphItem = DrawGraph2(ra.Worksheet, ra.Left, ra.Top + ra.Height, ra.Width, ra.Height)
    Case 2
        '軸描画
        DrawGraphItem = DrawAxis2(ra.Worksheet, ra.Left, ra.Top + ra.Height, ra.Width, ra.Height, 10)
    Case 3
        '原点描画
        DrawGraphItem = DrawAxis2(ra.Worksheet, ra.Left, ra.Top + ra.Height, ra.Width, ra.Height, 10)
    Case 4
        '部品描画
        DrawGraphItem = DrawParts(ra.Worksheet, ra.Left, ra.Top + ra.Height, g_part)
    End Select
    If Not DrawGraphItem = "" Then ra.Worksheet.Shapes(DrawGraphItem).Select
    '
    '画面チラつき防止処置解除
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Function

'軸線作画
Public Function DrawAxis( _
        ws As Worksheet, x0 As Double, y0 As Double, _
        w As Double, h As Double, Optional r As Double = 10) As String
    DrawAxis = DrawAxis2(ws, x0, y0, w, h, r)
End Function

Private Function DrawAxis2( _
        ws As Worksheet, x0 As Double, y0 As Double, _
        w As Double, h As Double, r As Double) As String
    Dim ns As Collection
    Set ns = New Collection
    '
    Dim sh As Object
    Set sh = ws.Shapes.AddLine(x0, y0 + 2 * r, x0, y0 - h)
    sh.line.ForeColor.RGB = RGB(0, 0, 0)
    ns.Add sh.name
    Set sh = ws.Shapes.AddLine(x0 - 2 * r, y0, x0 + w, y0)
    sh.line.ForeColor.RGB = RGB(0, 0, 0)
    ns.Add sh.name
    If r > 0 Then
        Set sh = ws.Shapes.AddShape(msoShapeOval, x0 - r, y0 - r, 2 * r, 2 * r)
        sh.line.ForeColor.RGB = RGB(0, 0, 0)
        sh.Fill.Visible = msoFalse
        ns.Add sh.name
    End If
    '
    Set sh = ws.Shapes.Range(ToArray(ns)).Group
    sh.name = "軸線 " & sh.id
    DrawAxis2 = sh.name
End Function

'方眼紙作成
Public Function DrawGraph( _
        ws As Worksheet, x0 As Double, y0 As Double, _
        w As Double, h As Double) As String
    DrawGraph = DrawGraph2(ws, x0, y0, w, h)
End Function

Private Function DrawGraph2( _
        ws As Worksheet, x0 As Double, y0 As Double, _
        w As Double, h As Double) As String
    Dim dp As Double
    dp = GetDrawParam(2) * GetDrawParam(3)
    If dp < 0.1 Then
        MsgBox ("間隔が狭すぎます。(" & dp & ")")
        Exit Function
    End If
    '
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    x1 = x0
    y1 = y0 - h
    x2 = x0 + w
    y2 = y0
    '
    Dim ns As Collection
    Set ns = New Collection
    '
    Dim sh As Object
    Dim p As Double
    Dim i As Integer
    For i = 1 To Int(w / dp)
        p = x1 + dp * i
        Set sh = ws.Shapes.AddLine(p, y1, p, y2)
        If i Mod 10 <> 0 Then sh.line.DashStyle = msoLineRoundDot
        sh.line.Weight = 0.25
        sh.line.ForeColor.RGB = RGB(0, 0, 255)
        ns.Add sh.name
    Next i
    For i = 1 To Int(h / dp)
        p = y2 - dp * i
        Set sh = ws.Shapes.AddLine(x1, p, x2, p)
        If i Mod 10 <> 0 Then sh.line.DashStyle = msoLineRoundDot
        sh.line.Weight = 0.25
        sh.line.ForeColor.RGB = RGB(0, 0, 255)
        ns.Add sh.name
    Next i
    Set sh = ws.Shapes.AddLine(x1, y1, x1, y2)
    sh.line.ForeColor.RGB = RGB(0, 0, 0)
    ns.Add sh.name
    Set sh = ws.Shapes.AddLine(x1, y2, x2, y2)
    sh.line.ForeColor.RGB = RGB(0, 0, 0)
    ns.Add sh.name
    '
    Set sh = ws.Shapes.Range(ToArray(ns)).Group
    sh.line.Transparency = 0.5
    sh.name = "方眼紙 " & sh.id
    DrawGraph2 = sh.name
End Function

Public Function ConvToPic2(ws As Worksheet, sh As ShapeRange) As String
    Dim s As String
    Dim x As Double
    Dim y As Double
    s = sh.name
    x = sh.Left
    y = sh.Top
    Dim sp As Shape
    sp.cu
    Selection.Cut
    Call Selection.Worksheet.PasteSpecial(0)
    Selection.name = s
    Selection.Left = x
    Selection.Top = y
    ConvToPic2 = Selection.name
    Application.CutCopyMode = 0
End Function


'図形を絵に変換
Public Function ConvToPic() As String
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim sh As Object
    On Error Resume Next
    Set sh = Selection.ShapeRange
    On Error GoTo 0
    If Not sh Is Nothing Then
        Dim s As String
        Dim x As Double
        Dim y As Double
        s = sh.name
        x = sh.Left
        y = sh.Top
        Selection.Cut
        Dim ws As Worksheet
        Set ws = Selection.Worksheet
        Call ws.PasteSpecial(0)
        Selection.name = s
        Selection.Left = x
        Selection.Top = y
        ConvToPic = Selection.name
        Application.CutCopyMode = 0
    End If
    '
    Application.ScreenUpdating = fsu
End Function

'図形を全て削除
Public Sub RemoveSharp(ws As Worksheet)
    On Error Resume Next
    Application.ScreenUpdating = False
    '
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Dim sh As Shape
        Set sh = ws.Shapes(i)
        ws.Shapes(i).Delete
    Next i
    '
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub


'----------------------------------------

'オブジェクトの書き出し
Public Sub ListShape(ra As Range, ws As Worksheet, igptn As String)
    If Not TypeName(Selection) = "Range" Then Exit Sub
    If igptn = "" Then igptn = g_ignore
    If igptn = "" Then igptn = "^#"
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    
    ce.Value = "Name"
    ce.Offset(, 1) = "Top"
    ce.Offset(, 2) = "Left"
    ce.Offset(, 3) = "Height"
    ce.Offset(, 4) = "Width"
    ce.Offset(, 5) = "Rotation"
    ce.Offset(, 6) = "Hide"
    ce.Offset(, 7) = "Type"
    ce.Offset(, 8) = "AlternativeText"
    Set ce = ce.Offset(1)
    
    Dim sp As Shape
    For Each sp In ws.Shapes
        If igptn <> "" Then
            If Not re_test(sp.name, igptn) Then
                Call ListShape2(ce, sp, "", igptn)
                Set ce = ce.Offset(1)
            End If
        End If
    Next sp
    '
    Application.ScreenUpdating = fsu
End Sub

Private Sub ListShape2(ce As Range, sp As Shape, ts As String, igptn As String)
    Dim s As String
    ce.Value = ts & sp.name
    ce.Offset(, 1).NumberFormatLocal = "0.0"
    ce.Offset(, 1) = sp.Top
    ce.Offset(, 2).NumberFormatLocal = "0.0"
    ce.Offset(, 2) = sp.Left
    ce.Offset(, 3).NumberFormatLocal = "0.0"
    ce.Offset(, 3) = sp.Height
    ce.Offset(, 4).NumberFormatLocal = "0.0"
    ce.Offset(, 4) = sp.Width
    ce.Offset(, 5).NumberFormatLocal = "0.0"
    ce.Offset(, 5) = sp.Rotation
    s = ""
    If Not sp.Visible Then s = "TRUE"
    ce.Offset(, 6) = s
    If sp.Type <> 1 Then
    ce.Offset(, 7) = shape_typename(sp.Type)
    Else
    ce.Offset(, 7) = shape_shapetypename(sp.AutoShapeType)
    End If
    ce.Offset(, 8) = sp.AlternativeText
    '
    If sp.Type = msoGroup Then
        Dim sp2 As Shape
        For Each sp2 In sp.GroupItems
            If igptn <> "" Then
                If Not re_test(sp2.name, igptn) Then
                    Call ListShape2(ce, sp2, ts & "    ", igptn)
                    Set ce = ce.Offset(1)
                End If
            End If
        Next sp2
    End If
End Sub

'オブジェクトの反映
Public Sub UpdateShape(ra As Range, Optional ws As Worksheet)
    If Not TypeName(Selection) = "Range" Then Exit Sub
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    If ws Is Nothing Then Set ws = ActiveSheet
    If ra Is Nothing Then Set ra = ActiveCell
    
    Dim dic As Dictionary
    Set dic = New Dictionary
    Dim sp As Shape
    For Each sp In ws.Shapes
        dic.Add sp.name, sp
        If sp.Type = msoGroup Then
            Dim sp2 As Shape
            For Each sp2 In sp.GroupItems
                dic.Add sp2.name, sp2
            Next sp2
        End If
    Next sp
    
    Dim ce As Range
    Set ce = ra.Cells(2, 1)
    Set ce = FarLeft(FarTop(ce))
    Dim s As String
    s = Trim(ce.Value)
    Do Until s = ""
        If dic.Exists(s) Then
            Set sp = dic(s)
            If sp.Top <> ce.Offset(, 1) Then sp.Top = ce.Offset(, 1)
            If sp.Left <> ce.Offset(, 2) Then sp.Left = ce.Offset(, 2)
            If sp.Height <> ce.Offset(, 3) Then sp.Height = ce.Offset(, 3)
            If sp.Width <> ce.Offset(, 4) Then sp.Width = ce.Offset(, 4)
            If sp.Rotation <> ce.Offset(, 5) Then sp.Rotation = ce.Offset(, 5)
            If sp.Visible <> (Not ce.Offset(, 6)) Then sp.Visible = Not ce.Offset(, 6)
            If sp.AlternativeText <> ce.Offset(, 8) Then sp.AlternativeText = ce.Offset(, 8)
        End If
        Set ce = ce.Offset(1)
        s = Trim(ce.Value)
    Loop
    '
    Application.ScreenUpdating = fsu
End Sub

Public Sub SetShapeSetting(Optional sr As ShapeRange)
    On Error Resume Next
    If sr Is Nothing Then Set sr = Selection.ShapeRange
    '
    With sr.TextFrame2
        With .TextRange.Font
        End With
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorBottom
        .HorizontalAnchor = msoAnchorNone
    End With
    With sr.TextFrame
        .VerticalOverflow = xlOartVerticalOverflowOverflow
        .HorizontalOverflow = xlOartHorizontalOverflowOverflow
    End With
    With sr.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
        .Solid
        .Visible = msoFalse
    End With
    With sr.line
        .Visible = msoTrue
        .Weight = 1
        .Visible = msoTrue
    End With
    sr.LockAspectRatio = msoTrue
    '
    Dim sp As Shape
    For Each sp In sr.Nodes
        sp.AlternativeText = sp.name
        sp.Placement = xlFreeFloating
    Next sp
End Sub

Public Sub DefaultShapeSetting(Optional sr As ShapeRange)
    Call SetShapeSetting(sr)
    Dim a As Variant
    On Error Resume Next
    Set sr = a.ShapeRange
    On Error GoTo 0
    If sr Is Nothing Then Exit Sub
    sr.SetShapesDefaultProperties
End Sub

'----------------------------------------
'shapetype
'----------------------------------------

Private Function shape_typename(id As Integer) As String
    shape_typename = id
    InitDrawing
    If id < 0 Then id = UBound(ptypename)
    If id <= UBound(ptypename) Then shape_typename = ptypename(id)
End Function

Private Function shape_shapetypename(id As Integer) As String
    shape_shapetypename = id
    InitDrawing
    If id < 0 Then id = UBound(pshapetypename)
    If id <= UBound(pshapetypename) Then shape_shapetypename = pshapetypename(id)
End Function

Private Function ShapeTypeID(s As String) As Integer
    InitDrawing
    Dim i As Integer
    For i = 1 To UBound(pshapetypename)
        If pshapetypename(i) Like s Then
            ShapeTypeID = i
            Exit For
        End If
        If pshapetypename(i) Like s Then
            ShapeTypeID = i
            Exit For
        End If
    Next i
End Function

Private Sub InitDrawing()
    ptype_col = Array("-", _
        "AutoShape", "Callout", "Chart", "Comment", "Freeform", "Group", "EmbeddedOLEObject", "FormControl", "Line", "LinkedOLEObject", "LinkedPicture", _
        "OLEControlObject", "Picture", "Placeholder", "TextEffect", "Media", "TextBox", "ScriptAnchor", "Table", "Canvas", "Diagram", "Ink", "InkComment", _
        "IgxGraphic", "Slicer", "WebVideo", "ContentApp", "Graphic", "LinkedGraphic", "3DModel", "Linked3DModel", "ShapeTypeMixed")
    
    ptypename = Array("-", _
        "オートシェイプ", "吹き出し", "グラフ", "コメント", "フリーフォーム", "Group", "埋め込み OLE オブジェクト", "フォーム コントロール", "Line", _
        "リンク OLE オブジェクト", "リンク画像", "OLE コントロール オブジェクト", "画像", "プレースホルダー", "テキスト効果", "メディア", "テキスト ボックス", _
        "スクリプト アンカー", "テーブル", "キャンバス", "図", "インク", "インク コメント", "SmartArt グラフィック", "Slicer", "Web ビデオ", _
        "コンテンツ Office アドイン", "グラフィック", "リンクされたグラフィック", "3D モデル", "リンクされた 3D モデル", "その他")

    pshapetypename = Array("-", _
        "四角形", "平行四辺形", "台形", "ひし形", "角丸四角形", "八角形", "二等辺三角形", "直角三角形", "楕円", "六角形", "十字形", "五角形", "円柱", "直方体", _
        "四角形：角度付き", "四角形：メモ", "スマイル", "円：塗りつぶしなし", "禁止マーク", "アーチ", "ハート", "稲妻", "太陽", "月", "円弧", "大かっこ", "中かっこ", _
        "ブローチ", "左大かっこ", "右大かっこ", "左中かっこ", "右中かっこ", "矢印：右", "矢印：左", "矢印：上", "矢印：下", "矢印：左右", "矢印：上下", _
        "矢印：四方向", "矢印：三方向", "矢印；折線", "矢印：Uターン", "矢印：二方向", "矢印：上向き折線", "矢印：右カーブ", "矢印：左カーブ", "矢印：上カーブ", _
        "矢印：下カーブ", "矢印：ストライプ", "矢印：V字型", "矢印：五方向", "矢印：山形", "吹き出し：右矢印", "吹き出し：左矢印", "吹き出し：上矢印", _
        "吹き出し：下矢印", "吹き出し：左右矢印", "吹き出し：上下矢印", "吹き出し：四方向矢印", "矢印：環状", "フローチャート：処理", "フローチャート：代替処理", _
        "フローチャート：判断", "フローチャート：データ", "フローチャート：定義済み処理", "フローチャート：内部記憶", "フローチャート：書類", _
        "フローチャート：複数書類", "フローチャート：端子", "フローチャート：準備", "フローチャート：手操作入力", "フローチャート：手作業", _
        "フローチャート：結合子", "フローチャート：他ページ結合子", "フローチャート：カード", "フローチャート：せん孔テープ", "フローチャート：和接合", _
        "フローチャート：論理和", "フローチャート：照合", "フローチャート：分類", "フローチャート：抜き出し", "フローチャート：組み合わせ", _
        "フローチャート：記憶データ", "フローチャート：論理積ゲート", "フローチャート：順次アクセス記憶", "フローチャート：磁気ディスク", _
        "フローチャート：直接アクセス記憶", "フローチャート：表示", "爆発 8pt", "爆発 14pt", "星 4pt", "星 5pt", "星 8pt", "星 16pt", "星 24pt", "星 32pt", _
        "リボン：上に曲がる", "リボン：下に曲がる", "リボン：カーブして上に曲がる", "リボン：カーブして下に曲がる", "スクロール：縦", "スクロール：横", "波線", _
        "小波", "吹き出し：四角形", "吹き出し：角丸四角形", "吹き出し：円形", "思考吹き出し：雲形", "吹き出し：線", "吹き出し：線", "吹き出し：折線", _
        "吹き出し：２つの折線", "吹き出し：線(強調線付き)", "吹き出し：線(強調線付き)", "吹き出し：折線(強調線付き)", "吹き出し：２つの折線(強調線付き)", _
        "吹き出し：線(枠なし)", "吹き出し：線(枠なし)", "吹き出し：折線(枠なし)", "吹き出し：２つの折線(枠なし)", "吹き出し：線(枠付き、強調線付き)", _
        "吹き出し：線(枠付き、強調線付き)", "吹き出し：折線(枠付き、強調線付き)", "吹き出し：２つの折線(枠付き、強調線付き)", _
        "ボタン", "[ホーム] ボタン", "[ヘルプ] ボタン", "[情報] ボタン", "[戻る] または [前へ] ボタン", "[進む] または [次へ] ボタン", "[開始] ボタン", _
        "[終了] ボタン", "[戻る] ボタン", "[文書] ボタン", "[サウンド] ボタン", "[ビデオ] ボタン", "吹き出し", "未サポート", "フローチャート：オフライン記憶", _
        "リボン：両端矢印", "斜め縞", "部分円", "台形：非対称", "十角形", "七角形", "十二角形", "星 6pt", "星 7pt", "星 10pt", "星 12pt", "四角形：1つを丸める", _
        "四角形：上の2つを丸める", "四角形：1つを切り取る1つを丸める", "四角形：1つを切り取る", "四角形：上の2つを切り取る", "四角形：対角を丸める", _
        "四角形：対角を切り取る", "フレーム", "フレーム(半分)", "涙型", "弦", "L字", "加算記号", "減算記号", "乗算記号", "除算記号", "次の値と等しい", "等号否定", _
        "四隅：三角形", "四隅：四角形", "四隅：四分円", "ギア 6pt", "ギア 9pt", "漏斗", "四分円", "矢印：反時計回り", "矢印：両方向時計回り", "矢印：曲線", "雲", _
        "正方形：対角", "正方形：水平・対角", "正方形：水平・垂直", "斜線", "その他")
End Sub

'図面シートを追加
Public Sub AddDrawingSheet(Optional sc As Integer = 25)
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim ws As Worksheet
    Set ws = ActiveSheet
    '
    ws.Cells.RowHeight = 20.3
    ws.Cells.ColumnWidth = 8#
    '
    Application.ScreenUpdating = fsu
    Exit Sub
    '
    If sc <= 680 Then
        ws.Cells.RowHeight = 0.6 * sc
    End If
    If sc <= 2560 Then
        Dim w As Double
        w = (sc - 7) / 10
        If sc < 17 Then w = 0.059 * sc
        ws.Cells.ColumnWidth = w
    End If
     '
    Application.ScreenUpdating = fsu
End Sub


