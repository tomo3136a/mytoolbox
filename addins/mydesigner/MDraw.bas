Attribute VB_Name = "MDraw"
'==================================
'描画
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'データ

'図形リストメンバー
'[0] 名称
'[1] 表示名
'[2] 0:文字列,1:整数,2:0.0##,3:論理値,4:色
Private Const c_ShapeInfoMember As String = "" _
        & ";Name,名前,0     ;Title,タイトル,0" _
        & ";" _
        & ";ID,ID,0         ;Type,種別,0     ;Style,スタイル,0" _
        & ";Top,上位置,2    ;Left,左位置,2   ;Back,後位置,2     ;Rotation,回転,2" _
        & ";Height,高さ,2   ;Width,幅,2      ;Depth,奥行き,2" _
        & ";" _
        & ";Visible,表示,3                  ;Transparency,透明度,2" _
        & ";LineVisible,枠線表示,3          ;LineColor,枠線色,4" _
        & ";FillVisible,塗りつぶし表示,3    ;FillColor,塗りつぶし色,4" _
        & ";" _
        & ";Text,テキスト,0     ;AltText,代替えテキスト,0" _
        & ";Scale,スケール,2    ;X0,原点X,2     ;Y0,原点Y,2     ;Z0,原点Z,2"

'図形リストヘッダ
Private Const c_ShapeInfoHeader As String = "" _
    & ";名称,           Name,Title" _
    & ";形状,           Name,ID,Type,Style,Title" _
    & ";位置,           Name,Top,Left,Back,Rotation" _
    & ";サイズ,         Name,Height,Width,Depth" _
    & ";表示,           Name,Visible,Transparency" _
    & ";枠線,           Name,LineVisible,LineColor" _
    & ";塗り,           Name,FillVisible,FillColor" _
    & ";テキスト,       Name,Text" _
    & ";代替えテキスト, Name,AltText" _
    & ";属性,           Name,Scale,X0,Y0,Z0"

'環境パラメータ項目
Public Enum E_DrawParam
    E_IGNORE = 1
    E_SCALE = 2
    E_AXES = 3
    E_FLAG = 4
    E_PART = 10
End Enum

'環境パラメータ
Private g_mask As String            '無効検索文字
Private g_scale As Double           'スケール
Private g_axes As Double            '軸間隔
Private g_flag As Integer           'モード(0:,1:,2:,3:)
Private g_part As String            '部品名

Private ptype_col As Variant        '図形タイプテーブル
Private ptypename As Variant        '図形タイプ名称テーブル
Private pshapetypename As Variant   '図形タイプテーブル

'----------------------------------------
'パラメータ制御
'----------------------------------------

'描画パラメータ初期化
Public Sub ResetDrawParam(Optional id As Integer)
    If id = 0 Or id = 1 Then g_mask = ""
    If id = 0 Or id = 2 Then g_scale = 0.1
    If id = 0 Or id = 3 Then g_axes = 10
    If id = 0 Or id = 4 Then g_flag = 0
    If id = 0 Or id = 10 Then g_part = ""
End Sub

'描画パラメータ設定
Public Sub SetDrawParam(id As Integer, ByVal val As String)
    Select Case id
    Case 1: g_mask = val
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
    Case 4: g_flag = (g_flag And (65535 - 1)) Or (val * 1)
    Case 5: g_flag = (g_flag And (65535 - 2)) Or (val * 2)
    Case 6: g_flag = (g_flag And (65535 - 4)) Or (val * 4)
    Case 7: g_flag = (g_flag And (65535 - 8)) Or (val * 8)
    Case 8: g_flag = (g_flag And (65535 - 16)) Or (val * 16)
    Case 9: g_flag = (g_flag And (65535 - 32)) Or (val * 32)
    Case 10: g_part = val
    End Select
End Sub

'描画パラメータ取得
Public Function GetDrawParam(id As Integer) As String
    Select Case id
    Case 1: GetDrawParam = g_mask
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

'描画パラメータフラグチェック
Public Function IsDrawParam(id As Integer) As Boolean
    IsDrawParam = ((g_flag And (65535 - (2 ^ (id - 4)))) <> 0)
End Function

'----------------------------------------
'図形属性制御
'----------------------------------------

'図形属性取得
Function GetShapeProperty(sr As ShapeRange, k As String) As String
    GetShapeProperty = ParamStrVal(sr.AlternativeText, k)
End Function

'図形属性設定
Sub SetShapeProperty(sr As ShapeRange, k As String, v As String)
    sr.AlternativeText = UpdateParamStr(sr.AlternativeText, k, v)
End Sub

'----------------------------------------
'図形基本設定
'----------------------------------------

'図形基本設定
'[3] テキスト, 塗りつぶしと線
Public Sub SetShapeStyle(Optional ByVal sr As ShapeRange)
    If sr Is Nothing Then
        If TypeName(Selection) = "Range" Then Exit Sub
        Set sr = Selection.ShapeRange
    End If
    Call SetShapeSetting(sr, 3)
End Sub

'標準図形設定
Public Sub DefaultShapeSetting(Optional ByVal sr As ShapeRange)
    If sr Is Nothing Then
        If TypeName(Selection) = "Range" Then Exit Sub
        Set sr = Selection.ShapeRange
    End If
    Call SetShapeSetting(sr, 511)
End Sub

Public Sub SetDefaultShapeStyle()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeOval, 10, 10, 10, 10)
    Call SetShapeSetting(ws.Shapes.Range(sh.name), 1 + 2 + 4 + 256)
    sh.Delete
    Set sh = ws.Shapes.AddLine(10, 10, 20, 20)
    Call SetShapeSetting(ws.Shapes.Range(sh.name), 1 + 2 + 4 + 256)
    sh.Delete
    Set sh = ws.Shapes.AddTextbox(msoTextOrientationDownward, 10, 10, 10, 10)
    Call SetShapeSetting(ws.Shapes.Range(sh.name), 1 + 2 + 4 + 256)
    sh.Delete
End Sub

'----------------------------------------
'図形基本設定
'[1] テキスト
'[2] 塗りつぶしと線
'[4] サイズとプロパティ
'[8] 代替え文字
'[256] デフォルト設定
Private Sub SetShapeSetting(Optional ByVal sr As ShapeRange, Optional mode As Integer = 255)
    
    Dim sh As Shape
    On Error Resume Next
    
    '設定(テキスト)
    If mode And 1 Then
        With sr.TextFrame2
            'With .TextRange.Font
            'End With
            .MarginLeft = 1
            .MarginRight = 1
            .MarginTop = 1
            .MarginBottom = 1
            .AutoSize = msoAutoSizeNone
            .WordWrap = msoFalse
            .VerticalAnchor = msoAnchorBottom
            .HorizontalAnchor = msoAnchorNone
        End With
        With sr.TextFrame
            .VerticalOverflow = xlOartVerticalOverflowOverflow
            .HorizontalOverflow = xlOartHorizontalOverflowOverflow
        End With
    End If
    
    '設定(塗りつぶしと線)
    If mode And 2 Then
        With sr.Fill
            .Visible = msoTrue
            '.ForeColor.RGB = RGB(255, 0, 0)
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
            '.Visible = msoFalse
        End With
        With sr.line
            .Visible = msoTrue
            .Weight = 1
            .Visible = msoTrue
        End With
    End If
    
    '設定(サイズとプロパティ)
    If mode And 4 Then
        sr.LockAspectRatio = msoTrue
        sr.Placement = xlMove
        For Each sh In sr
            sh.Placement = xlMove
        Next sh
    End If
    '
    '設定(代替え文字)
    If mode And 8 Then
        For Each sh In sr
            sh.AlternativeText = sh.name
        Next sh
    End If
    
    'デフォルト設定
    If mode And 256 Then sr.SetShapesDefaultProperties

    On Error GoTo 0

End Sub

'表示/非表示反転
Public Sub ToggleVisible(mode As Integer, Optional sr As ShapeRange)
    
    Dim sr2 As ShapeRange
    Set sr2 = sr
    If sr2 Is Nothing Then
        If TypeName(Selection) = "Range" Then Exit Sub
        Set sr2 = Selection.ShapeRange
    End If
    
    Select Case mode
    Case 0
        '表示/非表示反転
        With sr2.Fill
            If .Visible = msoTrue Then
                .Visible = msoFalse
            Else
                .Visible = msoTrue
            End If
        End With
    Case 3
        '3D表示/非表示反転
        With sr2.ThreeD
            If .Visible = msoTrue Then
                .Visible = msoFalse
            Else
                .Visible = msoTrue
                .SetPresetCamera (msoCameraIsometricTopUp)
                .RotationX = 45.2809
                .RotationY = -35.3962666667
                .RotationZ = -60.1624166667
            End If
        End With
    End Select

End Sub

'図形名更新
Public Sub UpdateShapeName(ws As Worksheet)
    
    Dim re As Object
    Set re = regex("\s+\d*$")
    
    Dim sh As Shape
    For Each sh In ws.Shapes
        Dim s As String
        s = re.Replace(sh.name, " " & sh.id)
        If s <> sh.name Then sh.name = s
        If sh.Type = msoGroup Then
            Dim sh2 As Shape
            For Each sh2 In sh.GroupItems
                s = re.Replace(sh2.name, " " & sh2.id)
                If s <> sh2.name Then sh2.name = s
            Next sh2
        End If
    Next sh

End Sub

'----------------------------------------
'図形制御
'----------------------------------------

'図形を削除
Public Sub RemoveSharps(Optional ByVal ws As Worksheet)
    
    '対象選択
    If ws Is Nothing Then
        If TypeName(Selection) <> "Range" Then
            Selection.Delete
            Exit Sub
        End If
        If MsgBox("全図形を削除しますか？", vbYesNo) <> vbYes Then Exit Sub
        Set ws = ActiveSheet
    End If
    
    '画面更新停止
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    
    '画面更新再開
    Application.ScreenUpdating = True
    On Error GoTo 0

End Sub

'図形を絵に変換
Public Sub ConvertToPicture()
    
    '対象選択
    If TypeName(Selection) = "Range" Then Exit Sub
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    If sr Is Nothing Then Exit Sub
    
    '画面更新停止
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Dim s As String
    Dim x As Double
    Dim y As Double
    s = sr.name
    x = sr.Left
    y = sr.Top
    Selection.Cut
    Dim ws As Worksheet
    Set ws = Selection.Worksheet
    ws.PasteSpecial 0
    Selection.name = s
    Selection.Left = x
    Selection.Top = y
    Application.CutCopyMode = 0
    
    '画面更新再開
    Application.ScreenUpdating = fsu

End Sub

'----------------------------------------
'図形部品描画
'----------------------------------------

'部品描画
Public Function DrawParts(ws As Worksheet, x0 As Double, y0 As Double, s As String) As String
    Dim cs As Worksheet
    Set cs = GetSheet("#shapes")
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
    Set sh = ws.Shapes.Range(ColToArr(ns)).Group
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
    Set sh = ws.Shapes.Range(ColToArr(ns)).Group
    sh.line.Transparency = 0.5
    sh.name = "方眼紙 " & sh.id
    DrawGraph2 = sh.name
End Function

'----------------------------------------
'図形属性制御
'----------------------------------------

'図形情報リストアップ
Public Sub ListShapeInfo(ByVal ws As Worksheet, Optional mode As Integer)
    
    '出力先セル/図形リスト取得
    Dim ce As Range
    Dim sr As Variant
    Dim sr2 As ShapeRange
    If TypeName(Selection) = "Range" Then
        Set ce = FarLeftTop(Selection)
        If ce.Row > 1 Then
            Dim s As String
            s = ce.Offset(-1).Value
            If s <> "" Then
                Dim ss() As String
                ss = Split(Replace(s, "]", ""), "[", 2)
                If UBound(ss) = 0 Then
                    Set ws = ActiveWorkbook.Sheets(ss(0))
                Else
                    Set ws = Workbooks(ss(1)).Sheets(ss(0))
                End If
            End If
        End If
        Set sr = ws.Shapes
    Else
        Set ce = GetCell("リスト出力位置を指定してください", "図形リスト出力")
        Set sr = Selection.ShapeRange
    End If
    If ce Is Nothing Or sr Is Nothing Then Exit Sub
    If ce.Value = "" Then
        ce.Value = sr.Parent.name & "[" & sr.Parent.Parent.name & "]"
        Set ce = ce.Offset(1)
    End If
    
    'テーブル項目取得
    Dim dic As Dictionary
    ArrStrToDict dic, c_ShapeInfoMember, 1
    
    'ヘッダ取得
    Dim hdr() As String
    StringToRow hdr, c_ShapeInfoHeader, mode
    If mode < 1 And ce.Value <> "" Then
        hdr = GetHeaderArray(ce, dic)
    End If
    
    'データ配列作成
    Dim rcnt As Long
    rcnt = 1 + sr.Count
    Dim sh As Shape
    For Each sh In sr
        If sh.Type = msoGroup Then rcnt = rcnt + sh.GroupItems.Count
    Next sh
    Dim arr As Variant
    ReDim arr(rcnt, UBound(hdr))
    
    'ヘッダ行設定
    Dim r As Long
    Dim c As Long
    'Dim s As String
    For c = 0 To UBound(hdr)
        s = UCase(hdr(c))
        If dic.Exists(s) Then
            arr(r, c) = dic(s)(1)
        Else
            arr(r, c) = hdr(c)
        End If
    Next c
    r = r + 1
    
    '図形名マスク設定
    Dim ptn As String
    Dim flg As Boolean
    ptn = g_mask
    flg = True
    If Left(ptn, 1) = "!" Then
        ptn = Mid(ptn, 2)
        flg = False
    End If
    If ptn = "" Then ptn = ".*"
    
    'レコード作成
    For Each sh In sr
        AddShapeRecord arr, r, sh, hdr, ptn, flg
    Next sh
    rcnt = r
    
    '画面更新停止
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    'テーブルデータクリア
    TableDataRange(ce).Clear
    
    '表示形式設定
    Dim ra As Range
    For c = 0 To UBound(hdr)
        Set ra = ce.Parent.Range(ce.Cells(2, c + 1), ce.Cells(rcnt, c + 1))
        s = UCase(hdr(c))
        If dic.Exists(s) Then
            Select Case CInt(dic(s)(2))
            Case 1: ra.NumberFormatLocal = "0"
            Case 2: ra.NumberFormatLocal = "0.0##"
            Case Else: ra.NumberFormatLocal = "@"
            End Select
        End If
    Next c
    
    'レコード書き込み
    ce.Resize(1 + rcnt, 1 + UBound(hdr)).Value = arr
    
    '配色
    For c = 0 To UBound(hdr)
        Set ra = ce.Parent.Range(ce.Cells(2, c + 1), ce.Cells(rcnt, c + 1))
        s = UCase(hdr(c))
        If dic.Exists(s) Then
            Select Case CInt(dic(s)(2))
            Case 4  '色
                For r = 2 To rcnt
                    s = ce.Cells(r, c + 1)
                    If s <> "" Then
                        ce.Cells(r, c + 1).Interior.color = val("&H" & s)
                    Else
                        ce.Cells(r, c + 1).ClearFormats
                    End If
                Next r
            End Select
        End If
    Next c
    
    'テーブルサイズ調整
    HeaderAutoFit ce
    
    '画面更新再開
    Application.ScreenUpdating = fsu

End Sub

'文字列から行配列を取得
Sub StringToRow(arr() As String, info As String, Optional mode As Integer = 0)
    
    Dim dic As Dictionary
    ArrStrToDict dic, info
    
    Dim kw As String
    kw = dic.Keys(mode)
    arr = dic(kw)
    arr = TakeArray(arr, 1)

End Sub

'図形レコードを配列に追加

Private Sub AddShapeRecord(arr As Variant, r As Long, sh As Shape, hdr As Variant, ptn As String, flg As Boolean)
        
    Dim c As Long
    Dim s As String
    If sh.Type = msoGroup Then
        For c = 0 To UBound(hdr)
            s = hdr(c)
            arr(r, c) = ShapeValue(sh, s, "")
        Next c
        r = r + 1
        Dim cnt As Long
        cnt = 0
        Dim v As Variant
        For Each v In sh.GroupItems
            Dim sh2 As Shape
            Set sh2 = v
            If re_test(sh2.name, ptn) = flg Then
                For c = 0 To UBound(hdr)
                    s = hdr(c)
                    arr(r, c) = ShapeValue(sh2, s, "  ")
                Next c
                r = r + 1
                cnt = cnt + 1
            End If
        Next v
        If cnt = 0 Then r = r - 1
    ElseIf re_test(sh.name, ptn) = flg Then
        For c = 0 To UBound(hdr)
            s = hdr(c)
            arr(r, c) = ShapeValue(sh, s, "")
        Next c
        r = r + 1
    End If

End Sub

'図形情報取得
Private Function ShapeValue(sh As Shape, k As String, Optional ts As String) As Variant
    
    Dim v As Variant
    v = "-"
    On Error Resume Next
    Select Case UCase(k)
    
    Case "NAME": v = ts & sh.name
    Case "TITLE": v = sh.Title
    Case "ID": v = sh.id
    Case "TYPE": v = shape_typename(sh.Type)
        If sh.Type = 1 Then v = shape_shapetypename(sh.AutoShapeType)
    Case "STYLE": v = sh.ShapeStyle
    
    Case "TOP": v = sh.Top
    Case "LEFT": v = sh.Left
    Case "BACK": v = sh.ThreeD.z
    Case "ROTATION": v = sh.Rotation
    
    Case "HEIGHT": v = sh.Height
    Case "WIDTH": v = sh.Width
    Case "DEPTH": v = sh.ThreeD.Depth
    
    Case "VISIBLE": v = CBool(sh.Visible)
    
    Case "LINEVISIBLE": v = CBool(sh.line.Visible)
    Case "LINECOLOR": v = Right("000000" & Hex(sh.line.ForeColor), 6)
    
    Case "FILLVISIBLE": v = CBool(sh.Fill.Visible)
    Case "FILLCOLOR": v = Right("000000" & Hex(sh.Fill.ForeColor), 6)
    Case "TRANSPARENCY": v = sh.Fill.Transparency
    
    Case "TEXT": v = sh.TextFrame2.TextRange.text
    Case "ALTTEXT": v = sh.AlternativeText
    
    Case "SCALE": v = Replace(re_match(sh.AlternativeText, "g:[+-]?[\d.]+"), "g:", "")
    Case "X0": v = Replace(re_match(sh.AlternativeText, "p:[+-]?[\d.]+,[+-]?[\d.]+"), "p:", "")
    Case "Y0": v = Replace(re_match(sh.AlternativeText, "d:[+-]?[\d.]+,[+-]?[\d.]+"), "d:", "")
    
    End Select
    On Error GoTo 0
    ShapeValue = v

End Function

'図形情報設定
Private Sub UpdateShapeValue(sh As Shape, k As String, ByVal v As Variant)
    
    'On Error Resume Next
    Select Case UCase(k)
    
    Case "NAME": sh.name = CStr(v)
    Case "TITLE": sh.Title = CStr(v)
    Case "ID":
    Case "TYPE":
    Case "STYLE":
    
    Case "TOP": sh.Top = CDbl(v)
    Case "LEFT": sh.Left = CDbl(v)
    Case "BACK": sh.ThreeD.z = CDbl(v)
    Case "ROTATION": sh.Rotation = CDbl(v)
    
    Case "HEIGHT": sh.Height = CDbl(v)
    Case "WIDTH": sh.Width = CDbl(v)
    Case "DEPTH": sh.ThreeD.Depth = CDbl(v)
    
    Case "VISIBLE": sh.Visible = CBool(v)
    
    Case "LINEVISIBLE": sh.line.Visible = CBool(v)
    Case "LINECOLOR": sh.line.ForeColor = v
    
    Case "FILLVISIBLE": sh.Fill.Visible = CBool(v)
    Case "FILLCOLOR": sh.Fill.ForeColor = v
    Case "TRANSPARENCY": sh.Fill.Transparency = v
    
    Case "TEXT": sh.TextFrame2.TextRange.text = v
    Case "ALTTEXT": sh.AlternativeText = v
    
    Case "SCALE": sh.AlternativeText = UpdateParamStr(sh.AlternativeText, "sc", CStr(v))
    Case "X0": sh.AlternativeText = UpdateParamStr(sh.AlternativeText, "x0", CStr(v))
    Case "Y0": sh.AlternativeText = UpdateParamStr(sh.AlternativeText, "y0", CStr(v))
    
    End Select
    'On Error GoTo 0

End Sub

'----------------------------------------

'図形情報リストの反映
Public Sub UpdateShapeInfo(ByVal ra As Range, Optional ByVal ws As Worksheet)
    
    If Not TypeName(Selection) = "Range" Then Exit Sub
    If ws Is Nothing Then Set ws = ActiveSheet
    If ra Is Nothing Then Set ra = ActiveCell
    
    'テーブル開始位置を取得
    Dim ce As Range
    Set ce = FarLeftTop(ra.Cells(2, 1))
    
    'テーブル項目取得
    Dim hdr_dic As Dictionary
    ArrStrToDict hdr_dic, c_ShapeInfoMember, 1
    
    'ヘッダ取得
    Dim hdr() As String
    hdr = GetHeaderArray(ce, hdr_dic)
    
    '図形リスト作成
    Dim dic As Dictionary
    Set dic = New Dictionary
    Dim sh As Shape, sh2 As Shape
    For Each sh In ws.Shapes
        dic.Add sh.name, sh
        If sh.Type = msoGroup Then
            For Each sh2 In sh.GroupItems
                dic.Add sh2.name, sh2
            Next sh2
        End If
    Next sh
    
    '画面更新停止
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Dim s As String
    s = Trim(ce.Value)
    Do Until s = ""
        If dic.Exists(s) Then
            Set sh = ws.Shapes(s)
            Dim f As Boolean
            f = sh.ThreeD.Visible
            If f Then sh.ThreeD.Visible = msoFalse
            Dim c As Integer
            For c = 1 To UBound(hdr)
                If ShapeValue(sh, hdr(c)) <> ce.Offset(, c) Then
                    UpdateShapeValue sh, hdr(c), ce.Offset(, c)
                End If
            Next c
            'If sp.Top <> ce.Offset(, 1) Then sp.Top = ce.Offset(, 1)
            'If sp.Left <> ce.Offset(, 2) Then sp.Left = ce.Offset(, 2)
            'If sp.Height <> ce.Offset(, 3) Then sp.Height = ce.Offset(, 3)
            'If sp.Width <> ce.Offset(, 4) Then sp.Width = ce.Offset(, 4)
            'If sp.Rotation <> ce.Offset(, 5) Then sp.Rotation = ce.Offset(, 5)
            'If sp.Visible <> (Not ce.Offset(, 6)) Then sp.Visible = Not ce.Offset(, 6)
            'If sp.AlternativeText <> ce.Offset(, 8) Then sp.AlternativeText = ce.Offset(, 8)
        End If
        If f Then sh.ThreeD.Visible = msoTrue
        Set ce = ce.Offset(1)
        s = Trim(ce.Value)
    Loop
    '
    '画面更新再開
    Application.ScreenUpdating = fsu

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


