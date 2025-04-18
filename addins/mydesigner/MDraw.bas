Attribute VB_Name = "MDraw"
'==================================
'描画
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'図形属性ID

Private Enum E_SAID
    N_NAME
    N_TITLE
    N_ID
    N_TYPE
    N_STYLE
    
    N_TOP
    N_LEFT
    N_BACK
    N_ROTATION
    
    N_HEIGHT
    N_WIDTH
    N_DEPTH
    
    N_VISIBLE
    N_LINEVISIBLE
    N_LINECOLOR
    
    N_FILLVISIBLE
    N_FILLCOLOR
    N_TRANSPARENCY
    
    N_TEXT
    N_ALTTEXT
    
    N_SCALE
    N_X0
    N_Y0
    N_Z0
    N_DX
    N_DY
    N_DZ
End Enum

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
        & ";Scale,スケール,2    ;X0,原点X,2     ;Y0,原点Y,2     ;Z0,原点Z,2" _
        & ";DX,サイズX,2        ;DY,サイズY,2   ;DZ,サイズZ,2"

'図形リストヘッダ
Private Const c_ShapeInfoHeader As String = "" _
    & ";名称,           Name" _
    & ";形状,           Name,ID,Type,Style,Title" _
    & ";位置,           Name,Left,Top,Back,Rotation" _
    & ";サイズ,         Name,Width,Height,Depth" _
    & ";表示,           Name,Visible,Transparency" _
    & ";枠線,           Name,LineVisible,LineColor" _
    & ";塗り,           Name,FillVisible,FillColor" _
    & ";テキスト,       Name,Text" _
    & ";代替えテキスト, Name,AltText" _
    & ";属性,           Name,Scale,X0,Y0,Z0,DX,DY,DZ"

'環境パラメータ項目
Public Enum E_DrawParam
    E_IGNORE = 1
    E_SCALE = 2
    E_AXES = 3
    E_FLAG = 4
    E_PART = 10
End Enum

'環境パラメータ
Private g_mask As String            '検索文字
Private g_scale As Double           'スケール
Private g_axes As Double            '軸間隔
Private g_flag As Long              'モード
                                    ' 4.A面, 5.B面
                                    ' 6.配置制約, 7.配線制約
                                    ' 8.PTH, 9:Note
Private g_part As String            '部品名

Private ptype_col As Variant        '図形タイプテーブル
Private ptypename As Variant        '図形タイプ名称テーブル
Private pshapetypename As Variant   '図形タイプテーブル

'----------------------------------------
'パラメータ制御
'----------------------------------------

'描画パラメータ初期化
Public Sub Draw_ResetParam(Optional id As Integer)
    If id = 0 Or id = 1 Then g_mask = ""
    If id = 0 Or id = 2 Then g_scale = 1
    If id = 0 Or id = 3 Then g_axes = 10
    If id = 0 Or id = 4 Then g_part = ""
    If id = 0 Or id = 5 Then g_flag = 1 + 2
End Sub

'描画パラメータ設定
Public Sub Draw_SetParam(id As Integer, ByVal val As String)
    Select Case id
    Case 2
        If val <= 0 Then
            MsgBox "比率の設定が間違っています。(設定値>0)" & Chr(10) _
                & "設定値： " & val
            Exit Sub
        End If
    Case 3
        If val <= 0 Then
            MsgBox "目盛りの設定が間違っています。(設定値>0)" & Chr(10) _
                & "設定値： " & val
            Exit Sub
        End If
    End Select
    
    Select Case id
    Case 1: g_mask = val
    Case 2: g_scale = val
    Case 3: g_axes = val
    Case 4: g_part = val
    End Select
End Sub

'描画パラメータ取得
Public Function Draw_GetParam(id As Integer) As String
    Dim msg As String
    Select Case id
    Case 2
        If g_scale <= 0 Then
            Draw_ResetParam id
            msg = "比率の設定を初期化しました。(設定値" & g_scale & ")"
        End If
    Case 3
        If g_axes <= 0 Then
            Draw_ResetParam id
            msg = "目盛りの設定を初期化しました。(設定値" & g_axes & ")"
        End If
    End Select
    If msg <> "" Then MsgBox msg
    
    Select Case id
    Case 1: Draw_GetParam = g_mask
    Case 2: Draw_GetParam = g_scale
    Case 3: Draw_GetParam = g_axes
    Case 4: Draw_GetParam = g_part
    End Select
End Function

'描画パラメータフラグ設定
Public Sub Draw_SetParamFlag(id As Integer, Optional ByVal val As Boolean = True)
    g_flag = g_flag And Not 2 ^ (id Mod 24)
    If val Then g_flag = g_flag Or 2 ^ (id Mod 24)
End Sub

'描画パラメータフラグチェック
Public Function Draw_IsParamFlag(id As Integer) As Boolean
    Draw_IsParamFlag = Not ((g_flag And 2 ^ (id Mod 24)) = 0)
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
'色操作
'----------------------------------------

Private Function FormatRGB(v As ColorFormat) As String
    Dim s As String
    s = v.Brightness
    s = v.SchemeColor
    s = v.Type
    s = Right("00000000" & Hex(v), 8)
    s = "#" & Mid(s, 1, 2) & Mid(s, 7, 2) & Mid(s, 5, 2) & Mid(s, 3, 2)
    FormatRGB = s & " " & v.Type & " " & v.SchemeColor & " " & v.Brightness
End Function

Private Function ToRGB(v As Variant) As Long
    Dim s As String
    s = Split(v, " ")(0)
    s = Right("00000000" & Replace(Replace(s, "#", ""), "&H", ""), 6)
    
    ToRGB = RGB(CLng("&H" & Mid(s, 1, 2)), CLng("&H" & Mid(s, 3, 2)), CLng("&H" & Mid(s, 5, 2)))
End Function

'塗りつぶし色表示取得
Private Function FormatInterior(v As Interior) As String
    If v.ThemeColor < 0 Then Exit Function
    Dim s As String
    s = Right("00000000" & Hex(v.Color), 6)
    s = "#" & Mid(s, 5, 2) & Mid(s, 3, 2) & Mid(s, 1, 2)
    If v.ThemeColor > 0 Then s = s & " " & v.ThemeColor & Format(v.TintAndShade, "+0%;-0%;"""";@")
    FormatInterior = s
End Function

'塗りつぶし色からフォント色設定
Private Sub SetFontColorFromInterior(ra As Range)
    
    If ra.Interior.ColorIndex <= 0 Then Exit Sub
    Dim s As String
    s = Right("00000000" & Hex(ra.Interior.Color), 6)
    Dim r As Long, g As Long, b As Long, a As Long
    r = CLng("&H0" & Mid(s, 5, 2))
    g = CLng("&H0" & Mid(s, 3, 2))
    b = CLng("&H0" & Mid(s, 1, 2))
    a = r * 0.3 + g * 0.6 + b * 0.1
    ra.Font.ColorIndex = IIf(a > 100, 1, 2)

End Sub

'色表示取得
Private Function FormatColor(v As ColorFormat) As String
    Dim s As String
    s = Right("00000000" & Hex(v), 6)
    s = "#" & Mid(s, 5, 2) & Mid(s, 3, 2) & Mid(s, 1, 2)
    If v.ObjectThemeColor > 0 Then s = s & " " & v.ObjectThemeColor & Format(v.Brightness, "+0%;-0%;"""";@")
    FormatColor = s
End Function

'----------------------------------------
'図形基本設定
'----------------------------------------

'テキストボックス基本設定
Public Sub SetTextBoxStyle()
    If TypeName(Selection) = "Range" Then Exit Sub
    Call SetShapeSetting(Selection.ShapeRange, &H33)
End Sub

'標準図形設定
Public Sub DefaultShapeSetting()
    If TypeName(Selection) = "Range" Then Exit Sub
    Call SetShapeSetting(Selection.ShapeRange, &H37)
End Sub

'図形基本設定
'[3] テキスト, 塗りつぶしと線
Public Sub SetDefaultShapeStyle()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    '
    With ws.Shapes.AddShape(msoShapeOval, 10, 10, 10, 10)
        SetShapeSetting ws.Shapes.Range(.name), &H107
        .Delete
    End With
    '
    With ws.Shapes.AddLine(10, 10, 20, 20)
        SetShapeSetting ws.Shapes.Range(.name), &H107
        .Delete
    End With
    '
    With ws.Shapes.AddTextbox(msoTextOrientationDownward, 10, 10, 10, 10)
        SetShapeSetting ws.Shapes.Range(.name), &H137
        .Delete
    End With
End Sub

'----------------------------------------
'図形基本設定
'[1] 設定(テキスト)
'[2] 設定(塗りつぶしと線)
'[4] サイズとプロパティ
'[8] 代替え文字
'[16]
'[32]
'[256] デフォルト設定
Private Sub SetShapeSetting( _
        Optional ByVal sr As ShapeRange, _
        Optional mode As Integer = 255 _
    )
    
    Dim sh As Shape
    On Error Resume Next
    
    '設定(テキスト)
    If (mode And 1) Then
        With sr.TextFrame2
            With .TextRange.Font.Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0
                .Transparency = 0
                .Solid
            End With
            .MarginLeft = 1
            .MarginRight = 1
            .MarginTop = 1
            .MarginBottom = 1
            .AutoSize = msoAutoSizeNone
            .WordWrap = msoFalse
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
        With sr.TextFrame
            .VerticalOverflow = xlOartVerticalOverflowOverflow
            .HorizontalOverflow = xlOartHorizontalOverflowOverflow
        End With
    End If
    
    '設定(塗りつぶし)
    If (mode And 2) Then
        With sr.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
    End If
    
    '設定(線)
    If (mode And 2) Then
        With sr.line
            .Visible = msoTrue
            .Weight = 1
            .Visible = msoTrue
        End With
    End If
    
    '設定(線)
    If (mode And &H10) Then
        With sr.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
            .Visible = msoFalse
        End With
        With sr.line
            .Visible = msoTrue
            .Weight = 1
            .Visible = msoTrue
        End With
    End If
    
    '設定(サイズとプロパティ)
    If (mode And 4) Then
        sr.LockAspectRatio = msoTrue
        sr.Placement = xlMove
        For Each sh In sr
            sh.Placement = xlMove
        Next sh
    End If
    
    'デフォルトに設定
    If (mode And &H20) Then
        sr.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        sr.TextFrame2.Orientation = msoTextOrientationHorizontal
    End If
    
    '設定(代替え文字)
    If (mode And &H8) Then
        For Each sh In sr
            sh.AlternativeText = sh.name
        Next sh
    End If
    
    'デフォルト設定
    If (mode And &H100) Then sr.SetShapesDefaultProperties

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
            If .PresetCamera <> msoPresetCameraMixed Then
                .ResetRotation
                .RotationX = 0
                .RotationY = 0
                .RotationZ = 0
                .Visible = False
                .Visible = True
            Else
                .SetPresetCamera (msoCameraIsometricTopUp)
                .RotationX = 45.2809
                .RotationY = -35.3962666667
                .RotationZ = -60.1624166667
            End If
        End With
    End Select

End Sub

'図形名更新
Public Sub UpdateShapeName(v As Variant, Optional bid As Boolean = True)
    
    Dim re As Object
    Set re = regex("\s+\d*$")
    
    Dim sr As Variant
    Select Case TypeName(v)
    Case "ShapeRange": Set sr = v
    Case "Worksheet": Set sr = v.Shapes
    Case "Shape": Set sr = v.Parent.Shapes.Range(v.name)
    Case Else: Exit Sub
    End Select
    If v Is Nothing Then Exit Sub
    
    Dim sh As Shape, sh2 As Shape
    Dim s As String
    For Each sh In sr
        s = re.Replace(sh.name, "")
        If bid Then s = s & " " & sh.id
        If s <> sh.name Then sh.name = s
        If sh.Type = msoGroup Then
            For Each sh2 In sh.GroupItems
                s = re.Replace(sh2.name, "")
                If bid Then s = s & " " & sh2.id
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
    If TypeName(Selection) = "Range" Then
        MsgBox "図形を選択してください。"
        Exit Sub
    End If
    If TypeName(Selection) = "DrawingObjects" Then
        MsgBox "複数の図形は選択できません。"
        Exit Sub
    End If
    
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

Public Sub PasteShapeNameList()

    If TypeName(Selection) = "Range" Then Exit Sub
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    If sr Is Nothing Then Exit Sub

End Sub

'原点合わせ
Public Sub OriginAlignment()
    
    If TypeName(Selection) = "Range" Then Exit Sub
    
    'オブジェクトコレクション作成
    Dim col As Collection
    Set col = New Collection
    Dim obj As Object
    If TypeName(Selection) = "DrawingObjects" Then
        For Each obj In Selection
            col.Add obj
        Next obj
    Else
        col.Add Selection
    End If
    
    '左下セル取得
    Dim s As String
    Dim ce As Range
    Dim x0 As Double, y0 As Double, y As Double
    Dim r As Long, c As Long
    For Each obj In col
        Set ce = obj.TopLeftCell
        y = obj.Top + obj.Height
        Do While y > ce.Top
            Set ce = ce.Offset(1)
        Loop
        If r = 0 Or r < ce.Row Then r = ce.Row
        If c = 0 Or c > ce.Column Then c = ce.Column
    Next obj
    Set ce = ce.Worksheet.Cells(r, c)
    
    '左下位置に移動
    For Each obj In col
        obj.Left = ce.Left
        obj.Top = ce.Top - obj.Height
    Next obj
    ce.Select

End Sub

'----------------------------------------
'部品描画
'----------------------------------------

'登録部品数取得
Sub DrawItemCount(ByRef cnt As Long)
    
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    cnt = ws.Shapes.Count

End Sub

'登録部品名取得
Sub DrawItemName(Index As Integer, ByRef name As String)
    
    If Index < 0 Then Exit Sub
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    If Index < ws.Shapes.Count Then
        name = ws.Shapes(1 + Index).name
    End If

End Sub

'登録部品選択
Sub DrawItemSelect(ByRef Index As Integer)
    
    Dim s As String
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If Not ws Is Nothing Then
        Dim i As Integer
        i = 1 + Index
        If i < 1 Then i = 1
        If i > ws.Shapes.Count Then i = ws.Shapes.Count
        If i > 0 Then s = ws.Shapes(i).name
    End If
    Call Draw_SetParam(4, s)

End Sub

'部品登録
Sub DrawItemEntry()
    
    If TypeName(Selection) = "Range" Then
        MsgBox "部品を選択してください。", vbOKOnly, app_name
        Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes", ThisWorkbook, True)
    
    '登録位置を計算
    Dim r As Long, i As Long
    Dim sh As Shape
    For Each sh In ws.Shapes
        i = sh.BottomRightCell.Row
        If i > r Then r = i
    Next sh
    Dim ce As Range
    Set ce = ws.Cells(r + 2, 2)

    '部品化
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    If sr.Count > 1 Then sr.Group
    Dim s As String
    s = InputBox("名前を入力してください。", app_name, sr.name)
    If s = "" Then Exit Sub
    sr.name = s
    UpdateShapeName sr
    sr.LockAspectRatio = msoTrue
    For Each sh In sr
        sh.Placement = xlMove
    Next sh

    'shapesシートに登録
    sr.Select
    Selection.Copy
    ws.Paste ce

    'アドインファイルなら保存
    If ws.Parent.name = ThisWorkbook.name Then
        ThisWorkbook.Save
    End If
    
End Sub

'部品削除
Sub DrawItemDelete()

    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    If g_part = "" Then Exit Sub
    
    ws.Shapes(g_part).Delete
    g_part = ""

    If ws.Parent.name = ThisWorkbook.name Then
        ThisWorkbook.Save
    End If
    
End Sub

'部品配置
Sub AddDrawItem()
    
    If g_part = "" Then Exit Sub
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    
    ws.Shapes(g_part).Copy
    
    Dim ra As Range
    If TypeName(Selection) = "Range" Then
        Set ra = Selection(1, 1)
    Else
        Set ra = Selection.TopLeftCell
    End If
    
    ScreenUpdateOn
    ScreenUpdateOff
    ra.Worksheet.Paste
    UpdateShapeName Selection.ShapeRange
    ScreenUpdateOn

End Sub

'部品コピー
Sub CopyDrawItem()
    
    If g_part = "" Then Exit Sub
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    
    ws.Shapes(g_part).Copy

End Sub

'部品シート複製
Sub DuplicateDrawItemSheet()
    
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    
    ws.Copy After:=ActiveSheet

End Sub

'部品シート取込
Sub ImportDrawItemSheet()
    
     Dim ws As Worksheet
    Set ws = SearchName(ThisWorkbook.Worksheets, "#shapes")
    
    If Not ws Is Nothing Then ws.Delete
    If Not UpdateAddinSheet(ActiveSheet) Then
    End If

End Sub

'----------------------------------------
'図形情報
'----------------------------------------

'図形情報リストアップ
Public Sub ListShapeInfo()

    If TypeName(Selection) <> "Range" Then
        Call AddShapeListName
        Call AddListShapeHeader(ActiveCell, 2)
    End If
    
    Call UpdateShapeInfo

End Sub

'----------------------------------------

'ヘッダ追加
Public Sub AddListShapeHeader(ByVal ce As Range, Optional mode As Long)
    
    'テーブル項目取得
    Dim dic As Dictionary
    ArrStrToDict dic, c_ShapeInfoMember, 1
    
    'ヘッダ取得
    Dim ra As Range
    Set ra = TableHeaderRange(TableLeftTop(ce))
    
    '存在する項目取得
    Dim hdr_dic As Dictionary
    Set hdr_dic = New Dictionary
    Dim s As String
    Dim v As Variant
    If ra.Count > 1 Then
        For Each v In ra.Value
            s = UCase(Trim(v))
            If s <> "" Then
                If dic.Exists(s) Then s = dic(s)(0)
                If Not hdr_dic.Exists(s) Then hdr_dic.Add s, 1
            End If
        Next v
    Else
        v = ra.Value
        s = UCase(Trim(v))
        If s <> "" Then
            If dic.Exists(s) Then s = dic(s)(0)
            If Not hdr_dic.Exists(s) Then hdr_dic.Add s, 1
        End If
    End If
    
    '追加ヘッダ項目取得
    Dim hdr() As String
    StringToRow hdr, c_ShapeInfoHeader, mode
    Dim hdr_col As Collection
    Set hdr_col = New Collection
    For Each v In hdr
        s = UCase(v)
        If dic.Exists(s) Then
            v = dic(s)(1)
            s = dic(s)(0)
        End If
        If Not hdr_dic.Exists(s) Then
            hdr_dic.Add s, 1
            hdr_col.Add v
        End If
    Next v
    If hdr_col.Count < 1 Then Exit Sub
    ReDim v(1 To 1, 1 To hdr_col.Count)
    Dim i As Long
    For i = 1 To hdr_col.Count
        v(1, i) = hdr_col(i)
    Next i
    If ra.Cells(1, 1).Value <> "" Then
        Set ra = ra.Offset(, ra.Columns.Count)
    End If
    Set ra = ra.Resize(1, hdr_col.Count)
    
    'ヘッダ追加
    ScreenUpdateOff
    ra.Value = v
    ScreenUpdateOn

End Sub

'ヘッダ追加
Public Sub AddShapeHeader(ByVal ce As Range, Optional mode As Long)
    
    'テーブル項目取得
    Dim dic As Dictionary
    ArrStrToDict dic, c_ShapeInfoMember, 1
    
    'ヘッダ取得
    Dim ra As Range
    Set ra = TableHeaderRange(TableLeftTop(ce))
    
    '存在する項目取得
    Dim hdr_dic As Dictionary
    Set hdr_dic = New Dictionary
    Dim s As String
    Dim v As Variant
    If ra.Count > 1 Then
        For Each v In ra.Value
            s = UCase(Trim(v))
            If s <> "" Then
                If dic.Exists(s) Then s = dic(s)(0)
                If Not hdr_dic.Exists(s) Then hdr_dic.Add s, 1
            End If
        Next v
    Else
        v = ra.Value
        s = UCase(Trim(v))
        If s <> "" Then
            If dic.Exists(s) Then s = dic(s)(0)
            If Not hdr_dic.Exists(s) Then hdr_dic.Add s, 1
        End If
    End If
    
    '追加ヘッダ項目取得
    Dim hdr() As String
    StringToRow hdr, c_ShapeInfoHeader, mode
    Dim hdr_col As Collection
    Set hdr_col = New Collection
    For Each v In hdr
        s = UCase(v)
        If dic.Exists(s) Then
            v = dic(s)(1)
            s = dic(s)(0)
        End If
        If Not hdr_dic.Exists(s) Then
            hdr_dic.Add s, 1
            hdr_col.Add v
        End If
    Next v
    If hdr_col.Count < 1 Then Exit Sub
    ReDim v(1 To 1, 1 To hdr_col.Count)
    Dim i As Long
    For i = 1 To hdr_col.Count
        v(1, i) = hdr_col(i)
    Next i
    If ra.Cells(1, 1).Value <> "" Then
        Set ra = ra.Offset(, ra.Columns.Count)
    End If
    Set ra = ra.Resize(1, hdr_col.Count)
    
    'ヘッダ追加
    ScreenUpdateOff
    ra.Value = v
    ScreenUpdateOn

End Sub

'----------------------------------------

'名前追加
Public Sub AddShapeListName()
    
    If TypeName(Selection) = "Range" Then Exit Sub
    
    '出力先セル/図形リスト取得
    Dim sr As Variant
    Set sr = Selection.ShapeRange
    Dim ce As Range
    Set ce = GetCell("リスト出力位置を指定してください", "図形リスト出力")
    If ce Is Nothing Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = sr.Parent
    
    'テーブル開始取得
    Set ce = TableLeftTop(ce)
    If ce.Value = "" Then ce.Value = "名前"
    
    'テーブル最終行取得
    Dim ce2 As Range
    Set ce2 = ce
    If ce2.Offset(1).Value <> "" Then Set ce2 = ce2.End(xlDown)
    
    '除外辞書作成
    Dim r_dic As Dictionary
    Set r_dic = New Dictionary
    
    'テーブルの項目は除外リストに追加
    Dim v As Variant
    Dim s As String
    If ce.Parent.Range(ce, ce2).Count > 1 Then
        For Each v In ce.Parent.Range(ce, ce2).Value
            s = UCase(Trim(v))
            If Not r_dic.Exists(s) Then r_dic.Add s, v
        Next v
    Else
        v = ce.Value
        s = UCase(Trim(v))
        If Not r_dic.Exists(s) Then r_dic.Add s, v
    End If
    
    '除外対象でなければテーブルに項目追加
    Dim sh As Shape, sh2 As Shape
    For Each sh In sr
        v = sh.name
        s = UCase(Trim(v))
        If Not r_dic.Exists(s) Then
            If s <> "" Then
                Set ce2 = ce2.Offset(1)
                ce2.Value = v
                r_dic.Add s, v
            End If
        End If
        If sh.Type = msoGroup Then
            For Each sh2 In sh.GroupItems
                v = sh2.name
                s = UCase(Trim(v))
                If Not r_dic.Exists(s) Then
                    If s <> "" Then
                        Set ce2 = ce2.Offset(1)
                        ce2.Value = "  " & v
                        r_dic.Add s, v
                    End If
                End If
            Next sh2
        End If
    Next sh
    
    '除外辞書削除
    r_dic.RemoveAll
    Set r_dic = Nothing
    
    ce.Select
    
End Sub

'----------------------------------------

'名前リストから図形選択
Public Sub SelectShapeName()

    If TypeName(Selection) <> "Range" Then Exit Sub
    
    '名前マスク設定
    Dim match_ptn As String
    Dim match_flg As Boolean
    match_ptn = g_mask
    match_flg = True
    If Left(match_ptn, 1) = "!" Then
        match_ptn = Mid(match_ptn, 2)
        match_flg = False
    End If
    If match_ptn = "" Then match_ptn = ".*"
    
    'テーブル開始・最終行取得
    Dim ce As Range, ce2 As Range
    Set ce = TableLeftTop(Selection)
    Set ce2 = ce
    If ce2.Value <> "" Then
        If ce2.Offset(1).Value <> "" Then
            Set ce2 = ce2.End(xlDown)
        End If
    End If
    
    Dim arr As Variant
    Dim cnt As Long
    Dim v As Variant
    Dim s As String
    Dim sh As Shape
    
    '名前リスト作成
    Dim s_arr() As String
    cnt = 0
    If ce.Parent.Range(ce, ce2).Count > 1 Then
        arr = ce.Parent.Range(ce, ce2).Value
        ReDim s_arr(1 To UBound(arr))
        On Error Resume Next
        For Each v In arr
            Set sh = Nothing
            Set sh = ce.Parent.Shapes(Trim(v))
            If Not sh Is Nothing Then
                s = sh.name
                If re_test(s, match_ptn) = match_flg Then
                    cnt = cnt + 1
                    s_arr(cnt) = s
                End If
            End If
        Next v
        On Error GoTo 0
    Else
        'テーブルに名前が無い場合
        ReDim s_arr(1 To ce.Parent.Shapes.Count)
        For Each sh In ce.Parent.Shapes
            s = sh.name
            If re_test(s, match_ptn) = match_flg Then
                cnt = cnt + 1
                s_arr(cnt) = s
            End If
        Next sh
    End If
    If cnt = 0 Then Exit Sub
    ReDim Preserve s_arr(1 To cnt)
        
    '選択
    ScreenUpdateOff
    ce.Parent.Shapes.Range(s_arr).Select
    ScreenUpdateOn
    
End Sub

'----------------------------------------

'図形情報更新
Public Sub UpdateShapeInfo()
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Dim ce As Range, ce2 As Range
    Dim r As Long, c As Long, rcnt As Long, ccnt As Long
    Dim v As Variant
    Dim s As String
    Dim sh As Shape, sh2 As Shape
    
    'テーブル開始位置取得
    Set ce = TableLeftTop(Selection)
    
    'テーブルヘッダ取得
    Dim hdr() As String
    Dim hdr_ra As Range
    Dim hdr_dic As Dictionary
    Set hdr_ra = TableHeaderRange(ce)
    ccnt = hdr_ra.Count
    If ccnt < 2 Then Exit Sub
    ReDim hdr(1 To ccnt)
    ArrStrToDict hdr_dic, c_ShapeInfoMember, 1
    c = 0
    For Each v In hdr_ra.Value
        c = c + 1
        s = UCase(Trim(v))
        If hdr_dic.Exists(s) Then
            hdr(c) = hdr_dic(UCase(Trim(v)))(0)
        Else
            hdr(c) = ""
        End If
    Next v
    
    '名前マスク設定
    Dim match_ptn As String
    Dim match_flg As Boolean
    match_ptn = g_mask
    match_flg = True
    If Left(match_ptn, 1) = "!" Then
        match_ptn = Mid(match_ptn, 2)
        match_flg = False
    End If
    If match_ptn = "" Then match_ptn = ".*"
    
    'テーブルデータ取得
    Dim tbl_arr As Variant
    tbl_arr = TableRange(hdr_ra).Value
    rcnt = UBound(tbl_arr, 1)
    For r = 2 To rcnt
        Set sh = Nothing
        On Error Resume Next
        s = Trim(tbl_arr(r, 1))
        Set sh = ce.Parent.Shapes(s)
        On Error GoTo 0
        If Not sh Is Nothing Then
            If re_test(sh.name, match_ptn) = match_flg Then
                For c = 2 To ccnt
                    s = hdr(c)
                    If s <> "" Then
                        tbl_arr(r, c) = ShapeValue(sh, s, "")
                    End If
                Next c
            End If
        Else
            For c = 2 To ccnt
                tbl_arr(r, c) = Empty
            Next c
        End If
    Next r
    
    '画面更新停止
    ScreenUpdateOff
    
    '表示形式設定
    Dim ra As Range
    For c = 1 To ccnt
        Set ra = ce.Parent.Range(ce.Cells(2, c), ce.Cells(rcnt, c))
        s = UCase(hdr(c))
        If hdr_dic.Exists(s) Then
            Select Case CInt(hdr_dic(s)(2))
            Case 1: ra.NumberFormatLocal = "0"
            Case 2: ra.NumberFormatLocal = "0.0##"
            Case Else: ra.NumberFormatLocal = "@"
            End Select
        End If
    Next c
    
    'テーブルデータ反映
    ce.Resize(rcnt, ccnt).Value = tbl_arr
    
    '配色
    For c = 1 To ccnt
        Set ra = ce.Parent.Range(ce.Cells(2, c), ce.Cells(rcnt, c))
        s = UCase(hdr(c))
        If hdr_dic.Exists(s) Then
            Select Case CInt(hdr_dic(s)(2))
            Case 4  '色
                For r = 2 To rcnt
                    s = ce.Cells(r, c)
                    If s <> "" Then
                        ce.Cells(r, c).Interior.Color = ToRGB(s)
                        SetFontColorFromInterior ce.Cells(r, c)
                    Else
                        ce.Cells(r, c).ClearFormats
                    End If
                Next r
            End Select
        End If
    Next c
    
    'テーブル調整
    WakuBorder TableRange(TableHeaderRange(ce))
    SetHeaderColor ce
    HeaderAutoFit ce
    
    '画面更新再開
    ScreenUpdateOn

End Sub

'シート取得
Public Sub LinkedSheet(ws As Worksheet, s As String)
    
    If s = "" Then Exit Sub
    
    Dim ss() As String
    ss = Split(Replace(s, "]", ""), "[", 2)
    If UBound(ss) < 0 Then Exit Sub
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If UBound(ss) = 1 Then
        For Each wb In Workbooks
            If UCase(wb.name) = UCase(ss(1)) Then Exit For
        Next wb
        If wb Is Nothing Then Exit Sub
    End If
    
    Dim ws1 As Worksheet
    For Each ws1 In wb.Sheets
        If UCase(ws1.name) = UCase(ss(0)) Then Exit For
    Next ws1
    If Not ws1 Is Nothing Then Set ws = ws1

End Sub

'文字列から行配列を取得
Sub StringToRow(arr() As String, info As String, Optional mode As Long = 0)
    
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

'----------------------------------------

'図形情報の適用
Public Sub ApplyShapeInfo(ByVal ra As Range, Optional ByVal ws As Worksheet)
    
    If Not TypeName(Selection) = "Range" Then Exit Sub
    If ws Is Nothing Then Set ws = ActiveSheet
    If ra Is Nothing Then Set ra = ActiveCell
    
    'テーブル開始位置を取得
    Dim ce As Range
    Set ce = TableLeftTop(ra)
    
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
        If Not dic.Exists(sh.name) Then dic.Add sh.name, 1
        If sh.Type = msoGroup Then
            For Each sh2 In sh.GroupItems
                If Not dic.Exists(sh2.name) Then dic.Add sh2.name, 1
            Next sh2
        End If
    Next sh
    Set sh = Nothing
    Set sh2 = Nothing
    
    '画面更新停止
    ScreenUpdateOn
    
    Dim s As String
    s = Trim(ce.Value)
    Do Until s = ""
        If dic.Exists(s) Then
            Set sh = ws.Shapes(s)
            If sh.Type = msoGroup Then
                Dim f As Boolean
                f = sh.ThreeD.Visible
                If f Then sh.ThreeD.Visible = msoFalse
            End If
            Dim c As Integer
            For c = 1 To UBound(hdr)
                Dim h As String
                h = hdr(c)
                Dim v1 As Variant, v2 As Variant
                v1 = ShapeValue(sh, h)
                v2 = ce.Offset(, c).Value
                If v1 <> v2 Then
                      ApplyShapeValue sh, h, v2
                End If
            Next c
        End If
        If f Then sh.ThreeD.Visible = msoTrue
        Set ce = ce.Offset(1)
        s = Trim(ce.Value)
    Loop
    Set dic = Nothing
    '
    '画面更新再開
    ScreenUpdateOn

End Sub


'----------------------------------------

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
    Case "BACK": v = sh.ThreeD.Z
    Case "ROTATION": v = sh.Rotation
    
    Case "HEIGHT": v = sh.Height
    Case "WIDTH": v = sh.Width
    Case "DEPTH": v = sh.ThreeD.Depth
    
    Case "VISIBLE": v = CBool(sh.Visible)
    
    Case "LINEVISIBLE": v = CBool(sh.line.Visible)
    Case "LINECOLOR": v = FormatColor(sh.line.ForeColor)
    
    Case "FILLVISIBLE": v = CBool(sh.Fill.Visible)
    Case "FILLCOLOR": v = FormatColor(sh.Fill.ForeColor)
    Case "TRANSPARENCY": v = sh.Fill.Transparency
    
    Case "TEXT": v = sh.TextFrame2.TextRange.text
    Case "ALTTEXT": v = sh.AlternativeText
    
    Case "SCALE": v = Replace(re_match(sh.AlternativeText, "sc:[+-]?[\d.]+"), "sc:", "")
    Case "X0": v = re_match(sh.AlternativeText, "p:([+-]?[\d.]+),([+-]?[\d.]+),([+-]?[\d.]+)", 0, 0)
    Case "Y0": v = re_match(sh.AlternativeText, "p:([+-]?[\d.]+),([+-]?[\d.]+),([+-]?[\d.]+)", 0, 1)
    Case "Z0": v = re_match(sh.AlternativeText, "p:([+-]?[\d.]+),([+-]?[\d.]+),([+-]?[\d.]+)", 0, 2)
    Case "DX": v = re_match(sh.AlternativeText, "d:([+-]?[\d.]+),([+-]?[\d.]+),([+-]?[\d.]+)", 0, 0)
    Case "DY": v = re_match(sh.AlternativeText, "d:([+-]?[\d.]+),([+-]?[\d.]+),([+-]?[\d.]+)", 0, 1)
    Case "DZ": v = re_match(sh.AlternativeText, "d:([+-]?[\d.]+),([+-]?[\d.]+),([+-]?[\d.]+)", 0, 2)
    
    End Select
    On Error GoTo 0
    ShapeValue = v

End Function

'図形情報設定
Private Sub ApplyShapeValue(sh As Shape, k As String, ByVal v As Variant)
    
    On Error Resume Next
    Select Case UCase(k)
    
    Case "NAME": sh.name = CStr(v)
    Case "TITLE": sh.Title = CStr(v)
    Case "ID":
    Case "TYPE":
    Case "STYLE":
    
    Case "TOP": sh.Top = CSng(v)
    Case "LEFT": sh.Left = CSng(v)
    Case "BACK": sh.ThreeD.Z = CSng(v)
    Case "ROTATION": sh.Rotation = CSng(v)
    
    Case "HEIGHT": sh.Height = CSng(v)
    Case "WIDTH": sh.Width = CSng(v)
    Case "DEPTH": sh.ThreeD.Depth = CSng(v)
    
    Case "VISIBLE": sh.Visible = CBool(v)
    
    Case "LINEVISIBLE": sh.line.Visible = CBool(v)
    Case "LINECOLOR": sh.line.ForeColor.RGB = ToRGB(v)
    
    Case "FILLVISIBLE": sh.Fill.Visible = CBool(v)
    Case "FILLCOLOR": sh.Fill.ForeColor.RGB = ToRGB(v)
    Case "TRANSPARENCY": sh.Fill.Transparency = v
    
    Case "TEXT": sh.TextFrame2.TextRange.text = v
    Case "ALTTEXT": sh.AlternativeText = v
    
    Case "SCALE": sh.AlternativeText = UpdateParamStr(sh.AlternativeText, "sc", CStr(v))
    'Case "X0": sh.AlternativeText = UpdateParamStr(sh.AlternativeText, "x0", CStr(v))
    'Case "Y0": sh.AlternativeText = UpdateParamStr(sh.AlternativeText, "y0", CStr(v))
    
    End Select
    On Error GoTo 0

End Sub

'----------------------------------------
'図形タイプ名称
'----------------------------------------

Private Function shape_typename(id As Integer) As String
    shape_typename = id
    If ptypename Is Empty Then InitDrawing
    If id < 0 Then id = UBound(ptypename)
    If id <= UBound(ptypename) Then shape_typename = ptypename(id)
End Function

Private Function shape_shapetypename(id As Integer) As String
    shape_shapetypename = id
    If pshapetypename Is Empty Then InitDrawing
    If id < 0 Then id = UBound(pshapetypename)
    If id <= UBound(pshapetypename) Then shape_shapetypename = pshapetypename(id)
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

'----------------------------------------
'図面シート
'----------------------------------------

'図面シートを追加
Public Sub AddDrawingSheet()
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=ActiveSheet)
    ws.Cells.ColumnWidth = 2.5
End Sub

'----------------------------------------
'タイムチャート
'----------------------------------------

'タイムチャート作図
Public Sub DrawTimeChart(Optional mode As Long)
    Dim ce As Range
    Set ce = ActiveCell
    If ce.Value = "" Then Exit Sub
    
    'データ範囲取得
    Dim ra As Range
    Set ra = ce
    If ce.Offset(, 1) <> "" Then
        Set ra = ce.Worksheet.Range(ce, ce.End(xlToRight))
    End If
    If ra.Count < 2 Then
        Dim v As Variant
        v = Application.InputBox("横サイズを指定してください。", app_name, "16", , , , , Type:=1)
        If TypeName(v) = "Boolean" Then Set ra = ra.Range(ra, ra.Offset(, v))
    End If
    
    Select Case mode
    Case 1: Call DrawTimeChart_1(ra)
    Case 2: Call DrawTimeChart_2(ra)
    End Select
End Sub

Private Sub DrawTimeChart_1(ByVal ra As Range)
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    
    Dim sh As Shape
    Dim fb As FreeformBuilder
    Dim ns As Collection
    Set ns = New Collection
    
    'データ範囲取得
    Set ra = ws.Range(ra, ra.End(xlToRight).Offset(, 1))
    
    '描画位置補正
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Do While Left(ce.Value, 1) = "!"
        Set ce = ce.Offset(-1)
    Loop
    If ce.Value = "" Then Set ce = ce.Offset(1)
    
    '状態位置取得
    Dim py(0 To 9) As Double
    Dim i As Long, j As Long
    j = 1
    For i = 0 To 9
        py(i) = ce.Offset(j).Top
        If ce.Row + j > 1 Then j = j - 1
    Next i
    
    Dim re As Object
    Set re = regex("[!#=.\[\]]")

    Dim x As Double, dx As Double, x0 As Double, x1 As Double, x2 As Double
    Dim y As Double, dy As Double, y0 As Double, y1 As Double, y2 As Double
    
    Dim xi As Double, yi As Double, yj As Double
    Dim yn As Long
    
    Dim ss As String, s As String, s0 As String
    Dim b_skip As Boolean
    Dim b_close As Boolean
    
    xi = ra.Left: yi = -1
    
    For Each ce In ra
        ss = CStr(ce.Value)
        i = Len(re.Replace(ss, ""))
        x = ce.Left: dx = (ce.Offset(, 1).Left - x) / IIf(i < 1, 1, i)
        x1 = x: x2 = x1 + dx
        '
        For i = 1 To Len(ss)
            s0 = s: s = Mid(ss, i, 1)
            x0 = x: If Not re.Test(s) Then x = x + dx
            '
            ' (x0,y0,s0)-(x1,y1,s)-[x,y,s]-(x2,y2,s)
            '0-9: レベル
            '-: 前を引き継ぐ
            '<: 位置保存
            '>: 位置保存
            '=: 位置保存
            Select Case s
            Case "0", "1", "2", "3", "4": yn = CLng(s): y = py(yn)
            Case "5", "6", "7", "8", "9": yn = CLng(s): y = py(yn)
            Case "\": yn = IIf(yn > 0, yn - 1, yn): y1 = y: y = py(yn): y0 = y
            Case "/": yn = IIf(yn < 9, yn + 1, yn): y1 = y: y = py(yn): y0 = y
            Case "=": xi = x0: y = y0: yn = yn Xor 1: yi = py(yn)
            Case "[": xi = x0: yi = y
            Case "<": xi = x0: yi = y
            Case ".": b_close = True
            End Select
            
            If s = "x" Then
                If yi < 0 Then y = py(yn): yn = yn Xor 1: y0 = py(yn): yi = y0
                If fb Is Nothing Then
                    Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, x0, y)
                End If
                fb.AddNodes msoSegmentLine, msoEditingAuto, x0, y
                fb.AddNodes msoSegmentLine, msoEditingAuto, x, yi
                Set sh = fb.ConvertToShape
                ns.Add sh.name
                Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, xi, yi)
                fb.AddNodes msoSegmentLine, msoEditingAuto, x0, yi
                fb.AddNodes msoSegmentLine, msoEditingAuto, x, y
                xi = x: x0 = x
            End If
            
            If x0 <> xi And y = yi And yi <> -1 And s <> "-" Then
                If Not fb Is Nothing Then
                    fb.AddNodes msoSegmentLine, msoEditingAuto, x0, y
                    fb.AddNodes msoSegmentLine, msoEditingAuto, xi, y
                    Set sh = fb.ConvertToShape
                    ns.Add sh.name
                    Set fb = Nothing
                End If
                xi = x0
                y = yi: yi = -1
            End If
            
            If fb Is Nothing Then
                If y <> 0 Then
                    Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, xi, y)
                End If
            ElseIf y <> y0 Then
                fb.AddNodes msoSegmentLine, msoEditingAuto, x0, y
            End If
            '
            If InStr("-<>[", s) = 0 Then
                If Not fb Is Nothing Then
                    fb.AddNodes msoSegmentLine, msoEditingAuto, x, y
                End If
                y0 = y
            End If
            
            '切断処理
            If b_close Then
                b_close = False
                If Not fb Is Nothing Then
                    Set sh = fb.ConvertToShape
                    ns.Add sh.name
                    Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, x, y)
                End If
            End If
            
        Next i
    Next ce
    
    If Not fb Is Nothing Then
        fb.AddNodes msoSegmentLine, msoEditingAuto, x, y
        Set sh = fb.ConvertToShape
        ns.Add sh.name
        Set fb = Nothing
    End If
    If x <> xi And yi <> -1 Then
        Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, xi, yi)
        fb.AddNodes msoSegmentLine, msoEditingAuto, x, yi
        Set sh = fb.ConvertToShape
        ns.Add sh.name
        Set fb = Nothing
    End If

    If ns.Count > 1 Then Set sh = ws.Shapes.Range(ColToArr(ns)).Group

End Sub

Private Sub DrawTimeChart_2(ByVal ra As Range)
    Dim ce As Range
    Set ce = ra(1, 1)
    
    'データ範囲取得
    If ce.Offset(, 1) <> "" Then
        Set ra = ce.Worksheet.Range(ce, ce.End(xlToRight))
    End If

    With ra
        .FormatConditions.Add Type:=xlExpression, Operator:=xlEqual, _
            Formula1:="=mod(" & ra(1, 1).Address(False, False) & ",2)=0"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Borders(xlBottom)
            .LineStyle = xlContinuous
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .FormatConditions(1).StopIfTrue = False
        
        .FormatConditions.Add Type:=xlExpression, Operator:=xlEqual, _
            Formula1:="=mod(" & ra(1, 1).Address(False, False) & ",2)=1"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Borders(xlTop)
            .LineStyle = xlContinuous
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .FormatConditions(1).StopIfTrue = False
    
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=" & ce.Address(False, False) & "<>" & ce.Offset(0, -1).Address(False, False)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Borders(xlLeft)
            .LineStyle = xlContinuous
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

'タイムチャートデータ生成
Public Sub GenerateTimeChart(Optional mode As Long)
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Dim ra As Range
    Set ra = Selection
    If ra(1, 1) <> "" And ra(1, 1).Offset(, 1) <> "" Then
       Set ra = ra.Worksheet.Range(ra(1, 1), ra(1, 1).End(xlToRight))
    End If
    
    Select Case mode
    Case 1: Call GenerateTimeChart_1(ra)
    Case 2: Call GenerateTimeChart_2(ra)
    Case 3: Call GenerateTimeChart_3(ra)
    Case 4: Call GenerateTimeChart_4(ra)
    Case 5: Call GenerateTimeChart_5(ra)
    Case 6: Call GenerateTimeChart_6(ra)
    Case 7: Call GenerateTimeChart_7(ra)
    Case 8: Call GenerateTimeChart_8(ra)
    Case 9: Call GenerateTimeChart_9(ra)
    End Select
End Sub

'
Private Function TimeChartRange(ra As Range, msg As String) As Range
    Dim c As Long, e As Long
    c = ra.Column
    e = c + ra.Columns.Count - 1
    On Error Resume Next
    Dim ce As Range
    Set ce = Application.InputBox(msg, app_name, Type:=8)
    On Error GoTo 0
    If ce Is Nothing Then Exit Function
    Set ce = ce.Worksheet.Cells(ce.Row, c)
    If ce(1, 1) <> "" And ce(1, 1).Offset(, 1) <> "" Then
       e = wsf.Max(e, ce(1, 1).End(xlToRight).Column)
    End If
    Set ra = ra.Worksheet.Range(ra(1, 1), ra(1, 1).Offset(, e - c))
    Set TimeChartRange = ce
End Function

'クロック
Private Sub GenerateTimeChart_1(ByVal ra As Range, Optional mode As Long)
    If ra.Count < 2 Then
        Dim v As Variant
        v = Application.InputBox("横サイズを指定してください。", app_name, "16", , , , , Type:=1)
        If TypeName(v) <> "Boolean" Then Set ra = ra.Worksheet.Range(ra, ra.Offset(, v - 1))
    End If
    ra.FormulaR1C1 = "=1-RC[-1]"
End Sub

'カウンタ
Private Sub GenerateTimeChart_2(ByVal ra As Range, Optional mode As Long)
    If ra.Count < 2 Then
        Dim v As Variant
        v = Application.InputBox("横サイズを指定してください。", app_name, "16", , , , , Type:=1)
        If TypeName(v) <> "Boolean" Then Set ra = ra.Worksheet.Range(ra, ra.Offset(, v - 1))
    End If
    
    Dim v1 As Variant, v2 As Variant, v3 As Variant
    v1 = Application.InputBox("カウント数を入力してください。", app_name, "16", Type:=0)
    If TypeName(v1) = "Boolean" Then Exit Sub
    If Left(v1, 1) = "=" Then v1 = Mid(v1, 2)
    v2 = Application.InputBox("ステップ数入力してください。", app_name, "1", Type:=0)
    If TypeName(v2) = "Boolean" Then Exit Sub
    If Left(v2, 1) = "=" Then v2 = Mid(v2, 2)
    v3 = Application.InputBox("開始値を入力してください。", app_name, "0", Type:=0)
    If TypeName(v3) = "Boolean" Then Exit Sub
    '
    ra.FormulaR1C1 = "=MOD(" & v2 & "+RC[-1]," & v1 & ")"
    ra.Cells(1, 1).FormulaR1C1 = v3
End Sub

'ロジック
Private Sub GenerateTimeChart_3(ByVal ra As Range, Optional mode As Long)
    Dim ce As Range
    Set ce = TimeChartRange(ra, "信号を選択してください。")
    ra.Formula = "=1-" & ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
End Sub

'NOT
Private Sub GenerateTimeChart_4(ByVal ra As Range, Optional mode As Long)
    Dim ce As Range
    Set ce = TimeChartRange(ra, "NOT:信号を選択してください。")
    ra.Formula = "=1-" & ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
End Sub

'AND
Private Sub GenerateTimeChart_5(ByVal ra As Range, Optional mode As Long)
    Dim exp As String
    Dim ce As Range
    Dim i As Long
    For i = 1 To 8
        Set ce = TimeChartRange(ra, "AND:信号を選択してください。")
        If ce Is Nothing Then Exit For
        If exp <> "" Then exp = exp & "*"
        exp = exp & ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
    Next i
    ra.Formula = "=" & exp
End Sub

'OR
Private Sub GenerateTimeChart_6(ByVal ra As Range, Optional mode As Long)
    Dim exp As String
    Dim ce As Range
    Dim i As Long
    For i = 1 To 8
        Set ce = TimeChartRange(ra, "OR:信号を選択してください。")
        If ce Is Nothing Then Exit For
        If exp <> "" Then exp = exp & "*"
        exp = exp & "(1-" & ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False) & ")"
    Next i
    ra.Formula = "=1-" & exp
End Sub

'XOR
Private Sub GenerateTimeChart_7(ByVal ra As Range, Optional mode As Long)
    Dim exp As String
    Dim ce As Range
    Dim i As Long
    For i = 1 To 8
        Set ce = TimeChartRange(ra, "XOR:信号を選択してください。")
        If ce Is Nothing Then Exit For
        If exp <> "" Then exp = exp & "+"
        exp = exp & ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
    Next i
    ra.Formula = "=mod(" & exp & ",2)"
End Sub

'SEL
Private Sub GenerateTimeChart_8(ByVal ra As Range, Optional mode As Long)
    Dim exp As String, s As String
    Dim ce As Range
    Set ce = TimeChartRange(ra, "SEL:SEL信号を選択してください。")
    If ce Is Nothing Then Exit Sub
    exp = ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
    
    Set ce = TimeChartRange(ra, "SEL:信号A(SEL=0)を選択してください。")
    If ce Is Nothing Then Exit Sub
    s = ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
    
    Set ce = TimeChartRange(ra, "SEL:信号B(SEL=1)を選択してください。")
    If ce Is Nothing Then Exit Sub
    
    exp = exp & "*(" & ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
    exp = exp & "-" & s & ")+" & s
    ra.Formula = "=" & exp
End Sub

'DFF
Private Sub GenerateTimeChart_9(ByVal ra As Range, Optional mode As Long)
    Dim exp As String
    Dim ce As Range
    Set ce = TimeChartRange(ra, "DFF:クロック信号を選択してください。")
    If ce Is Nothing Then Exit Sub
    exp = ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
    
    Set ce = TimeChartRange(ra, "DFF:イネーブル信号を選択してください。")
    If Not ce Is Nothing Then
        exp = exp & "*" & ra(1, 1).Offset(ce.Row - ra.Row, -1).Address(False, False)
    End If
    
    Set ce = TimeChartRange(ra, "DFF:データ信号を選択してください。")
    If ce Is Nothing Then Exit Sub
    
    exp = exp & "*(" & ra(1, 1).Offset(ce.Row - ra.Row, -1).Address(False, False)
    exp = exp & "-" & ra(1, 1).Offset(0, -1).Address(False, False) & ")"
    exp = exp & "+" & ra(1, 1).Offset(0, -1).Address(False, False)
    ra.Formula = "=" & exp
End Sub

'----------------------------------------
'図操作
'----------------------------------------

'裏表反転
Sub FlipShapes()

    If TypeName(Selection) = "Range" Then
        MsgBox "図形を選択してください。", vbOKOnly, app_name
        Exit Sub
    End If

    '対象図形取得
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    
    'グループ解除
    Dim rx As Double, ry As Double, rz As Double
    Dim rv As Integer
    Dim grp_name As String
    If sr.Count = 1 Then
        If sr.Type <> msoGroup Then Exit Sub
        '3D情報取得
        With sr
            rx = .ThreeD.RotationX
            ry = .ThreeD.RotationY
            rz = .ThreeD.RotationZ
            rv = .ThreeD.Visible
        End With
        grp_name = sr.name
        Set sr = sr.Ungroup
    End If
    
    'コレクション作成
    Dim col As Collection
    Set col = New Collection
    Dim sh As Shape
    For Each sh In sr
        col.Add sh.name
    Next sh
    
    'ZOhder変更
    Dim i As Long
    For i = 1 To col.Count
        sr.Item(col(i)).ZOrder msoSendToBack
    Next i
    
    '3D 奥行設定
    Dim v As Variant
    For Each v In col
        Set sh = ActiveSheet.Shapes(v)
        If sh.Type = msoGroup Then
            Dim sh2 As Shape
            For Each sh2 In sh.GroupItems
                If sh2.ThreeD.Z < -100000 Then
                ElseIf sh2.ThreeD.Depth < -10000 Then
                Else
                    sh2.ThreeD.Z = sh2.ThreeD.Depth - sh2.ThreeD.Z
                End If
            Next sh2
        ElseIf sh.ThreeD.Z < -100000 Then
        ElseIf sh.ThreeD.Depth < -10000 Then
        Else
            sh.ThreeD.Z = sh.ThreeD.Depth - sh.ThreeD.Z
        End If
    Next v
    
    '再グループ化
    If grp_name <> "" Then
        If sr.Count > 1 Then sr.Group
        sr.name = grp_name
        grp_name = ""
    End If
    
    '3Dリセット
    If rv Then
        With sr.ThreeD
            .SetPresetCamera msoCameraIsometricTopUp
            .RotationX = rx
            .RotationY = ry
            .RotationZ = rz
        End With
    End If
    
    '左右反転
    sr.flip msoFlipHorizontal
    sr.Select

End Sub

