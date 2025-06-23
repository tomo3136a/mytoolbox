Attribute VB_Name = "Style"
Option Explicit
Option Private Module


'---------------------------------------------
'スタイル
'---------------------------------------------

'スタイルを選択
Sub ProcStyle(id As Long)
    Select Case id
    Case 1: LoadStyle1
    Case Else: ShapeSetting id
    End Select
End Sub

'スタイル設定
Private Sub LoadStyle1(Optional id As Long = 1)
    Dim old As Variant
    Set old = Selection
    Dim sh As Shape
    '図形
    Set sh = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 10, 10, 10, 10)
    With sh
        .ShapeStyle = msoShapeStylePreset1
        SetTextStyle sh
        SetPropStyle sh
        .SetShapesDefaultProperties
        .Delete
    End With
    'ライン
    Set sh = ActiveSheet.Shapes.AddLine(10, 10, 20, 20)
    With sh
        .ShapeStyle = msoShapeStylePreset1
        SetPropStyle sh
        .SetShapesDefaultProperties
        .Delete
    End With
    'テキストボックス
    Set sh = ActiveSheet.Shapes.AddTextbox(msoTextOrientationDownward, 10, 10, 10, 10)
    With sh
        .ShapeStyle = msoShapeStylePreset1
        SetTextStyle sh
        SetPropStyle sh
        .SetShapesDefaultProperties
        .Delete
    End With
    old.Select
    MsgBox "スタイル" & id & "に設定しました。", vbOKOnly, app_name
End Sub

'プロパティ：移動・サイズ変更なし
Private Sub ShapeSetting(mode As Long)
    Dim sh As Shape, sh2 As Shape
    For Each sh In Selection.ShapeRange
        SetShapeSetting mode, sh
        If sh.Type = msoGroup Then
            For Each sh2 In sh.GroupItems
                SetShapeSetting mode, sh
            Next sh2
        End If
    Next sh
End Sub

Private Sub SetShapeSetting(mode As Long, sh As Shape)
    Select Case mode
    Case 2: sh.Placement = xlMove
    Case 3: SetStyle34 sh
    Case 4: SetStyle34 sh
        sh.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    Case 5:
        sh.TextFrame2.AutoSize = msoAutoSizeNone
        sh.TextFrame2.WordWrap = msoFalse
        sh.TextFrame2.VerticalAnchor = msoAnchorTop
        sh.TextFrame2.HorizontalAnchor = msoAnchorNone
        sh.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End Select
End Sub

Private Sub SetStyle34(sh As Shape)
    With sh
        With .TextFrame2
            .MarginLeft = 1
            .MarginRight = 1
            .MarginTop = 1
            .MarginBottom = 1
            .AutoSize = msoAutoSizeNone
            .WordWrap = msoFalse
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
        With .TextFrame
            .VerticalOverflow = xlOartVerticalOverflowOverflow
            .HorizontalOverflow = xlOartHorizontalOverflowOverflow
        End With
        .TextFrame2.AutoSize = msoAutoSizeNone
        .TextFrame2.Orientation = msoTextOrientationHorizontal
    End With
End Sub

'設定(テキスト)
Private Sub SetTextStyle(sh As Shape)
    With sh
        With .TextFrame2
            With .TextRange.Font.Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0
                .Transparency = 0
                .Solid
            End With
            With .TextRange.ParagraphFormat
                .Alignment = msoAlignLeft
            End With
            .MarginLeft = 1
            .MarginRight = 1
            .MarginTop = 1
            .MarginBottom = 1
            .AutoSize = msoAutoSizeNone
            .WordWrap = msoFalse
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
        End With
        With .TextFrame
            .VerticalOverflow = xlOartVerticalOverflowOverflow
            .HorizontalOverflow = xlOartHorizontalOverflowOverflow
        End With
        .TextFrame2.AutoSize = msoAutoSizeNone
        .TextFrame2.Orientation = msoTextOrientationHorizontal
    End With
End Sub
    
'設定(塗りつぶし)
Private Sub SetFillStyle(sh As Shape)
    With sh
        With .Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
    End With
End Sub
    
'設定(線)
Private Sub SetLineStyle(sh As Shape)
    With sh
        With .line
            .Visible = msoTrue
            .Weight = 1
            .Visible = msoTrue
        End With
    End With
End Sub
    
'設定(サイズとプロパティ)
Private Sub SetPropStyle(sh As Shape)
    With sh
        .LockAspectRatio = msoTrue
        .Placement = xlMove
    End With
End Sub

Private Sub SetPropRangeStyle(sh As Shape)
    sh.Select
    With Selection.ShapeRange
        .LockAspectRatio = msoTrue
        Dim sh2 As Shape
        For Each sh2 In Selection.ShapeRange
            sh2.Placement = xlMove
        Next sh2
    End With
End Sub

'---------------------------------------------
'カラーマーカ
'---------------------------------------------

'セルにカラーマーカを設定
Sub AddMarker(ra As Range, id As Integer, Optional ByVal kw As String)
    
    If kw = "" Then kw = Replace(Mid(Date, 5), "/", "")
    If InStr(1, kw, "_") = 0 Then kw = kw & "_" & (id Mod 10)
    
    Dim wb As Workbook
    Set wb = ra.Parent.Parent
    
    On Error Resume Next
    If wb.Styles(kw) Is Nothing Then
        With wb.Styles.Add(kw)
            .IncludeNumber = False
            .IncludeFont = False
            .IncludeAlignment = False
            .IncludeBorder = False
            .IncludePatterns = True
            .IncludeProtection = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                Select Case id Mod 10
                Case 0: .Color = RGB(255, 241, 0)   '黄
                Case 1: .Color = RGB(240, 125, 136) '赤
                Case 2: .Color = RGB(85, 171, 229)  '青
                Case 3: .Color = RGB(95, 190, 125)  '薄緑
                Case 4: .Color = RGB(185, 192, 203) '灰色
                Case 5: .Color = RGB(255, 140, 0)   '橙
                Case 6: .Color = RGB(51, 186, 177)  '青緑
                Case 7: .Color = RGB(163, 179, 103) '茶
                Case 8: .Color = RGB(168, 149, 226) '紫
                Case 9: .Color = RGB(2, 104, 2)     '緑
                End Select
                .TintAndShade = 0
            End With
        End With
    End If
    On Error GoTo 0
    ra.Style = kw

End Sub

'カラーマーカ削除
Sub DelMarker(ByVal ra As Range)
        
    Dim wb As Workbook
    Set wb = ra.Worksheet.Parent
        
    ScreenUpdateOff
    Dim ce As Range
    For Each ce In ra
        Dim kw As String
        kw = ra.Style
        If kw <> "Normal" Then wb.Styles(kw).Delete
    Next ce
    ScreenUpdateOn

End Sub

'カラーマーカリスト取得
Sub ListMarker()
    
    Dim ra As Range
    Set ra = GetCell("リストの出力先を指定してください。")
    If ra Is Nothing Then Exit Sub
    
    Dim wb As Workbook
    Set wb = ra.Parent.Parent
    
    Dim arr As Variant
    arr = re_extract(wb.Styles, "^\d{4}_\d{1,2}$")
    arr = wsf.Transpose(arr)
    If Not TypeName(arr) = "Variant()" Then Exit Sub
    ra.Resize(UBound(arr, 1), 1).Value = arr
    
    ScreenUpdateOff
    Dim v As Variant
    For Each v In arr
        ra.Style = v
        Set ra = ra.Offset(1)
    Next v
    ScreenUpdateOn
    
End Sub
