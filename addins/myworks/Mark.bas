Attribute VB_Name = "Mark"
Option Explicit
Option Private Module

Private Const mark_dx As Integer = 32
Private Const mark_dy As Integer = 20

'---------------------------------------------
'版数マーク
'---------------------------------------------

'版数設定値取得
Sub GetRevMark(ByRef text As Variant)
    Dim s As String
    s = GetParam("rev", "text")
    If s = "" Then
        s = "1"
        Call SetParam("rev", "text", s)
    End If
    text = s
End Sub

'版数設定値設定
Sub SetRevMark(ByRef text As String, Optional id As Integer)
    Dim s As String
    s = Trim(text)
    If s = "" Then Exit Sub
    Call SetParam("rev", "text", s)
    Call SetParam("rev", "id", id)
End Sub

'版数マーク追加
Sub AddRevMark(ra As Range)
    Dim s As String
    s = GetParam("rev", "text")
    Dim i As Integer
    i = 1 + LastRevIndex(s)
    Call DrawRevMark(Selection, s, i)
End Sub

Private Function LastRevIndex(text As String) As Integer
    'AddRevMarkの内部処理
    '指定した版数の最大個別図形番号取得
    Dim re As Object
    Set re = regex("^rev:" & text & "\b")
    '
    Dim id As Integer
    id = 0
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        Dim sp As Shape
        For Each sp In ws.Shapes
            If re.Test(sp.AlternativeText) Then
                Dim s As String
                s = re_match(sp.AlternativeText, "[/_-](\d+)", 0, 0)
                If s <> "" Then
                    Dim i As Integer
                    i = Val(s)
                    If i > id Then id = i
                End If
            End If
        Next sp
    Next ws
    LastRevIndex = id
End Function

Private Sub DrawRevMark(ra As Range, rev As String, id As Integer)
    'AddRevMarkの内部処理
    '版数マークの図形配置
    If id < 1 Then id = 1
    '
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    '
    Dim dx As Integer, dy As Integer
    dx = mark_dx
    dy = mark_dy
    Dim x As Long, y As Long
    Call RevMarkPos(ce, x, y, dx + 2, dy + 2)
    '
    Dim ws As Worksheet
    Set ws = ra.Parent
    Dim sp As Shape
    Set sp = ws.Shapes.AddShape(msoShapeIsoscelesTriangle, x, y, dx, dy)
    
    Dim a As String
    a = "rev:" & rev & "-" & id
    
    With ws.Shapes.Range(Array(ws.Shapes.Count))
        .ShapeStyle = msoShapeStylePreset1
        With .line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Weight = 1
            .Transparency = 0
        End With
        With .TextFrame2
            .WordWrap = msoFalse
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .TextRange.Characters.text = rev
            With .TextRange.Characters(1, Len(rev))
                With .Font.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(255, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
                With .ParagraphFormat
                    .FirstLineIndent = 0
                    .Alignment = msoAlignLeft
                End With
                With .Font
                    .Bold = msoTrue
                    .NameComplexScript = "+mn-cs"
                    .NameFarEast = "+mn-ea"
                    .Size = 14
                    .name = "+mn-lt"
                End With
            End With
        End With
        .LockAspectRatio = msoTrue
        With .TextFrame
            .VerticalOverflow = xlOartVerticalOverflowOverflow
            .HorizontalOverflow = xlOartHorizontalOverflowOverflow
        End With
        .Fill.Visible = msoFalse
        .AlternativeText = a
    End With
    sp.Placement = xlMove
End Sub

Private Sub RevMarkPos(ce As Range, ByRef x As Long, ByRef y As Long, dx As Integer, dy As Integer)
    'AddRevMarkの内部処理
    '版数マークの配置位置の調整
    x = ce.Left
    y = ce.Top
    On Error Resume Next
    If ce = "" Then
        Do While TestRevMarkPos(x, y, dx, dy)
            x = x + dx
        Loop
    ElseIf ce.Offset(, -1) = "" Then
        x = x - dx
        y = y - dy / 2
        Do While TestRevMarkPos(x, y, dx, dy)
            y = y + dy
        Loop
    ElseIf ce.Offset(, 1) = "" Then
        x = ce.Offset(, 1).Left
        y = y - dy / 2
        Do While TestRevMarkPos(x, y, dx, dy)
            x = x + dx
        Loop
    ElseIf ce.Offset(-1) = "" Then
        y = y - dy
        Do While TestRevMarkPos(x, y, dx, dy)
            x = x + dx
        Loop
    ElseIf ce.Offset(1) = "" Then
        y = ce.Offset(1).Top
        Do While TestRevMarkPos(x, y, dx, dy)
            x = x + dx
        Loop
    Else
        y = y - dy / 2
        Do While TestRevMarkPos(x, y, dx, dy)
            y = y + dy
        Loop
    End If
    On Error GoTo 0
End Sub

Private Function TestRevMarkPos(x As Long, y As Long, dx As Integer, dy As Integer) As Boolean
    'AddRevMark/RevMarkPosの内部処理
    '他の改版マークがかなっていないか調査。重なる場合はtrueを返す
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        Dim sp As Shape
        For Each sp In ws.Shapes
            If sp.AutoShapeType = msoShapeIsoscelesTriangle Then
                If Abs(sp.Top - y) < dy / 2 And Abs(sp.Left - x) < dx / 2 Then
                    TestRevMarkPos = True
                    Exit Function
                End If
            End If
        Next sp
    Next ws
    TestRevMarkPos = False
End Function

'版数マークリスト
Sub ListRevMark(ra As Range, Optional rev As String)
    If rev = "" Then Call GetRevMark(rev)
    '
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    '
    Dim bLink As Boolean
    Dim res As Integer
    res = MsgBox("配置セルへリンクしますか。", vbYesNoCancel + vbDefaultButton2, "版数マークリスト")
    If res = vbYes Then
        bLink = True
    ElseIf res = vbCancel Then
        Exit Sub
    End If
    '
    ScreenUpdateOff
    Dim ws As Worksheet
    For Each ws In ce.Parent.Parent.Sheets
        Dim sp As Shape
        For Each sp In ws.Shapes
            If sp.AutoShapeType = msoShapeIsoscelesTriangle Then
                If sp.TextFrame2.TextRange.text = rev Then
                    Dim s As String
                    s = sp.TopLeftCell.Address(False, False)
                    ce.Offset.Value = ws.name
                    If bLink Then
                        ce.Parent.Hyperlinks.Add _
                            Anchor:=ce.Offset(, 1), _
                            Address:="", _
                            SubAddress:=(ws.name & "!" & s), _
                            TextToDisplay:=s, _
                            ScreenTip:=rev & " 版"
                    Else
                        'TODO:簡易クリア
                        ce.Offset(, 1).Value = ""
                        ce.Offset(, 1).Font.ColorIndex = 0
                        ce.Offset(, 1).Font.Underline = False
                        ce.Offset(, 1).Value = s
                    End If
                    ce.Offset(, 2).Value = re_replace(sp.AlternativeText, "\s+", " ")
                    Set ce = ce.Offset(1)
                End If
            End If
        Next sp
    Next ws
    ScreenUpdateOn
End Sub

