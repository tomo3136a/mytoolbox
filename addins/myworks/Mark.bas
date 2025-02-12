Attribute VB_Name = "Mark"
'==================================
'�Ő��}�[�N
'==================================

Option Explicit
Option Private Module

Private Const mark_dx As Integer = 10
Private Const mark_dy As Integer = 20
Private rev_comment As String

'----------------------------------------
'�@�\�Ăяo��
'----------------------------------------

Public Sub RevProc(id As Long, ra As Range, Optional ByRef res As Long)
    Dim v As Variant
    Dim rev As String
    Select Case id
    Case 1
        '�Ő��}�[�N�ǉ�
        v = InputBox("�ύX��������͂��Ă��������B", "�Ő��}�[�N", rev_comment)
        If StrPtr(v) = 0 Then Exit Sub
        rev_comment = Trim(v)
        Call AddRevMark(ra, rev_comment)
    Case 2
        '�Ő��ݒ�
        Call GetRevMark(rev)
        v = InputBox("�Ő�����͂��Ă��������B", "�Ő��}�[�N", rev)
        If StrPtr(v) = 0 Then Exit Sub
        rev = Trim(v)
        If rev = "" Then Exit Sub
        Call SetRevMark(rev)
        res = 1
    Case 3
        '�Ő����X�g�쐬
        Call GetRevMark(rev)
        v = InputBox("���X�g����Ő�����͂��Ă��������B", "�Ő��}�[�N", rev)
        If StrPtr(v) = 0 Then Exit Sub
        rev = Trim(v)
        If rev = "" Then Exit Sub
        Call ListRevMark(ra, rev)
    End Select
End Sub

'---------------------------------------------

'�Ő��ݒ�l�擾
Public Sub GetRevMark(ByRef v As Variant)
    If Not ExistsRtParam("rev", "text") Then Call SetRtParam("rev", "text", "1")
    v = GetRtParam("rev", "text")
End Sub

'�Ő��ݒ�l�ݒ�
Private Sub SetRevMark(v As String, Optional id As Integer)
    Dim s As String
    s = Trim(v)
    If s = "" Then Exit Sub
    Call SetRtParam("rev", "text", s)
    Call SetRtParam("rev", "index", CStr(id))
End Sub

'�Ő��}�[�N�ǉ�
Private Sub AddRevMark(ra As Range, comment As String)
    Dim s As String
    s = GetRtParam("rev", "text")
    If s = "" Then Exit Sub
    Call DrawRevMark(Selection, s, 1 + LastRevIndex(s), comment)
End Sub

'�w�肵���Ő��̍ő�ʐ}�`�ԍ��擾
Private Function LastRevIndex(s As String) As Integer
    Dim re As Object, re2 As Object
    Set re = regex("\brev:" & s & "\b")
    '
    Dim id As Integer
    id = 0
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        Dim sp As Shape
        For Each sp In ws.Shapes
            Dim s2 As String
            s2 = sp.AlternativeText
            If re.Test(s2) Then
                Dim s3 As String
                s3 = re_match(s2, "\bidx:(\d+)\b", 0, 0)
                If s3 <> "" Then
                    Dim i As Integer
                    i = Val(s3)
                    If i > id Then id = i
                End If
            End If
        Next sp
    Next ws
    LastRevIndex = id
End Function

'�Ő��}�[�N�̐}�`�z�u
Private Sub DrawRevMark(ra As Range, rev As String, id As Integer, comment As String)
    If id < 1 Then id = 1
    '
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    '
    Dim dx As Integer, dy As Integer
    dx = mark_dx * (1 + Len(rev))
    dy = mark_dy
    Dim x As Long, y As Long
    Call RevMarkPos(ce, x, y, dx + 2, dy + 2)
    '
    Dim ws As Worksheet
    Set ws = ra.Parent
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeIsoscelesTriangle, x, y, dx, dy)
    
    Dim a As String
    a = Trim(comment)
    a = UpdateParamStr(a, "rev", rev)
    a = UpdateParamStr(a, "idx", CStr(id))
    
    Dim sr As ShapeRange
    Set sr = ws.Shapes.Range(Array(ws.Shapes.Count))
    With sr
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
        .name = "����"
    End With
    sh.Placement = xlMove
    
    If False Then
        Dim s As String
        s = sr.name
        x = sr.Left
        y = sr.Top
        sr.Select
        Selection.Cut
        Selection.Worksheet.PasteSpecial 0
        Selection.name = s
        Selection.Left = x
        Selection.Top = y
        Application.CutCopyMode = 0
    End If
    
End Sub

Private Sub RevMarkPos(ce As Range, ByRef x As Long, ByRef y As Long, dx As Integer, dy As Integer)
    'AddRevMark�̓�������
    '�Ő��}�[�N�̔z�u�ʒu�̒���
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
    'AddRevMark/RevMarkPos�̓�������
    '���̉��Ń}�[�N�����Ȃ��Ă��Ȃ��������B�d�Ȃ�ꍇ��true��Ԃ�
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

Function NextRecord(ce As Range, Optional hdr As String) As Range
Dim ra As Range
    Set ra = ce.CurrentRegion
    ra.Select
    Set ra = TableLeftTop(ra)
    ra.Select
    Set ra = TableHeaderRange(ra)
    ra.Select
    Set ra = TableRange(ra)
    ra.Select
    Set ra = LeftBottom(ra)
    ra.Select
    If ra.Value = "" Then
        If hdr <> "" Then
            Dim ss As Variant
            ss = Split(hdr, ",")
            ra.Resize(1, UBound(ss) - LBound(ss) + 1).Value = ss
        End If
    End If
    If ra.Value <> "" Then Set ra = ra.Offset(1)
    ra.Select
    Set NextRecord = ra
End Function

'�Ő��}�[�N���X�g
Private Sub ListRevMark(ra As Range, Optional rev As String)
    If rev = "" Then Call GetRevMark(rev)
    '
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Set ce = NextRecord(ce, "�Ő�,�V�[�g,���W,����")
    '
    Dim bLink As Boolean
    Dim res As Integer
    res = MsgBox("�z�u�Z���փ����N���܂����B", vbYesNoCancel + vbDefaultButton2, "�Ő��}�[�N���X�g")
    If res = vbYes Then
        bLink = True
    ElseIf res = vbCancel Then
        bLink = False
    End If
    '
    ScreenUpdateOff
    Dim s As String
    'Dim ss As Variant
    'If ce = "" Then
    '    s = "�Ő�,�V�[�g,���W,����"
    '    ss = Split(s, ",")
    '    ce.Resize(1, UBound(ss) - LBound(ss) + 1).Value = ss
    '    Set ce = ce.Offset(1)
    'Else
    '    Set ce = NextRecord(ce)
    'End If
    ce.Select
    
    Dim ws As Worksheet
    For Each ws In ce.Parent.Parent.Sheets
        Dim sp As Shape
        For Each sp In ws.Shapes
            If sp.AutoShapeType = msoShapeIsoscelesTriangle Then
                If sp.TextFrame2.TextRange.text = rev Then
                    s = sp.AlternativeText
                    ce.Value = ParamStrVal(s, "rev")
                    Set ce = ce.Offset(, 1)
                    ce.Value = ws.name
                    Set ce = ce.Offset(, 1)
                    
                    s = sp.TopLeftCell.Address(False, False)
                    If bLink Then
                        ce.Parent.Hyperlinks.Add _
                            Anchor:=ce, _
                            Address:="", _
                            SubAddress:=(ws.name & "!" & s), _
                            TextToDisplay:=s, _
                            ScreenTip:=rev & " ��"
                    Else
                        'TODO:�ȈՃN���A
                        ce.Value = ""
                        ce.Font.ColorIndex = 0
                        ce.Font.Underline = False
                        ce.Value = s
                    End If
                    Set ce = ce.Offset(, 1)
                    ce.Value = Trim(RemoveParamStrAll(sp.AlternativeText))
                    s = sp.AlternativeText
                    s = re_replace(s, "\s*\w+:[^$\r\n]*[$\r\n]?", "")
                    ce.Value = Trim(s)
                    Set ce = ce.Offset(1, -3)
                End If
            End If
        Next sp
    Next ws
    ScreenUpdateOn
End Sub

