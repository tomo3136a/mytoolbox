Attribute VB_Name = "MDraw"
'==================================
'�`��
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'�}�`����ID

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
'�f�[�^

'�}�`���X�g�����o�[
'[0] ����
'[1] �\����
'[2] 0:������,1:����,2:0.0##,3:�_���l,4:�F
Private Const c_ShapeInfoMember As String = "" _
        & ";Name,���O,0     ;Title,�^�C�g��,0" _
        & ";" _
        & ";ID,ID,0         ;Type,���,0     ;Style,�X�^�C��,0" _
        & ";Top,��ʒu,2    ;Left,���ʒu,2   ;Back,��ʒu,2     ;Rotation,��],2" _
        & ";Height,����,2   ;Width,��,2      ;Depth,���s��,2" _
        & ";" _
        & ";Visible,�\��,3                  ;Transparency,�����x,2" _
        & ";LineVisible,�g���\��,3          ;LineColor,�g���F,4" _
        & ";FillVisible,�h��Ԃ��\��,3    ;FillColor,�h��Ԃ��F,4" _
        & ";" _
        & ";Text,�e�L�X�g,0     ;AltText,��ւ��e�L�X�g,0" _
        & ";Scale,�X�P�[��,2    ;X0,���_X,2     ;Y0,���_Y,2     ;Z0,���_Z,2" _
        & ";DX,�T�C�YX,2        ;DY,�T�C�YY,2   ;DZ,�T�C�YZ,2"

'�}�`���X�g�w�b�_
Private Const c_ShapeInfoHeader As String = "" _
    & ";����,           Name" _
    & ";�`��,           Name,ID,Type,Style,Title" _
    & ";�ʒu,           Name,Left,Top,Back,Rotation" _
    & ";�T�C�Y,         Name,Width,Height,Depth" _
    & ";�\��,           Name,Visible,Transparency" _
    & ";�g��,           Name,LineVisible,LineColor" _
    & ";�h��,           Name,FillVisible,FillColor" _
    & ";�e�L�X�g,       Name,Text" _
    & ";��ւ��e�L�X�g, Name,AltText" _
    & ";����,           Name,Scale,X0,Y0,Z0,DX,DY,DZ"

'���p�����[�^����
Public Enum E_DrawParam
    E_IGNORE = 1
    E_SCALE = 2
    E_AXES = 3
    E_FLAG = 4
    E_PART = 10
End Enum

'���p�����[�^
Private g_mask As String            '��������
Private g_scale As Double           '�X�P�[��
Private g_axes As Double            '���Ԋu
Private g_flag As Long              '���[�h
                                    ' 4.A��, 5.B��
                                    ' 6.�z�u����, 7.�z������
                                    ' 8.PTH, 9:Note
Private g_part As String            '���i��

Private ptype_col As Variant        '�}�`�^�C�v�e�[�u��
Private ptypename As Variant        '�}�`�^�C�v���̃e�[�u��
Private pshapetypename As Variant   '�}�`�^�C�v�e�[�u��

'----------------------------------------
'�p�����[�^����
'----------------------------------------

'�`��p�����[�^������
Public Sub Draw_ResetParam(Optional id As Integer)
    If id = 0 Or id = 1 Then g_mask = ""
    If id = 0 Or id = 2 Then g_scale = 0.1
    If id = 0 Or id = 3 Then g_axes = 10
    If id = 0 Or id = 4 Then g_part = ""
    If id = 0 Or id = 5 Then g_flag = 1 + 2
End Sub

'�`��p�����[�^�ݒ�
Public Sub Draw_SetParam(id As Integer, ByVal val As String)
    Select Case id
    Case 1: g_mask = val
    Case 2
        If val <= 0 Then
            MsgBox "�䗦�̐ݒ肪�Ԉ���Ă��܂��B(�ݒ�l>0)" & Chr(10) _
                & "�ݒ�l�F " & val
            Exit Sub
        End If
        g_scale = val
    Case 3
        If val <= 0 Then
            MsgBox "�ڐ���̐ݒ肪�Ԉ���Ă��܂��B(�ݒ�l>0)" & Chr(10) _
                & "�ݒ�l�F " & val
            Exit Sub
        End If
        g_axes = val
    Case 4: g_part = val
    End Select
End Sub

'�`��p�����[�^�擾
Public Function Draw_GetParam(id As Integer) As String
    Select Case id
    Case 1: Draw_GetParam = g_mask
    Case 2
        If g_scale <= 0 Then
            Draw_ResetParam id
            MsgBox "�䗦�̐ݒ�����������܂����B(�ݒ�l" & g_scale & ")"
        End If
        Draw_GetParam = g_scale
    Case 3
        If g_axes <= 0 Then
            Draw_ResetParam id
            MsgBox "�ڐ���̐ݒ�����������܂����B(�ݒ�l" & g_axes & ")"
        End If
        Draw_GetParam = g_axes
    Case 4
        Draw_GetParam = g_part
    End Select
End Function

'�`��p�����[�^�t���O�ݒ�
Public Sub Draw_SetParamFlag(id As Integer, Optional ByVal val As Boolean = True)
    g_flag = g_flag And Not 2 ^ (id Mod 24)
    If val Then g_flag = g_flag Or 2 ^ (id Mod 24)
End Sub

'�`��p�����[�^�t���O�`�F�b�N
Public Function Draw_IsParamFlag(id As Integer) As Boolean
    Draw_IsParamFlag = Not ((g_flag And 2 ^ (id Mod 24)) = 0)
End Function

'----------------------------------------
'�}�`��������
'----------------------------------------

'�}�`�����擾
Function GetShapeProperty(sr As ShapeRange, k As String) As String
    GetShapeProperty = ParamStrVal(sr.AlternativeText, k)
End Function

'�}�`�����ݒ�
Sub SetShapeProperty(sr As ShapeRange, k As String, v As String)
    sr.AlternativeText = UpdateParamStr(sr.AlternativeText, k, v)
End Sub

'----------------------------------------
'�F����
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

'�h��Ԃ��F�\���擾
Private Function FormatInterior(v As Interior) As String
    If v.ThemeColor < 0 Then Exit Function
    Dim s As String
    s = Right("00000000" & Hex(v.Color), 6)
    s = "#" & Mid(s, 5, 2) & Mid(s, 3, 2) & Mid(s, 1, 2)
    If v.ThemeColor > 0 Then s = s & " " & v.ThemeColor & Format(v.TintAndShade, "+0%;-0%;"""";@")
    FormatInterior = s
End Function

'�h��Ԃ��F����t�H���g�F�ݒ�
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

'�F�\���擾
Private Function FormatColor(v As ColorFormat) As String
    Dim s As String
    s = Right("00000000" & Hex(v), 6)
    s = "#" & Mid(s, 5, 2) & Mid(s, 3, 2) & Mid(s, 1, 2)
    'FormatColor = s & " " & v.Type & " " & v.SchemeColor & " " & v.ObjectThemeColor & " " & v.Brightness
    If v.ObjectThemeColor > 0 Then s = s & " " & v.ObjectThemeColor & Format(v.Brightness, "+0%;-0%;"""";@")
    FormatColor = s
End Function

'----------------------------------------
'�}�`��{�ݒ�
'----------------------------------------

'�}�`��{�ݒ�
'[3] �e�L�X�g, �h��Ԃ��Ɛ�
Public Sub SetShapeStyle(Optional ByVal sr As ShapeRange)
    If sr Is Nothing Then
        If TypeName(Selection) = "Range" Then Exit Sub
        Set sr = Selection.ShapeRange
    End If
    Call SetShapeSetting(sr, 3)
End Sub

'�W���}�`�ݒ�
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
    With ws.Shapes.AddShape(msoShapeOval, 10, 10, 10, 10)
        SetShapeSetting ws.Shapes.Range(.name), &H107
        .Delete
    End With
    With ws.Shapes.AddLine(10, 10, 20, 20)
        SetShapeSetting ws.Shapes.Range(.name), &H107
        .Delete
    End With
    With ws.Shapes.AddTextbox(msoTextOrientationDownward, 10, 10, 10, 10)
        SetShapeSetting ws.Shapes.Range(.name), &H107
        .Delete
    End With
End Sub

'----------------------------------------
'�}�`��{�ݒ�
'[1] �e�L�X�g
'[2] �h��Ԃ��Ɛ�
'[4] �T�C�Y�ƃv���p�e�B
'[8] ��ւ�����
'[256] �f�t�H���g�ݒ�
Private Sub SetShapeSetting(Optional ByVal sr As ShapeRange, Optional mode As Integer = 255)
    
    Dim sh As Shape
    On Error Resume Next
    
    '�ݒ�(�e�L�X�g)
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
    
    '�ݒ�(�h��Ԃ��Ɛ�)
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
    
    '�ݒ�(�T�C�Y�ƃv���p�e�B)
    If mode And 4 Then
        sr.LockAspectRatio = msoTrue
        sr.Placement = xlMove
        For Each sh In sr
            sh.Placement = xlMove
        Next sh
    End If
    '
    '�ݒ�(��ւ�����)
    If mode And 8 Then
        For Each sh In sr
            sh.AlternativeText = sh.name
        Next sh
    End If
    
    '�f�t�H���g�ݒ�
    If mode And 256 Then sr.SetShapesDefaultProperties

    On Error GoTo 0

End Sub

'�\��/��\�����]
Public Sub ToggleVisible(mode As Integer, Optional sr As ShapeRange)
    
    Dim sr2 As ShapeRange
    Set sr2 = sr
    If sr2 Is Nothing Then
        If TypeName(Selection) = "Range" Then Exit Sub
        Set sr2 = Selection.ShapeRange
    End If
    
    Select Case mode
    Case 0
        '�\��/��\�����]
        With sr2.Fill
            If .Visible = msoTrue Then
                .Visible = msoFalse
            Else
                .Visible = msoTrue
            End If
        End With
    Case 3
        '3D�\��/��\�����]
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

'�}�`���X�V
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
'�}�`����
'----------------------------------------

'�}�`���폜
Public Sub RemoveSharps(Optional ByVal ws As Worksheet)
    
    '�ΏۑI��
    If ws Is Nothing Then
        If TypeName(Selection) <> "Range" Then
            Selection.Delete
            Exit Sub
        End If
        If MsgBox("�S�}�`���폜���܂����H", vbYesNo) <> vbYes Then Exit Sub
        Set ws = ActiveSheet
    End If
    
    '��ʍX�V��~
    On Error Resume Next
    Application.ScreenUpdating = False
    
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    
    '��ʍX�V�ĊJ
    Application.ScreenUpdating = True
    On Error GoTo 0

End Sub

'�}�`���G�ɕϊ�
Public Sub ConvertToPicture()
    
    '�ΏۑI��
    If TypeName(Selection) = "Range" Then Exit Sub
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    If sr Is Nothing Then Exit Sub
    
    '��ʍX�V��~
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
    
    '��ʍX�V�ĊJ
    Application.ScreenUpdating = fsu

End Sub

Public Sub PasteShapeNameList()

    If TypeName(Selection) = "Range" Then Exit Sub
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    If sr Is Nothing Then Exit Sub
    

End Sub

'----------------------------------------
'���i�`��
'----------------------------------------

'�o�^���i���擾
Sub DrawItemCount(ByRef cnt As Long)
    
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    cnt = ws.Shapes.Count

End Sub

'�o�^���i���擾
Sub DrawItemName(Index As Integer, ByRef name As String)
    
    If Index < 0 Then Exit Sub
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    If Index < ws.Shapes.Count Then
        name = ws.Shapes(1 + Index).name
    End If

End Sub

'�o�^���i�I��
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

'���i�o�^
Sub DrawItemEntry()
    
    If TypeName(Selection) = "Range" Then Exit Sub
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes", ThisWorkbook)
    
    '�o�^�ʒu���v�Z
    Dim ce As Range
    Dim r As Long
    Dim sh As Shape
    For Each sh In ws.Shapes
        Set ce = sh.BottomRightCell
        If ce.Row > r Then r = ce.Row
    Next sh
    Set ce = ws.Cells(r + 2, 2)


    '���i��
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    If sr.Count > 1 Then
        Dim s As String
        s = InputBox("���O����͂��Ă��������B")
        If s = "" Then Exit Sub
        sr.Group
        sr.name = s
    End If
    UpdateShapeName sr
    sr.LockAspectRatio = msoTrue
    For Each sh In sr
        sh.Placement = xlMove
    Next sh

    'shapes�V�[�g�ɓo�^
    sr.Select
    Selection.Copy
    ws.Paste ce

    '�A�h�C���t�@�C���Ȃ�ۑ�
    If ws.Parent.name = ThisWorkbook.name Then
        ThisWorkbook.Save
    End If
    
End Sub

'���i�폜
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

'���i�z�u
Sub AddDrawItem()
    
    If g_part = "" Then Exit Sub
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    
    Dim ra As Range
    If TypeName(Selection) = "Range" Then
        Set ra = Selection
    Else
        Set ra = Selection.TopLeftCell
    End If
    
    ws.Shapes(g_part).Copy
    ra.Worksheet.Paste
    UpdateShapeName Selection.ShapeRange

End Sub

'���i�R�s�[
Sub CopyDrawItem()
    
    If g_part = "" Then Exit Sub
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    
    ws.Shapes(g_part).Copy

End Sub

'���i�V�[�g����
Sub DuplicateDrawItemSheet()
    
    Dim ws As Worksheet
    Set ws = GetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    
    ws.Copy After:=ActiveSheet

End Sub

'���i�V�[�g�捞
Sub ImportDrawItemSheet()
    
     Dim ws As Worksheet
    Set ws = SearchName(ThisWorkbook.Worksheets, "#shapes")
    
    If Not ws Is Nothing Then ws.Delete
    If Not UpdateAddinSheet(ActiveSheet) Then
    End If

End Sub

'----------------------------------------
'�}�`���
'----------------------------------------

'�}�`��񃊃X�g�A�b�v
Public Sub ListShapeInfo()

    If TypeName(Selection) <> "Range" Then
        Call AddShapeListName
        Call AddListShapeHeader(ActiveCell, 2)
    End If
    
    Call UpdateShapeInfo

End Sub

'----------------------------------------

'�w�b�_�ǉ�
Public Sub AddListShapeHeader(ByVal ce As Range, Optional mode As Integer)
    
    '�e�[�u�����ڎ擾
    Dim dic As Dictionary
    ArrStrToDict dic, c_ShapeInfoMember, 1
    
    '�w�b�_�擾
    Dim ra As Range
    Set ra = TableHeaderRange(TableLeftTop(ce))
    
    '���݂��鍀�ڎ擾
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
    
    '�ǉ��w�b�_���ڎ擾
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
    
    '�w�b�_�ǉ�
    ScreenUpdateOff
    ra.Value = v
    ScreenUpdateOn

End Sub

'�w�b�_�ǉ�
Public Sub AddShapeHeader(ByVal ce As Range, Optional mode As Integer)
    
    '�e�[�u�����ڎ擾
    Dim dic As Dictionary
    ArrStrToDict dic, c_ShapeInfoMember, 1
    
    '�w�b�_�擾
    Dim ra As Range
    Set ra = TableHeaderRange(TableLeftTop(ce))
    
    '���݂��鍀�ڎ擾
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
    
    '�ǉ��w�b�_���ڎ擾
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
    
    '�w�b�_�ǉ�
    ScreenUpdateOff
    ra.Value = v
    ScreenUpdateOn

End Sub

'----------------------------------------

'���O�ǉ�
Public Sub AddShapeListName()
    
    If TypeName(Selection) = "Range" Then Exit Sub
    
    '�o�͐�Z��/�}�`���X�g�擾
    Dim sr As Variant
    Set sr = Selection.ShapeRange
    Dim ce As Range
    Set ce = GetCell("���X�g�o�͈ʒu���w�肵�Ă�������", "�}�`���X�g�o��")
    If ce Is Nothing Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = sr.Parent
    
    '�e�[�u���J�n�擾
    Set ce = TableLeftTop(ce)
    If ce.Value = "" Then ce.Value = "���O"
    
    '�e�[�u���ŏI�s�擾
    Dim ce2 As Range
    Set ce2 = ce
    If ce2.Offset(1).Value <> "" Then Set ce2 = ce2.End(xlDown)
    
    '���O�����쐬
    Dim r_dic As Dictionary
    Set r_dic = New Dictionary
    
    '�e�[�u���̍��ڂ͏��O���X�g�ɒǉ�
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
    
    '���O�ΏۂłȂ���΃e�[�u���ɍ��ڒǉ�
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
    
    '���O�����폜
    r_dic.RemoveAll
    Set r_dic = Nothing
    
    ce.Select
    
End Sub

'----------------------------------------

'���O���X�g����}�`�I��
Public Sub SelectShapeName()

    If TypeName(Selection) <> "Range" Then Exit Sub
    
    '���O�}�X�N�ݒ�
    Dim match_ptn As String
    Dim match_flg As Boolean
    match_ptn = g_mask
    match_flg = True
    If Left(match_ptn, 1) = "!" Then
        match_ptn = Mid(match_ptn, 2)
        match_flg = False
    End If
    If match_ptn = "" Then match_ptn = ".*"
    
    '�e�[�u���J�n�E�ŏI�s�擾
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
    
    '���O���X�g�쐬
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
        '�e�[�u���ɖ��O�������ꍇ
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
        
    '�I��
    ScreenUpdateOff
    ce.Parent.Shapes.Range(s_arr).Select
    ScreenUpdateOn
    
End Sub

'----------------------------------------

'�}�`���X�V
Public Sub UpdateShapeInfo()
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Dim ce As Range, ce2 As Range
    Dim r As Long, c As Long, rcnt As Long, ccnt As Long
    Dim v As Variant
    Dim s As String
    Dim sh As Shape, sh2 As Shape
    
    '�e�[�u���J�n�ʒu�擾
    Set ce = TableLeftTop(Selection)
    
    '�e�[�u���w�b�_�擾
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
    
    '���O�}�X�N�ݒ�
    Dim match_ptn As String
    Dim match_flg As Boolean
    match_ptn = g_mask
    match_flg = True
    If Left(match_ptn, 1) = "!" Then
        match_ptn = Mid(match_ptn, 2)
        match_flg = False
    End If
    If match_ptn = "" Then match_ptn = ".*"
    
    '�e�[�u���f�[�^�擾
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
    
    '��ʍX�V��~
    ScreenUpdateOff
    
    '�\���`���ݒ�
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
    
    '�e�[�u���f�[�^���f
    ce.Resize(rcnt, ccnt).Value = tbl_arr
    
    '�z�F
    For c = 1 To ccnt
        Set ra = ce.Parent.Range(ce.Cells(2, c), ce.Cells(rcnt, c))
        s = UCase(hdr(c))
        If hdr_dic.Exists(s) Then
            Select Case CInt(hdr_dic(s)(2))
            Case 4  '�F
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
    
    '�e�[�u������
    WakuBorder TableRange(TableHeaderRange(ce))
    SetHeaderColor ce
    HeaderAutoFit ce
    
    '��ʍX�V�ĊJ
    ScreenUpdateOn

End Sub

'�V�[�g�擾
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

'�����񂩂�s�z����擾
Sub StringToRow(arr() As String, info As String, Optional mode As Integer = 0)
    
    Dim dic As Dictionary
    ArrStrToDict dic, info
    
    Dim kw As String
    kw = dic.Keys(mode)
    arr = dic(kw)
    arr = TakeArray(arr, 1)

End Sub

'�}�`���R�[�h��z��ɒǉ�
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

'�}�`���̓K�p
Public Sub ApplyShapeInfo(ByVal ra As Range, Optional ByVal ws As Worksheet)
    
    If Not TypeName(Selection) = "Range" Then Exit Sub
    If ws Is Nothing Then Set ws = ActiveSheet
    If ra Is Nothing Then Set ra = ActiveCell
    
    '�e�[�u���J�n�ʒu���擾
    Dim ce As Range
    Set ce = TableLeftTop(ra)
    
    '�e�[�u�����ڎ擾
    Dim hdr_dic As Dictionary
    ArrStrToDict hdr_dic, c_ShapeInfoMember, 1
    
    '�w�b�_�擾
    Dim hdr() As String
    hdr = GetHeaderArray(ce, hdr_dic)
    
    '�}�`���X�g�쐬
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
    
    '��ʍX�V��~
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
    '��ʍX�V�ĊJ
    ScreenUpdateOn

End Sub


'----------------------------------------

'�}�`���擾
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
    'Case "LINECOLOR": v = FormatRGB(sh.line.ForeColor)
    Case "LINECOLOR": v = FormatColor(sh.line.ForeColor)
    
    Case "FILLVISIBLE": v = CBool(sh.Fill.Visible)
    'Case "FILLCOLOR": v = FormatRGB(sh.Fill.ForeColor)
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

'�}�`���ݒ�
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
    Case "BACK": sh.ThreeD.z = CSng(v)
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
'�}�`�^�C�v����
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
        "�I�[�g�V�F�C�v", "�����o��", "�O���t", "�R�����g", "�t���[�t�H�[��", "Group", "���ߍ��� OLE �I�u�W�F�N�g", "�t�H�[�� �R���g���[��", "Line", _
        "�����N OLE �I�u�W�F�N�g", "�����N�摜", "OLE �R���g���[�� �I�u�W�F�N�g", "�摜", "�v���[�X�z���_�[", "�e�L�X�g����", "���f�B�A", "�e�L�X�g �{�b�N�X", _
        "�X�N���v�g �A���J�[", "�e�[�u��", "�L�����o�X", "�}", "�C���N", "�C���N �R�����g", "SmartArt �O���t�B�b�N", "Slicer", "Web �r�f�I", _
        "�R���e���c Office �A�h�C��", "�O���t�B�b�N", "�����N���ꂽ�O���t�B�b�N", "3D ���f��", "�����N���ꂽ 3D ���f��", "���̑�")

    pshapetypename = Array("-", _
        "�l�p�`", "���s�l�ӌ`", "��`", "�Ђ��`", "�p�ێl�p�`", "���p�`", "�񓙕ӎO�p�`", "���p�O�p�`", "�ȉ~", "�Z�p�`", "�\���`", "�܊p�`", "�~��", "������", _
        "�l�p�`�F�p�x�t��", "�l�p�`�F����", "�X�}�C��", "�~�F�h��Ԃ��Ȃ�", "�֎~�}�[�N", "�A�[�`", "�n�[�g", "���", "���z", "��", "�~��", "�傩����", "��������", _
        "�u���[�`", "���傩����", "�E�傩����", "����������", "�E��������", "���F�E", "���F��", "���F��", "���F��", "���F���E", "���F�㉺", _
        "���F�l����", "���F�O����", "���G�ܐ�", "���FU�^�[��", "���F�����", "���F������ܐ�", "���F�E�J�[�u", "���F���J�[�u", "���F��J�[�u", _
        "���F���J�[�u", "���F�X�g���C�v", "���FV���^", "���F�ܕ���", "���F�R�`", "�����o���F�E���", "�����o���F�����", "�����o���F����", _
        "�����o���F�����", "�����o���F���E���", "�����o���F�㉺���", "�����o���F�l�������", "���F��", "�t���[�`���[�g�F����", "�t���[�`���[�g�F��֏���", _
        "�t���[�`���[�g�F���f", "�t���[�`���[�g�F�f�[�^", "�t���[�`���[�g�F��`�ςݏ���", "�t���[�`���[�g�F�����L��", "�t���[�`���[�g�F����", _
        "�t���[�`���[�g�F��������", "�t���[�`���[�g�F�[�q", "�t���[�`���[�g�F����", "�t���[�`���[�g�F�葀�����", "�t���[�`���[�g�F����", _
        "�t���[�`���[�g�F�����q", "�t���[�`���[�g�F���y�[�W�����q", "�t���[�`���[�g�F�J�[�h", "�t���[�`���[�g�F����E�e�[�v", "�t���[�`���[�g�F�a�ڍ�", _
        "�t���[�`���[�g�F�_���a", "�t���[�`���[�g�F�ƍ�", "�t���[�`���[�g�F����", "�t���[�`���[�g�F�����o��", "�t���[�`���[�g�F�g�ݍ��킹", _
        "�t���[�`���[�g�F�L���f�[�^", "�t���[�`���[�g�F�_���σQ�[�g", "�t���[�`���[�g�F�����A�N�Z�X�L��", "�t���[�`���[�g�F���C�f�B�X�N", _
        "�t���[�`���[�g�F���ڃA�N�Z�X�L��", "�t���[�`���[�g�F�\��", "���� 8pt", "���� 14pt", "�� 4pt", "�� 5pt", "�� 8pt", "�� 16pt", "�� 24pt", "�� 32pt", _
        "���{���F��ɋȂ���", "���{���F���ɋȂ���", "���{���F�J�[�u���ď�ɋȂ���", "���{���F�J�[�u���ĉ��ɋȂ���", "�X�N���[���F�c", "�X�N���[���F��", "�g��", _
        "���g", "�����o���F�l�p�`", "�����o���F�p�ێl�p�`", "�����o���F�~�`", "�v�l�����o���F�_�`", "�����o���F��", "�����o���F��", "�����o���F�ܐ�", _
        "�����o���F�Q�̐ܐ�", "�����o���F��(�������t��)", "�����o���F��(�������t��)", "�����o���F�ܐ�(�������t��)", "�����o���F�Q�̐ܐ�(�������t��)", _
        "�����o���F��(�g�Ȃ�)", "�����o���F��(�g�Ȃ�)", "�����o���F�ܐ�(�g�Ȃ�)", "�����o���F�Q�̐ܐ�(�g�Ȃ�)", "�����o���F��(�g�t���A�������t��)", _
        "�����o���F��(�g�t���A�������t��)", "�����o���F�ܐ�(�g�t���A�������t��)", "�����o���F�Q�̐ܐ�(�g�t���A�������t��)", _
        "�{�^��", "[�z�[��] �{�^��", "[�w���v] �{�^��", "[���] �{�^��", "[�߂�] �܂��� [�O��] �{�^��", "[�i��] �܂��� [����] �{�^��", "[�J�n] �{�^��", _
        "[�I��] �{�^��", "[�߂�] �{�^��", "[����] �{�^��", "[�T�E���h] �{�^��", "[�r�f�I] �{�^��", "�����o��", "���T�|�[�g", "�t���[�`���[�g�F�I�t���C���L��", _
        "���{���F���[���", "�΂ߎ�", "�����~", "��`�F��Ώ�", "�\�p�`", "���p�`", "�\��p�`", "�� 6pt", "�� 7pt", "�� 10pt", "�� 12pt", "�l�p�`�F1���ۂ߂�", _
        "�l�p�`�F���2���ۂ߂�", "�l�p�`�F1��؂���1���ۂ߂�", "�l�p�`�F1��؂���", "�l�p�`�F���2��؂���", "�l�p�`�F�Ίp���ۂ߂�", _
        "�l�p�`�F�Ίp��؂���", "�t���[��", "�t���[��(����)", "�܌^", "��", "L��", "���Z�L��", "���Z�L��", "��Z�L��", "���Z�L��", "���̒l�Ɠ�����", "�����ے�", _
        "�l���F�O�p�`", "�l���F�l�p�`", "�l���F�l���~", "�M�A 6pt", "�M�A 9pt", "�R�l", "�l���~", "���F�����v���", "���F���������v���", "���F�Ȑ�", "�_", _
        "�����`�F�Ίp", "�����`�F�����E�Ίp", "�����`�F�����E����", "�ΐ�", "���̑�")
End Sub

'----------------------------------------
'�}�ʃV�[�g
'----------------------------------------

'�}�ʃV�[�g��ǉ�
Public Sub AddDrawingSheet()
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=ActiveSheet)
    ws.Cells.ColumnWidth = 2.5
End Sub

'�z���}�`
Public Sub DrawLineToLine()
    Dim ce As Range, ws As Worksheet
    Set ce = ActiveCell
    Set ws = ce.Worksheet
    Call DrawLineToLine_1(ce, 1)
    'Call DrawLineToLine_1(ce, 2)
End Sub

Private Sub DrawLineToLine_1(ByVal ce As Range, mode As Long)
    Dim ws As Worksheet
    Set ws = ce.Worksheet
    
    Dim sh As Shape
    Dim fb As FreeformBuilder
    
    Dim ce2 As Range
    Set ce2 = ce
    Do While Left(ce2.Value, 1) = "#"
        Set ce2 = ce2.Offset(-1)
        Do While ce2.Row > 1 And ce2.Value = ""
            Set ce2 = ce2.Offset(-1)
        Loop
    Loop
    
    Dim py(1 To 9) As Double
    py(1) = ce2.Top
    py(2) = ce2.Offset(1).Top
    py(3) = ce2.Offset(2).Top
    py(4) = ce2.Offset(3).Top
    py(5) = ce2.Offset(4).Top
    py(6) = ce2.Offset(5).Top
    py(7) = ce2.Offset(6).Top
    py(8) = ce2.Offset(7).Top
    py(9) = ce2.Offset(8).Top
    Dim yn As Long
    yn = 1
    
    Dim re As Object
    Set re = regex("[!#$%&()<>\[\]]")
    
    Dim y As Double, y0 As Double

    Dim x1 As Double, x2 As Double, dx As Double, x As Double
    Dim ss As String, s0 As String, s1 As String
    Dim i As Long, m As Long
    
    ss = re.Replace(CStr(ce.Value), "")
    m = Len(re.Replace(ss, ""))
    Do While m > 0
        x = ce.Left
        If m < 1 Then m = 1
        dx = (ce.Offset(, 1).Left - x) / m
        For i = 1 To Len(ss)
            s0 = s1: s1 = Mid(ss, i, 1)
            Select Case s1
            Case "1": yn = 1: y = py(yn)
            Case "2": yn = 2: y = py(yn)
            Case "3": yn = 3: y = py(yn)
            Case "4": yn = 4: y = py(yn)
            Case "5": yn = 5: y = py(yn)
            Case "6": yn = 6: y = py(yn)
            Case "7": yn = 7: y = py(yn)
            Case "8": yn = 8: y = py(yn)
            Case "9": yn = 9: y = py(yn)
            Case "/": yn = IIf(yn > 1, yn - 1, yn): y = py(yn): y0 = y
            Case "\": yn = IIf(yn < 9, yn + 1, yn): y = py(yn): y0 = y
            End Select
            
            If fb Is Nothing Then
                If y <> 0 Then Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, x, y)
            ElseIf y0 <> y Then
                fb.AddNodes msoSegmentLine, msoEditingAuto, x, y
            End If
            x = x + dx
            If Not fb Is Nothing And s1 <> "-" Then
                fb.AddNodes msoSegmentLine, msoEditingAuto, x, y
                y0 = y
            End If
        Next i
        Set ce = ce.Offset(, 1)
        ss = CStr(ce.Value)
        m = Len(re.Replace(ss, ""))
    Loop
    
    If Not fb Is Nothing And s1 = "-" Then
        fb.AddNodes msoSegmentLine, msoEditingAuto, x, y
    End If
    If Not fb Is Nothing Then Set sh = fb.ConvertToShape
    Set fb = Nothing
End Sub

