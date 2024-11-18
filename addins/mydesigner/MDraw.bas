Attribute VB_Name = "MDraw"
'==================================
'�`��
'==================================

Option Explicit
Option Private Module

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
        & ";Scale,�X�P�[��,2    ;X0,���_X,2     ;Y0,���_Y,2     ;Z0,���_Z,2"

'�}�`���X�g�w�b�_
Private Const c_ShapeInfoHeader As String = "" _
    & ";����,           Name,Title" _
    & ";�`��,           Name,ID,Type,Style,Title" _
    & ";�ʒu,           Name,Top,Left,Back,Rotation" _
    & ";�T�C�Y,         Name,Height,Width,Depth" _
    & ";�\��,           Name,Visible,Transparency" _
    & ";�g��,           Name,LineVisible,LineColor" _
    & ";�h��,           Name,FillVisible,FillColor" _
    & ";�e�L�X�g,       Name,Text" _
    & ";��ւ��e�L�X�g, Name,AltText" _
    & ";����,           Name,Scale,X0,Y0,Z0"

'���p�����[�^����
Public Enum E_DrawParam
    E_IGNORE = 1
    E_SCALE = 2
    E_AXES = 3
    E_FLAG = 4
    E_PART = 10
End Enum

'���p�����[�^
Private g_mask As String            '������������
Private g_scale As Double           '�X�P�[��
Private g_axes As Double            '���Ԋu
Private g_flag As Integer           '���[�h(0:,1:,2:,3:)
Private g_part As String            '���i��

Private ptype_col As Variant        '�}�`�^�C�v�e�[�u��
Private ptypename As Variant        '�}�`�^�C�v���̃e�[�u��
Private pshapetypename As Variant   '�}�`�^�C�v�e�[�u��

'----------------------------------------
'�p�����[�^����
'----------------------------------------

'�`��p�����[�^������
Public Sub ResetDrawParam(Optional id As Integer)
    If id = 0 Or id = 1 Then g_mask = ""
    If id = 0 Or id = 2 Then g_scale = 0.1
    If id = 0 Or id = 3 Then g_axes = 10
    If id = 0 Or id = 4 Then g_flag = 0
    If id = 0 Or id = 10 Then g_part = ""
End Sub

'�`��p�����[�^�ݒ�
Public Sub SetDrawParam(id As Integer, ByVal val As String)
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
    Case 4: g_flag = (g_flag And (65535 - 1)) Or (val * 1)
    Case 5: g_flag = (g_flag And (65535 - 2)) Or (val * 2)
    Case 6: g_flag = (g_flag And (65535 - 4)) Or (val * 4)
    Case 7: g_flag = (g_flag And (65535 - 8)) Or (val * 8)
    Case 8: g_flag = (g_flag And (65535 - 16)) Or (val * 16)
    Case 9: g_flag = (g_flag And (65535 - 32)) Or (val * 32)
    Case 10: g_part = val
    End Select
End Sub

'�`��p�����[�^�擾
Public Function GetDrawParam(id As Integer) As String
    Select Case id
    Case 1: GetDrawParam = g_mask
    Case 2
        If g_scale <= 0 Then
            ResetDrawParam id
            MsgBox "�䗦�̐ݒ�����������܂����B(�ݒ�l" & g_scale & ")"
        End If
        GetDrawParam = g_scale
    Case 3
        If g_axes <= 0 Then
            MsgBox "�ڐ���̐ݒ�����������܂����B(�ݒ�l" & g_scale & ")"
            ResetDrawParam id
        End If
        GetDrawParam = g_axes
    End Select
End Function

'�`��p�����[�^�t���O�`�F�b�N
Public Function IsDrawParam(id As Integer) As Boolean
    IsDrawParam = ((g_flag And (65535 - (2 ^ (id - 4)))) <> 0)
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

'----------------------------------------
'�}�`���i�`��
'----------------------------------------

'���i�`��
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

'�A�C�e�����
Public Function DrawGraphItem(id As Integer, Optional ra As Range) As String
    If ra Is Nothing Then Exit Function
    '��ʃ`�����h�~���u
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Interactive = False
    Application.Cursor = xlWait
    '
    Select Case id
    Case 1
        '���ᎆ�`��
        DrawGraphItem = DrawGraph2(ra.Worksheet, ra.Left, ra.Top + ra.Height, ra.Width, ra.Height)
    Case 2
        '���`��
        DrawGraphItem = DrawAxis2(ra.Worksheet, ra.Left, ra.Top + ra.Height, ra.Width, ra.Height, 10)
    Case 3
        '���_�`��
        DrawGraphItem = DrawAxis2(ra.Worksheet, ra.Left, ra.Top + ra.Height, ra.Width, ra.Height, 10)
    Case 4
        '���i�`��
        DrawGraphItem = DrawParts(ra.Worksheet, ra.Left, ra.Top + ra.Height, g_part)
    End Select
    If Not DrawGraphItem = "" Then ra.Worksheet.Shapes(DrawGraphItem).Select
    '
    '��ʃ`�����h�~���u����
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Function

'�������
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
    sh.name = "���� " & sh.id
    DrawAxis2 = sh.name
End Function

'���ᎆ�쐬
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
        MsgBox ("�Ԋu���������܂��B(" & dp & ")")
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
    sh.name = "���ᎆ " & sh.id
    DrawGraph2 = sh.name
End Function

'----------------------------------------
'�}�`��������
'----------------------------------------

'�}�`��񃊃X�g�A�b�v
Public Sub ListShapeInfo(ByVal ws As Worksheet, Optional mode As Integer)
    
    '�o�͐�Z��/�}�`���X�g�擾
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
        Set ce = GetCell("���X�g�o�͈ʒu���w�肵�Ă�������", "�}�`���X�g�o��")
        Set sr = Selection.ShapeRange
    End If
    If ce Is Nothing Or sr Is Nothing Then Exit Sub
    If ce.Value = "" Then
        ce.Value = sr.Parent.name & "[" & sr.Parent.Parent.name & "]"
        Set ce = ce.Offset(1)
    End If
    
    '�e�[�u�����ڎ擾
    Dim dic As Dictionary
    ArrStrToDict dic, c_ShapeInfoMember, 1
    
    '�w�b�_�擾
    Dim hdr() As String
    StringToRow hdr, c_ShapeInfoHeader, mode
    If mode < 1 And ce.Value <> "" Then
        hdr = GetHeaderArray(ce, dic)
    End If
    
    '�f�[�^�z��쐬
    Dim rcnt As Long
    rcnt = 1 + sr.Count
    Dim sh As Shape
    For Each sh In sr
        If sh.Type = msoGroup Then rcnt = rcnt + sh.GroupItems.Count
    Next sh
    Dim arr As Variant
    ReDim arr(rcnt, UBound(hdr))
    
    '�w�b�_�s�ݒ�
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
    
    '�}�`���}�X�N�ݒ�
    Dim ptn As String
    Dim flg As Boolean
    ptn = g_mask
    flg = True
    If Left(ptn, 1) = "!" Then
        ptn = Mid(ptn, 2)
        flg = False
    End If
    If ptn = "" Then ptn = ".*"
    
    '���R�[�h�쐬
    For Each sh In sr
        AddShapeRecord arr, r, sh, hdr, ptn, flg
    Next sh
    rcnt = r
    
    '��ʍX�V��~
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    '�e�[�u���f�[�^�N���A
    TableDataRange(ce).Clear
    
    '�\���`���ݒ�
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
    
    '���R�[�h��������
    ce.Resize(1 + rcnt, 1 + UBound(hdr)).Value = arr
    
    '�z�F
    For c = 0 To UBound(hdr)
        Set ra = ce.Parent.Range(ce.Cells(2, c + 1), ce.Cells(rcnt, c + 1))
        s = UCase(hdr(c))
        If dic.Exists(s) Then
            Select Case CInt(dic(s)(2))
            Case 4  '�F
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
    
    '�e�[�u���T�C�Y����
    HeaderAutoFit ce
    
    '��ʍX�V�ĊJ
    Application.ScreenUpdating = fsu

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

'�}�`���ݒ�
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

'�}�`��񃊃X�g�̔��f
Public Sub UpdateShapeInfo(ByVal ra As Range, Optional ByVal ws As Worksheet)
    
    If Not TypeName(Selection) = "Range" Then Exit Sub
    If ws Is Nothing Then Set ws = ActiveSheet
    If ra Is Nothing Then Set ra = ActiveCell
    
    '�e�[�u���J�n�ʒu���擾
    Dim ce As Range
    Set ce = FarLeftTop(ra.Cells(2, 1))
    
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
        dic.Add sh.name, sh
        If sh.Type = msoGroup Then
            For Each sh2 In sh.GroupItems
                dic.Add sh2.name, sh2
            Next sh2
        End If
    Next sh
    
    '��ʍX�V��~
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
    '��ʍX�V�ĊJ
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

'�}�ʃV�[�g��ǉ�
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


