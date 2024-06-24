Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI

'----------------------------------------
'ribbon helper
'----------------------------------------

'ID�ԍ��擾
Private Function RB_ID(control As IRibbonControl) As Integer
    RB_ID = Val(Right(control.id, 1))
End Function

'TAG�ԍ��擾
Private Function RB_TAG(control As IRibbonControl) As Integer
    RB_TAG = Val(control.Tag)
End Function

'���{��ID�ԍ��擾
Private Function RibbonID(control As IRibbonControl) As Integer
    Dim s As String
    s = control.Tag
    If s = "" Then s = control.id
    Dim vs As Variant
    vs = Split(s, ".")
    If UBound(vs) >= 0 Then
        RibbonID = Val(vs(UBound(vs)))
        Exit Function
    End If
    vs = Split(s, "_")
    If UBound(vs) >= 0 Then
        RibbonID = Val(vs(UBound(vs)))
        Exit Function
    End If
    RibbonID = Val(s)
End Function

'���{�����X�V
Private Sub RefreshRibbon()
    If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
    DoEvents
End Sub

'----------------------------------------
'Initialize
'----------------------------------------

Private Sub works_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    '
    '�V���[�g�J�b�g�L�[�ݒ�
    Application.OnKey "{F1}"
    Application.OnKey "{F1}", "works_ShortcutKey1"
    '
    Application.OnKey "+{F1}"
    Application.OnKey "+{F1}", "works_ShortcutKey2"
    '
    '������
    SetParam "path", 1, True        '�����N����
    SetParam "path", 2, True        '�t�H���_����
    SetParam "path", 3, True        '�ċA����
    SetParam "info", 1, True        '�V�[�g�ǉ�
    SetParam "mark", "color", 10    '�}�[�J�J���[�͉��F
End Sub

Private Sub works_ShortcutKey1()
    'works�^�u�Ɉړ�
    If Not g_ribbon Is Nothing Then g_ribbon.ActivateTab "TabWorks"
End Sub

Private Sub works_ShortcutKey2()
    'home�^�u�Ɉړ�
    SendKeys "%"
    SendKeys "H"
    SendKeys "%"
End Sub

'----------------------------------------
'1x:���|�[�g�@�\
'----------------------------------------

'���|�[�g�T�C��
Private Sub works11_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call ReportSign(Selection)
End Sub

'�y�[�W�t�H�[�}�b�g
Private Sub works12_onAction(ByVal control As IRibbonControl)
    Call PagePreview
End Sub

'�e�L�X�g�ϊ�
Private Sub works13_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call MenuTextConv(RibbonID(control), Selection)
End Sub

'�g������
Private Sub works14_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call MenuUserFormat(RibbonID(control), Selection)
End Sub

'��^���}��
Private Sub works15_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call MenuUserFormula(RibbonID(control), Selection)
End Sub

'�\���E��\��
Private Sub works16_onAction(ByVal control As IRibbonControl)
    Call ShowHide(RibbonID(control))
End Sub

'�p�X��
Private Sub works17_onAction(ByVal control As IRibbonControl)
    Call PathMenu(RibbonID(control), Selection)
End Sub

Private Sub works17_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetParam "path", RibbonID(control), pressed
End Sub

Private Sub works17_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetParamBool("path", RibbonID(control))
End Sub

'���擾
Private Sub works18_onAction(ByVal control As IRibbonControl)
    Call AddInfoSheet(RibbonID(control))
End Sub

Private Sub works18_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetParam "info", RibbonID(control), pressed
End Sub

Private Sub works18_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetParamBool("info", RibbonID(control))
End Sub

'�G�N�X�|�[�g
Private Sub works19_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call MenuExport(Selection, RibbonID(control))
End Sub

Private Sub works19_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetParam "export", RibbonID(control), pressed
End Sub

Private Sub works19_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetParamBool("export", RibbonID(control))
End Sub

'----------------------------------------
'2x:�r���g
'----------------------------------------

'�ړ��E�I��
Private Sub works21_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    SelectTable RibbonID(control), Selection
    Application.ScreenUpdating = True
End Sub

'�g�ݒ�
Private Sub works22_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Call TableWaku(RibbonID(control), Selection)
    Application.ScreenUpdating = True
End Sub

'�͂��N���A
Private Sub works27_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Select Case RibbonID(control)
    Case 1
        '�͂��N���A
        Call TableWaku(7, Selection)
    Case 2
        '�f�[�^�N���A
        Call TableWaku(8, Selection)
    Case 3
        '�\�N���A
        Call TableWaku(9, Selection)
    Case Else
        '�͂��E�f�[�^�N���A
        Call TableWaku(7, Selection)
        Call TableWaku(9, Selection)
    End Select
    Application.ScreenUpdating = True
End Sub

'�}�[�W���\���E�ݒ�
Private Sub works28_onAction(ByVal control As IRibbonControl)
    SetTableMargin xlRows
    SetTableMargin xlColumns
    g_ribbon.InvalidateControl control.id
End Sub

Private Sub works28_getLabel(ByRef control As Office.IRibbonControl, ByRef label As Variant)
   label = "�s: " & GetTableMargin(xlRows) & ", ��: " & GetTableMargin(xlColumns)
End Sub

'----------------------------------------
'3x:�e���v���[�g�@�\
'----------------------------------------

Private Sub works3_onAction(ByVal control As IRibbonControl)
    Call TemplateMenu(RibbonID(control))
    Select Case RibbonID(control)
    'Case 8 '�X�V
    '    RBTable_Init
    '    g_ribbon.InvalidateControl "b4.1"
    '    g_ribbon.InvalidateControl "b4.2"
    '    g_ribbon.InvalidateControl "b4.3"
    '    g_ribbon.InvalidateControl "b4.4"
    '    g_ribbon.InvalidateControl "b4.5"
    '    g_ribbon.InvalidateControl "b4.6"
    '    g_ribbon.InvalidateControl "b4.7"
    '    g_ribbon.InvalidateControl "b4.8"
    '    g_ribbon.InvalidateControl "b4.9"
    Case 9 '�J��
        g_ribbon.InvalidateControl "b3.9"
    End Select
End Sub

Private Sub works3_getLabel(ByRef control As IRibbonControl, ByRef label As Variant)
    If ThisWorkbook.IsAddin Then label = "�u�b�N�J��" Else label = "�u�b�N����"
    g_ribbon.Invalidate
    DoEvents
End Sub

Private Sub works3_getEnabled(control As IRibbonControl, ByRef enable As Variant)
    enable = ThisWorkbook.IsAddin
End Sub

'----------------------------------------
'4:marker
'----------------------------------------

Private Sub works4_getLabel(control As IRibbonControl, ByRef label As Variant)
    Dim name() As Variant
    name = Array("-", "��", "��", "��", "�D�F", "��", "��", "�W����", "��", "��", "���F")
    label = name(Val(GetParam("mark", "color")))
End Sub

Private Sub works4_onGetImage(control As IRibbonControl, ByRef bitmap As Variant)
    bitmap = "AppointmentColor" & Val(GetParam("mark", "color"))
End Sub

Private Sub works4_onAction(control As IRibbonControl)
    Select Case RibbonID(control)
    Case 1
        If TypeName(Selection) <> "Range" Then Exit Sub
        Call AddMarker(Selection, Val(GetParam("mark", "color")))
    Case 2
        If TypeName(Selection) <> "Range" Then Exit Sub
        Call ListMarker(Selection)
    Case 3
        If TypeName(Selection) <> "Range" Then Exit Sub
        Dim ce As Range
        For Each ce In Selection.Cells
            Call DelMarker(ce.Value)
            ce.Clear
        Next ce
    Case 4
        Call DelMarkerAll
    End Select
End Sub

Private Sub works41_onAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    Call SetParam("mark", "color", Mid(selectedId, InStr(1, selectedId, ".") + 1))
    If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
    DoEvents
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call AddMarker(Selection, Val(GetParam("mark", "color")))
End Sub

'----------------------------------------
'5:revision mark
'----------------------------------------

Private Sub works5_getLabel(control As IRibbonControl, ByRef label As Variant)
    Call GetRevMark(label)
End Sub

Private Sub works5_onAction(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    '
    Dim rev As String
    Select Case RibbonID(control)
    Case 1
        '�Ő��}�[�N�ǉ�
        Call AddRevMark(Selection)
    Case 2
        '�Ő��ݒ�
        Call GetRevMark(rev)
        rev = InputBox("�Ő�����͂��Ă��������B", "�Ő��}�[�N�ݒ�", rev)
        If rev = "" Then Exit Sub
        Call SetRevMark(rev)
        If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
        DoEvents
        Call AddRevMark(Selection)
    Case 3
        '�Ő����X�g�쐬
        Call GetRevMark(rev)
        rev = InputBox("�Ő�����͂��Ă��������B", "�Ő��}�[�N���X�g", rev)
        If rev = "" Then Exit Sub
        Call ListRevMark(Selection, rev)
    End Select
End Sub

