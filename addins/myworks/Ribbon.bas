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

'���{�����X�V
Private Sub RefreshRibbon(Optional id As String)
    If Not g_ribbon Is Nothing Then
        If id = "" Then
            g_ribbon.Invalidate
        Else
            g_ribbon.InvalidateControl id
        End If
    End If
    DoEvents
End Sub

'���{��ID�ԍ��擾
Private Function RibbonID(control As IRibbonControl, Optional n As Long) As Long
    Dim vs As Variant
    vs = Split(re_replace(control.id, "[^0-9.]", ""), ".")
    If UBound(vs) >= n Then RibbonID = Val("0" & vs(UBound(vs) - n))
End Function

'----------------------------------------
'�C�x���g
'----------------------------------------

'�N�������s
Private Sub works_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    '
    '�V���[�g�J�b�g�L�[�ݒ�
    On Error Resume Next
    Application.OnKey "{F1}"
    Application.OnKey "{F1}", "works_ShortcutKey1"
    '
    Application.OnKey "+{F1}"
    Application.OnKey "+{F1}", "works_ShortcutKey2"
    On Error GoTo 0
    '
    '������
    SetRtParam "path", 1, True          '�����N����
    SetRtParam "path", 2, True          '�t�H���_����
    SetRtParam "path", 3, True          '�ċA����
    SetRtParam "info", "sheet", True    '�V�[�g�ǉ�
    SetRtParam "mark", "color", 0       '�}�[�J�J���[�͉��F
End Sub

Private Sub works_ShortcutKey1()
    'works�^�u�Ɉړ�
    If g_ribbon Is Nothing Then Exit Sub
    g_ribbon.ActivateTab "TabWorks"
End Sub

Private Sub works_ShortcutKey2()
    'home�^�u�Ɉړ�
    SendKeys "%"
    SendKeys "H"
    SendKeys "%"
End Sub

'----------------------------------------
'���@�\�O���[�v1
'���|�[�g�@�\
'----------------------------------------

'���|�[�g�T�C��
Private Sub works11_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    ReportSign Selection
End Sub

'�y�[�W�t�H�[�}�b�g
Private Sub works12_onAction(ByVal control As IRibbonControl)
    PagePreview
End Sub

'�e�L�X�g�ϊ�
Private Sub works13_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    MenuTextConv Selection, RibbonID(control)
End Sub

'�g������
Private Sub works14_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    MenuUserFormat Selection, RibbonID(control)
End Sub

'��^���}��
Private Sub works15_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    MenuUserFormula Selection, RibbonID(control)
End Sub

'�\���E��\��
Private Sub works16_onAction(ByVal control As IRibbonControl)
    ShowHide RibbonID(control)
End Sub

'�p�X��
Private Sub works17_onAction(ByVal control As IRibbonControl)
    PathMenu Selection, RibbonID(control)
End Sub

Private Sub works17_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetRtParam "path", RibbonID(control), CStr(pressed)
End Sub

Private Sub works17_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetRtParamBool("path", RibbonID(control))
End Sub

'���擾
Private Sub works18_onAction(ByVal control As IRibbonControl)
    AddInfoTable RibbonID(control)
End Sub

Private Sub works18_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetRtParam "info", control.Tag, CStr(pressed)
End Sub

Private Sub works18_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetRtParamBool("info", control.Tag)
End Sub

'�G�N�X�|�[�g
Private Sub works19_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    MenuExport Selection, RibbonID(control)
End Sub

Private Sub works19_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    SetRtParam "export", RibbonID(control), CStr(pressed)
End Sub

Private Sub works19_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetRtParamBool("export", RibbonID(control))
End Sub

'----------------------------------------
''���@�\�O���[�v2
'�r���g
'----------------------------------------

'�ړ��E�I��
Private Sub works21_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    TableSelect Selection, RibbonID(control)
End Sub

'�g�ݒ�
Private Sub works22_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Call TableWaku(Selection, RibbonID(control))
    Application.ScreenUpdating = True
End Sub

'��ǉ�
Private Sub works23_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Call AddColumn(Selection, RibbonID(control))
    Application.ScreenUpdating = True
End Sub

'�͂��N���A
Private Sub works27_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Select Case RibbonID(control)
    Case 1
        '�͂��N���A
        Call TableWaku(Selection, 7)
    Case 2
        '�f�[�^�N���A
        Call TableWaku(Selection, 8)
    Case 3
        '�\�N���A
        Call TableWaku(Selection, 9)
    Case Else
        '�͂��E�f�[�^�N���A
        Call TableWaku(Selection, 7)
        Call TableWaku(Selection, 9)
    End Select
    Application.ScreenUpdating = True
End Sub

'�}�[�W���\���E�ݒ�
Private Sub works28_onAction(ByVal control As IRibbonControl)
    SetTableMargin xlRows
    SetTableMargin xlColumns
    RefreshRibbon control.id
End Sub

Private Sub works28_getLabel(ByRef control As Office.IRibbonControl, ByRef label As Variant)
   label = "�s: " & GetTableMargin(xlRows) & ", ��: " & GetTableMargin(xlColumns)
End Sub

'----------------------------------------
'���@�\�O���[�v3
'�e���v���[�g�@�\
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
'���@�\�O���[�v4
'marker
'----------------------------------------

Private Sub works4_getLabel(control As IRibbonControl, ByRef label As Variant)
    Dim name() As String
    name = Split("���F,��,��,����,�D�F,��,��,�W����,��,��", ",")
    label = name(Val(GetRtParam("mark", "color")))
End Sub

Private Sub works4_onGetImage(control As IRibbonControl, ByRef bitmap As Variant)
    Dim id As Integer
    id = ((Val(GetRtParam("mark", "color")) + 9) Mod 10) + 1
    bitmap = "AppointmentColor" & id
End Sub

Private Sub works4_onAction(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Select Case RibbonID(control)
    Case 1
        Call AddMarker(Selection, Val(GetRtParam("mark", "color")))
    Case 2
        Call ListMarker(Selection)
    Case 3
        ScreenUpdateOff
        Dim ce As Range
        For Each ce In Selection.Cells
            Call DelMarker(ce.Value)
            ce.Clear
        Next ce
        ScreenUpdateOn
    End Select
End Sub

Private Sub works41_onAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    Call SetRtParam("mark", "color", Mid(selectedId, 1 + InStr(1, selectedId, ".")))
    If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
    DoEvents
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call AddMarker(Selection, Val(GetRtParam("mark", "color")))
End Sub

'----------------------------------------
'���@�\�O���[�v5
'revision mark
'----------------------------------------

Private Sub works5_getLabel(control As IRibbonControl, ByRef label As Variant)
    Call GetRevMark(label)
End Sub

Private Sub works5_onAction(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim res As Long
    Call RevProc(Selection, RibbonID(control), res)
    If Not res Then Exit Sub
    RefreshRibbon
    Call RevProc(Selection, 1)
End Sub

