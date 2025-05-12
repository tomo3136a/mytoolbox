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
    If UBound(vs) >= n Then RibbonID = Val("0" & vs(n))
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
    SetRtParam "path.1", True       '�����N����
    SetRtParam "path.2", True       '�t�H���_����
    SetRtParam "path.3", True       '�ċA����
    SetRtParam "info.sheet", True   '�V�[�g�ǉ�
    SetRtParam "mark.color", 0      '�}�[�J�J���[�͉��F
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
Private Sub works1_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 1:
        '���|�[�g�T�C��
        If TypeName(Selection) <> "Range" Then Exit Sub
        ReportSign Selection
    Case 2:
        '�y�[�W�t�H�[�}�b�g
        PagePreview
    Case 3:
        '�e�L�X�g�ϊ�
        If TypeName(Selection) <> "Range" Then Exit Sub
        TextConvProc Selection, RibbonID(control, 2)
    Case 4:
        '�g������
        If TypeName(Selection) <> "Range" Then Exit Sub
        UserFormatProc Selection, RibbonID(control, 2)
    Case 5:
        '��^���}��
        If TypeName(Selection) <> "Range" Then Exit Sub
        UserFormulaProc Selection, RibbonID(control, 2)
    Case 6:
        '�\���E��\��
        ShowHide RibbonID(control, 2)
    Case 7:
        '�p�X��
        If TypeName(Selection) <> "Range" Then Exit Sub
        PathProc Selection, RibbonID(control, 2)
    Case 8:
        '���擾
        AddInfoTable RibbonID(control, 2)
    Case 9:
        '�G�N�X�|�[�g
        If TypeName(Selection) <> "Range" Then Exit Sub
        ExportProc Selection, RibbonID(control, 2)
    End Select
End Sub

Private Sub works1_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Select Case RibbonID(control, 1)
    Case 7: SetRtParam "path." & RibbonID(control, 2), CStr(pressed)    '�p�X��
    Case 8: SetRtParam "info." & control.Tag, CStr(pressed)             '���擾
    Case 9: SetRtParam "export." & RibbonID(control, 2), CStr(pressed)  '�G�N�X�|�[�g
    End Select
End Sub

Private Sub works1_getPressed(control As IRibbonControl, ByRef returnedVal)
    Select Case RibbonID(control, 1)
    Case 7: returnedVal = GetRtParamBool("path." & RibbonID(control, 2))    '�p�X��
    Case 8: returnedVal = GetRtParamBool("info." & control.Tag)             '���擾
    Case 9: returnedVal = GetRtParamBool("export." & RibbonID(control, 2))  '�G�N�X�|�[�g
    End Select
End Sub

'----------------------------------------
''���@�\�O���[�v2
'�r���g
'----------------------------------------

'�g�ݒ�
Private Sub works2_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    ScreenUpdateOff
    Select Case RibbonID(control, 1)
    Case 1: Call WakuProc(Selection, RibbonID(control, 2))
    Case 2: Call SelectProc(Selection, RibbonID(control, 2))
    Case 3: Call AddColumn(Selection, RibbonID(control, 2))
    Case 7
        Select Case RibbonID(control, 2)
        Case 1: Call WakuProc(Selection, 7)    '�͂��N���A
        Case 2: Call WakuProc(Selection, 8)    '�f�[�^�N���A
        Case 3: Call WakuProc(Selection, 9)    '�\�N���A
        Case Else                               '�͂��E�f�[�^�N���A
            Call WakuProc(Selection, 7)
            Call WakuProc(Selection, 9)
        End Select
    Case 8
        SetTableMargin xlRows
        SetTableMargin xlColumns
        RefreshRibbon control.id
    End Select
    ScreenUpdateOn
End Sub

Private Sub works2_getLabel(ByRef control As Office.IRibbonControl, ByRef label As Variant)
   label = "�s: " & GetTableMargin(xlRows) & ", ��: " & GetTableMargin(xlColumns)
End Sub

'----------------------------------------
'���@�\�O���[�v3
'�e���v���[�g�@�\
'----------------------------------------

Private Sub works3_onAction(ByVal control As IRibbonControl)
    Call TemplateMenu(RibbonID(control, 1))
    Select Case RibbonID(control, 1)
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
    RefreshRibbon
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
    Dim ss() As String
    ss = Split("���F,��,��,����,�D�F,��,��,�W����,��,��", ",")
    label = ss(Val(GetRtParam("mark.color")))
End Sub

Private Sub works4_onGetImage(control As IRibbonControl, ByRef bitmap As Variant)
    Dim id As Integer
    id = ((Val(GetRtParam("mark.color")) + 9) Mod 10) + 1
    bitmap = "AppointmentColor" & id
End Sub

Private Sub works4_onAction(control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 1
        If TypeName(Selection) <> "Range" Then Exit Sub
        Call AddMarker(Selection, Val(GetRtParam("mark.color")))
    Case 3
        Call ListMarker
    Case 4
        If TypeName(Selection) <> "Range" Then Exit Sub
        Call DelMarker(Selection)
    End Select
End Sub

Private Sub works4_onSelected(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    Call SetRtParam("mark.color", "" & selectedIndex)
    RefreshRibbon
    DoEvents
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call AddMarker(Selection, Val(GetRtParam("mark.color")))
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
    Call RevProc(Selection, RibbonID(control, 1), res)
    RefreshRibbon
    If Not res Then Exit Sub
    Call RevProc(Selection, 1)
End Sub

'----------------------------------------
'���@�\�O���[�v6
'test
'----------------------------------------

Private Sub works6_onAction(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    'ScreenUpdateOff
    Select Case RibbonID(control, 1)
    Case 1: Call TestProc(Selection, RibbonID(control, 1))
    Case 2: Call TestProc(Selection, RibbonID(control, 1))
    Case 3: Call TestProc(Selection, RibbonID(control, 1))
    Case 4: Call TestProc(Selection, RibbonID(control, 1))
    Case 5: Call TestProc(Selection, RibbonID(control, 1))
    End Select
    'ScreenUpdateOn
End Sub

Private Sub TestProc(ra As Range, id As Long)
    Select Case id
    Case 1: Call Cells_GenerateValue(ra, 1)
    End Select
End Sub

