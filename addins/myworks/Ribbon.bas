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

'�R�}���hID�ԍ��擾
Private Function RB_CID(control As IRibbonControl) As Integer
    RB_CID = val(Right(control.id, 1))
End Function

'TAG�ԍ��擾
Private Function RB_TAG(control As IRibbonControl) As Integer
    RB_TAG = val(control.Tag)
End Function

'ID�ԍ��擾
Private Function RB_ID(control As IRibbonControl) As Integer
    If control.Tag = "" Then
        RB_ID = val(Right(control.id, 1))
    Else
        RB_ID = val(control.Tag)
    End If
End Function

'���{��ID�ԍ��擾
Private Function RibbonID(control As IRibbonControl) As Integer
    Dim s As String
    s = control.Tag
    If s = "" Then s = control.id
    Dim vs As Variant
    vs = Split(s, ".")
    If UBound(vs) >= 0 Then
        RibbonID = val(vs(UBound(vs)))
        Exit Function
    End If
    vs = Split(s, "_")
    If UBound(vs) >= 0 Then
        RibbonID = val(vs(UBound(vs)))
        Exit Function
    End If
    RibbonID = val(s)
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
    Application.OnKey "{F1}", "works_ShortcutKey"
    '
    RBTable_Init
    SetDataListParam 1, True
End Sub

Private Sub works_ShortcutKey()
    g_ribbon.ActivateTab "TabWorks"
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
    Call Cells_Conv(Selection, RB_ID(control))
End Sub

'�\���E��\��
Private Sub works14_onAction(ByVal control As IRibbonControl)
    Call ShowHide(RibbonID(control))
End Sub

'�ڎ��V�[�g�쐬
Private Sub works15_onAction(ByVal control As IRibbonControl)
    Call AddInfoSheet(RibbonID(control))
End Sub

Private Sub works15_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Call SetDataListParam(RibbonID(control), pressed)
End Sub

Private Sub works15_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetDataListParam(RibbonID(control))
End Sub

'�p�X��
Private Sub works16_onAction(ByVal control As IRibbonControl)
    Call GetPath(Selection, RibbonID(control))
End Sub

Private Sub works16_onChecked(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Call SetPathParam(RibbonID(control), pressed)
End Sub

Private Sub works16_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetPathParam(RibbonID(control))
End Sub

'�g�������t������
Private Sub works17_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call AddUserFormat(Selection, RibbonID(control))
End Sub

'��^���}��
Private Sub works18_onAction(ByVal control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Call InsertFormula(Selection, RibbonID(control))
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
        '�e�[�u���N���A
        Call TableWaku(8, Selection)
    Case Else
        '�\�N���A
        Call TableWaku(7, Selection)
        Call TableWaku(8, Selection)
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
    Call TemplateMenu(RB_ID(control))
    Select Case RB_ID(control)
    Case 8 '�X�V
        RBTable_Init
        g_ribbon.InvalidateControl "b4.1"
        g_ribbon.InvalidateControl "b4.2"
        g_ribbon.InvalidateControl "b4.3"
        g_ribbon.InvalidateControl "b4.4"
        g_ribbon.InvalidateControl "b4.5"
        g_ribbon.InvalidateControl "b4.6"
        g_ribbon.InvalidateControl "b4.7"
        g_ribbon.InvalidateControl "b4.8"
        g_ribbon.InvalidateControl "b4.9"
    Case 9 '�J��
        g_ribbon.InvalidateControl "b39"
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
'common
'----------------------------------------

Private Sub works4_onAction(control As IRibbonControl)
    Call RBTable_onAction(RibbonID(control))
End Sub

Private Sub works4_getVisible(control As IRibbonControl, ByRef Visible As Variant)
    Call RBTable_getVisible(RibbonID(control), Visible)
End Sub

Private Sub works4_getLabel(control As IRibbonControl, ByRef label As Variant)
    Call RBTable_getLabel(RibbonID(control), label)
End Sub

Private Sub works4_onGetImage(control As IRibbonControl, ByRef bitmap As Variant)
    Call RBTable_onGetImage(RibbonID(control), bitmap)
End Sub

Private Sub works4_getSize(control As IRibbonControl, ByRef Size As Variant)
    Call RBTable_getSize(RibbonID(control), Size)
End Sub

