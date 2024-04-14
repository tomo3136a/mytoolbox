Attribute VB_Name = "Ribbon"
'========================================
'rebbon interface
'========================================

Option Explicit
Option Private Module

'----------------------------------------

Private g_ribbon As IRibbonUI

'----------------------------------------

Private Sub RB_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    
    '�V���[�g�J�b�g�L�[�ݒ�
    Application.OnKey "{F1}"
    Application.OnKey "{F1}", "RB_ShortcutKey"
    
    RBTable_Init
End Sub

Private Sub RB_ShortcutKey()
    g_ribbon.ActivateTab "TabWorks"
End Sub

'�R�}���hID�ԍ��擾
Private Function RB_CID(control As IRibbonControl) As Integer
    RB_CID = CInt("0" & Right(control.id, 1))
End Function

'TAG�ԍ��擾
Private Function RB_TAG(control As IRibbonControl) As Integer
    RB_TAG = CInt("0" & control.tag)
End Function

'ID�ԍ��擾
Private Function RB_ID(control As IRibbonControl) As Integer
    If control.tag = "" Then
        RB_ID = CInt("0" & Right(control.id, 1))
    Else
        RB_ID = CInt("0" & control.tag)
    End If
End Function

'���{�����X�V
Private Sub RefreshAddInsRibbon()
    g_ribbon.Invalidate
    DoEvents
End Sub


'----------------------------------------
'���|�[�g�@�\
'----------------------------------------

'���|�[�g�T�C��
Private Sub RB11_onAction(ByVal control As IRibbonControl)
    Call ReportSign(Selection)
End Sub

'�y�[�W�t�H�[�}�b�g
Private Sub RB12_onAction(ByVal control As IRibbonControl)
    Call PagePreview
End Sub

'�e�L�X�g�ϊ�
Private Sub RB13_onAction(ByVal control As IRibbonControl)
    Call Cells_Conv(Selection, RB_ID(control))
End Sub

'�\���E��\��
Private Sub RB14_onAction(ByVal control As IRibbonControl)
    Call ShowHide(RB_ID(control))
End Sub

'�ڎ��V�[�g�쐬
Private Sub RB15_onAction(ByVal control As IRibbonControl)
    Call AddInfoSheet(RB_ID(control))
End Sub

'�p�X��
Private Sub RB16_onAction(ByVal control As IRibbonControl)
    Call GetPath(Selection, RB_ID(control))
End Sub

Private Sub RB16b_onAction(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Call SetPathParam(RB_ID(control), pressed)
End Sub

Private Sub RB16b_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetPathParam(RB_ID(control))
End Sub

'----------------------------------------
'�r���g
'----------------------------------------

'�g�ݒ�
Private Sub RB21_onAction(ByVal control As IRibbonControl)
    Call KeisenWaku(RB_ID(control), Selection)
End Sub

'�t�B���^
Private Sub RB22_onAction(ByVal control As IRibbonControl)
    Call KeisenWaku(6, Selection)
End Sub

'������
Private Sub RB23_onAction(ByVal control As IRibbonControl)
    Call KeisenWaku(7, Selection)
End Sub

'�g�Œ�
Private Sub RB24_onAction(ByVal control As IRibbonControl)
    Call KeisenWaku(8, Selection)
End Sub

'���o���F
Private Sub RB25_onAction(ByVal control As IRibbonControl)
    Call KeisenWaku(9, Selection)
End Sub

Private Sub RB25_getImage(ByVal control As IRibbonControl, ByRef bitmap As Variant)
    Dim pic As IPictureDisp
    Set pic = Application.CommandBars.GetImageMso("FontFillBackColorPicker", 32, 32)
    Set bitmap = pic
End Sub

'�͂��N���A
Private Sub RB26_onAction(ByVal control As IRibbonControl)
    Call KeisenWaku(10, Selection)
End Sub

'�͂��N���A
Private Sub RB28_onAction(ByVal control As IRibbonControl)
    Call KeisenWaku(11, Selection)
End Sub

'�}�[�W��
Private Sub RB27_onAction(ByVal control As IRibbonControl)
    SetTableMargin xlRows
    SetTableMargin xlColumns
    g_ribbon.InvalidateControl "RB27"
End Sub

Private Sub RB27_getLabel(ByRef control As Office.IRibbonControl, ByRef label As Variant)
   label = "�s: " & GetTableMargin(xlRows) & ", ��: " & GetTableMargin(xlColumns)
End Sub

'----------------------------------------
'�e���v���[�g�@�\
'----------------------------------------

Private Sub RB3_onAction(ByVal control As IRibbonControl)
    Call TemplateMenu(RB_ID(control))
    Select Case RB_ID(control)
    Case 8 '�X�V
        RBTable_Init
        g_ribbon.InvalidateControl "RB4_1"
        g_ribbon.InvalidateControl "RB4_2"
        g_ribbon.InvalidateControl "RB4_3"
        g_ribbon.InvalidateControl "RB4_4"
        g_ribbon.InvalidateControl "RB4_5"
        g_ribbon.InvalidateControl "RB4_6"
        g_ribbon.InvalidateControl "RB4_7"
        g_ribbon.InvalidateControl "RB4_8"
        g_ribbon.InvalidateControl "RB4_9"
    Case 9 '�J��
        'g_dev_visible = Not g_dev_visible
        g_ribbon.InvalidateControl "RB9"
    End Select
End Sub

Private Sub RB33_getLabel(ByRef control As IRibbonControl, ByRef label As Variant)
    If ThisWorkbook.IsAddin Then label = "�u�b�N�J��" Else label = "�u�b�N����"
    g_ribbon.Invalidate
    DoEvents
End Sub

Private Sub RB3_getEnabled(control As IRibbonControl, ByRef enable As Variant)
    enable = ThisWorkbook.IsAddin
End Sub

'----------------------------------------
'common
'----------------------------------------

Private Sub RB_onAction(control As IRibbonControl)
    Call RBTable_onAction(RB_ID(control))
End Sub

Private Sub RB_getVisible(control As IRibbonControl, ByRef Visible As Variant)
    Call RBTable_getVisible(RB_ID(control), Visible)
End Sub

Private Sub RB_getLabel(control As IRibbonControl, ByRef label As Variant)
    Call RBTable_getLabel(RB_ID(control), label)
End Sub

Private Sub RB_onGetImage(control As IRibbonControl, ByRef bitmap As Variant)
    Call RBTable_onGetImage(RB_ID(control), bitmap)
End Sub

Private Sub RB_getSize(control As IRibbonControl, ByRef Size As Variant)
    Call RBTable_getSize(RB_ID(control), Size)
End Sub

