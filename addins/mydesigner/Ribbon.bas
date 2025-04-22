Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI
Private g_select As Integer

'----------------------------------------
'ribbon helper
'----------------------------------------

'���{�����X�V
Private Sub RefreshRibbon(Optional id As String)
    If g_ribbon Is Nothing Then
    ElseIf id = "" Then g_ribbon.Invalidate
    Else: g_ribbon.InvalidateControl id
    End If
    DoEvents
End Sub

'���{��ID�ԍ��擾
Private Function RibbonID(control As IRibbonControl, Optional n As Long) As Long
    Dim vs As Variant
    vs = Split(re_replace(control.id, "[^0-9.]", ""), ".")
    If UBound(vs) >= n Then RibbonID = val("0" & vs(UBound(vs) - n))
End Function

'----------------------------------------
'�C�x���g
'----------------------------------------

'�N�������s
Private Sub Designer_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    Draw_ResetParam
    IDF_ResetParam
End Sub

'�e�L�X�g����
Private Sub Designer_onChange(ByRef control As IRibbonControl, ByRef text As String)
    Call Draw_SetParam(RibbonID(control), text)
End Sub

Private Sub Designer_getText(ByRef control As IRibbonControl, ByRef text As Variant)
    text = Draw_GetParam(RibbonID(control))
End Sub

'�`�F�b�N�{�b�N�X
Private Sub Designer_onActionPressed(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Call Draw_SetParam(RibbonID(control), IIf(pressed, 1, 0))
End Sub

Private Sub Designer_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Draw_IsParamFlag(RibbonID(control))
End Sub

'----------------------------------------
'�����j���[
'  1.x: �}�`����@�\
'  2.x: �}�`���X�g
'  3.x: �}�`�̍X�V
'  4.x: ���i�z�u�@�\
'  5.x: ��}�@�\(IDF)
'----------------------------------------

Private Sub Designer_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 1: Draw1_Menu RibbonID(control)
    Case 2: Draw2_Menu RibbonID(control)
    Case 3: Draw3_Menu RibbonID(control)
    End Select
End Sub

'----------------------------------------
'���@�\�O���[�v1
'�}�`����@�\
'----------------------------------------

Private Sub Designer1_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 1: Draw1_Menu RibbonID(control)
    Case 2: Draw2_Menu RibbonID(control)
    Case 3: Draw3_Menu RibbonID(control)
    End Select
End Sub

Private Sub Draw1_Menu(id As Long, Optional opt As Variant)
    Select Case id
    Case 1: SetDefaultShapeStyle        '�W���}�`�ݒ�
    Case 2:
    Case 3: RemoveSharps                '�}�`�폜
    Case 4: ConvertToPicture            '�}�`���G�ɕϊ�
    Case 5: SetTextBoxStyle             '�e�L�X�g�{�b�N�X��{�ݒ�
    Case 6: ToggleVisible 0             '�h��Ԃ��\��ON/OFF
    Case 7: ToggleVisible 3             '3D�\��ON/OFF
    Case 8: OriginAlignment             '���_���킹
    Case 9: UpdateShapeName ActiveSheet '�}�`���ꊇ�X�V
    Case 10: FlipShapes                 '�\�����]
    End Select
End Sub

'----------------------------------------
'���@�\�O���[�v2
'�}�`����@�\
'  1x: �}�`���X�g
'  2x: �}�`�̍X�V
'----------------------------------------

Private Sub Designer2_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 2: Draw2_Menu RibbonID(control)
    Case 3: Draw3_Menu RibbonID(control)
    End Select
End Sub

Private Sub Draw2_Menu(id As Long, Optional opt As Variant)
    Select Case id
    Case 1: ListShapeInfo               '�ꗗ�擾
    Case 2: AddShapeListName            '���O�ǉ�
    Case 3: ApplyShapeInfo ActiveCell   '�}�`���K�p
    Case 4: SelectShapeName             '�}�`���I��
    Case 5: UpdateShapeInfo             '�f�[�^�擾
    End Select
End Sub

'----------------------------------------
''���@�\�O���[�v3
'
'----------------------------------------

Private Sub Draw3_Menu(id As Long, Optional opt As Variant)
    AddListShapeHeader ActiveCell, id   '�w�b�_���ڒǉ�
End Sub

'----------------------------------------
''���@�\�O���[�v4
'���i�z�u�@�\
'----------------------------------------

'�}�`�A�C�e��
Private Sub Designer4_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 4: Draw4_Menu RibbonID(control)
    End Select
    RefreshRibbon "c41"
End Sub

Private Sub Draw4_Menu(id As Long, Optional opt As Variant)
    Select Case id
    Case 1: AddDrawItem             '�z�u
    Case 2: CopyDrawItem            '�R�s�[
    Case 3: DrawItemEntry           '�o�^
    Case 4: DrawItemDelete          '�폜
    Case 5: DuplicateDrawItemSheet  '�ݒ胍�[�J����
    Case 6: ImportDrawItemSheet     '�ݒ�V�[�g�捞
    'Case 7: ToggleAddinBook
    Case Else:
    End Select
End Sub

'���i�I��
Private Sub Designer4_getItemCount(control As IRibbonControl, ByRef returnedVal)
    Dim cnt As Long
    DrawItemCount cnt
    If cnt < 1 Then cnt = 1
    returnedVal = cnt
End Sub

Private Sub Designer4_getItemID(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    returnedVal = Index
End Sub

Private Sub Designer4_getItemLabel(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    Dim s As String
    DrawItemName Index, s
    returnedVal = s
End Sub

Private Sub Designer4_getSelectedItemID(control As IRibbonControl, ByRef returnedVal)
    Dim cnt As Long
    DrawItemCount cnt
    If cnt > 0 Then cnt = cnt - 1
    If g_select > cnt Then g_select = cnt
    DrawItemSelect g_select
    returnedVal = g_select
End Sub

Private Sub Designer4_onActionDropDown(control As IRibbonControl, id As String, Index As Integer)
    g_select = Index
    DrawItemSelect g_select
    If Not g_ribbon Is Nothing Then g_ribbon.InvalidateControl control.id
End Sub

'----------------------------------------
''���@�\�O���[�v5
'�^�C�~���O�}
'----------------------------------------

Private Sub Designer5_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 5: Draw5_Menu RibbonID(control)
    End Select
End Sub

Private Sub Draw5_Menu(id As Long, Optional opt As Variant)
    Select Case id
    Case 1: GenerateTimeChart 1     '�^�C���`���[�g�f�[�^�쐬(�N���b�N)
    Case 2: GenerateTimeChart 2     '�^�C���`���[�g�f�[�^�쐬(�J�E���^)
    Case 3: GenerateTimeChart 3     '�^�C���`���[�g�f�[�^�쐬(���W�b�N)
    Case 7: AddDrawingSheet         '���ᎆ�V�[�g�ǉ�
    Case 8: DrawTimeChart 1         '��}
    Case 9: DrawTimeChart 2         '��}(�}�`)
    
    Case 5: GenerateTimeChart 5     '���o
    Case 6: GenerateTimeChart 6     '����

    Case 11: ApplyTimeChart 1       '���]
    Case 12: ApplyTimeChart 2       '�N���A
    Case 13: ApplyTimeChart 3       '�v���Z�b�g
    
    Case 21: GenerateTimeChart 1    '�f�[�^�쐬(�N���b�N)
    Case 22: GenerateTimeChart 2    '�f�[�^�쐬(�J�E���^)
    Case 23: GenerateTimeChart 3    '�f�[�^�쐬(���W�b�N)
    Case 24: GenerateTimeChart 4    '�f�[�^�쐬(���o)
    Case 25: GenerateTimeChart 5    '�f�[�^�쐬(����)
    
    Case 31: GenerateTimeChart 11   '�f�[�^�쐬(NOT)
    Case 32: GenerateTimeChart 12   '�f�[�^�쐬(AND)
    Case 33: GenerateTimeChart 13   '�f�[�^�쐬(OR)
    Case 34: GenerateTimeChart 14   '�f�[�^�쐬(XOR)
    Case 35: GenerateTimeChart 15   '�f�[�^�쐬(MUX)
    Case 36: GenerateTimeChart 16   '�f�[�^�쐬(D-FF)
    Case 37: GenerateTimeChart 17   '�f�[�^�쐬(SR-FF)
    Case 38: GenerateTimeChart 18   '�f�[�^�쐬(SYNC)
    Case 39: GenerateTimeChart 19   '�f�[�^�쐬(EDGE)
    Case Else: HelpTimingChart
    End Select
End Sub

'----------------------------------------
''���@�\�O���[�v6
'��}�@�\(IDF)
'----------------------------------------

Private Sub Designer6_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 5: Draw6_Menu RibbonID(control)
    End Select
End Sub

Private Sub Draw6_Menu(id As Long, Optional opt As Variant)
    Select Case id
    Case 1: DrawIDF                     'IDF��}
    Case 2: AddSheetIDF                 'IDF�V�[�g�ǉ�
    Case 3: MacroIDF                    'IDF�}�N��
    Case 4: ImportIDF                   'IDF�t�@�C���ǂݍ���
    Case 5: ExportIDF                   'IDF�t�@�C�������o��
    Case 6: AddRecordIDF                'IDF�s�ǉ�
    Case 7: AddRecordIDF 1              'IDF�s�ǉ�
    Case 8: AddRecordIDF 2              'IDF�s�ǉ�
    Case 10: ResetShapeSize             '�T�C�Y�C��
    Case 11: ResizeShapeScale           '�X�P�[���ύX
    End Select
End Sub

'�e�L�X�g����
Private Sub Designer6_onChange(ByRef control As IRibbonControl, ByRef text As String)
    Call IDF_SetParam(RibbonID(control), text)
End Sub

Private Sub Designer6_getText(ByRef control As IRibbonControl, ByRef text As Variant)
    text = IDF_GetParam(RibbonID(control))
End Sub

'�`�F�b�N�{�b�N�X
Private Sub Designer6_onActionPressed(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Call IDF_SetFlag(RibbonID(control), pressed)
End Sub

Private Sub Designer6_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = IDF_IsFlag(RibbonID(control))
End Sub

'----------------------------------------
''���@�\�O���[�vn
'�g���@�\
'----------------------------------------

'�_�C�i�~�b�N���j���[
Private Sub Designer2_getMenuContent(ByVal control As IRibbonControl, ByRef returnedVal)
    Dim xml As String

    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
          "<button id=""but1"" imageMso=""Help"" label=""Test1"" onAction=""Test1Macro""/>" & _
          "<button id=""but2"" imageMso=""Help"" label=""Test2"" onAction=""Test2Macro""/>" & _
          "<button id=""but3"" imageMso=""Help"" label=""Test3"" onAction=""Test3Macro""/>" & _
          "<button id=""but4"" imageMso=""Help"" label=""Test4"" onAction=""Test4Macro""/>" & _
          "<button id=""but5"" imageMso=""Help"" label=""Help"" onAction=""HelpMacro""/>" & _
          "<button id=""but6"" imageMso=""FindDialog"" label=""Find"" onAction=""FindMacro""/>" & _
          "</menu>"

    returnedVal = xml
End Sub

Sub Test1Macro(control As IRibbonControl)
    Dim shra As ShapeRange
    Dim obj As Object
    On Error Resume Next
    Set obj = Selection
    On Error GoTo 0
    If obj Is Nothing Then Exit Sub
    If TypeName(obj) = "Range" Then Exit Sub
    '
    Dim sz As Integer
    Set shra = obj.ShapeRange
    
    sz = shra.TextFrame2.TextRange.Characters.Count
    If sz > 0 Then
        shra.TextFrame2.TextRange.Characters(sz, 1).Font.Superscript = True
    End If
    'obj.TextFrame.Characters.text = "test"
    
    'MsgBox "Test1 macro"
End Sub

Sub Test2Macro(control As IRibbonControl)
    ToggleAddinBook
    RefreshRibbon "c41"
End Sub

Sub Test3Macro(control As IRibbonControl)
    Dim sn As String
    sn = "#shapes"
    Dim ws As Worksheet
    Set ws = GetSheet(sn)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = sn
    End If
    
End Sub

Sub Test4Macro(control As IRibbonControl)
End Sub

Sub HelpMacro(control As IRibbonControl)
End Sub

Sub FindMacro(control As IRibbonControl)
End Sub

