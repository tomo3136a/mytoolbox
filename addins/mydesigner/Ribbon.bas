Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI
Private g_select As Integer

'----------------------------------------
'common
'----------------------------------------

Private Function SelectRange() As Range
    Dim ra As Range
    If TypeName(Selection) = "Range" Then
        Dim s As String
        s = TypeName(Selection)
        Set ra = Selection
    Else
        Set ra = Range(Selection.TopLeftCell, Selection.BottomRightCell)
        ra.Select
    End If
    Set SelectRange = ra
End Function

'----------------------------------------
'ribbon helper
'----------------------------------------

Private Sub RefreshRibbon()
    If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
    DoEvents
End Sub

Private Function RibbonID(control As IRibbonControl) As Integer
    Dim s As String
    s = control.Tag
    If s = "" Then s = control.id
    Dim vs As Variant
    vs = Split(s, ".")
    If UBound(vs) >= 0 Then
        RibbonID = val("0" & vs(UBound(vs)))
        Exit Function
    End If
    vs = Split(s, "_")
    If UBound(vs) >= 0 Then
        RibbonID = val("0" & vs(UBound(vs)))
        Exit Function
    End If
End Function

'----------------------------------------
'�ݒ�
'----------------------------------------

Private Sub Designer_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    Call ResetDrawParam
End Sub

'�e�L�X�g����
Private Sub Designer_onChange(ByRef control As IRibbonControl, ByRef text As String)
    Call SetDrawParam(RibbonID(control), text)
End Sub

Private Sub Designer_getText(ByRef control As IRibbonControl, ByRef text As Variant)
    text = GetDrawParam(RibbonID(control))
End Sub

'�`�F�b�N�{�b�N�X
Private Sub Designer_onAction(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Dim v As Integer
    If pressed Then v = 1
    Call SetDrawParam(RibbonID(control), v)
End Sub

'----------------------------------------
'1x. �}�`����@�\
'----------------------------------------

Private Sub Designer1_onAction(ByVal control As IRibbonControl)
    Dim obj As Object
    Set obj = Selection
    '
    Select Case RibbonID(control)
    Case 1
        '�}�`�̃��X�g�A�b�v
        Call ListShape(ActiveCell, ActiveSheet, "")
    Case 2
        '�}�`�̍X�V
        Call UpdateShape(ActiveCell)
    Case 3
        '�}�`��S�č폜
        Call RemoveSharp(ActiveSheet)
    Case 4
        '�}�`���G�ɕϊ�
        Call ConvToPic
    Case 5
        '�e�L�X�g�{�b�N�X��{�ݒ�
        Call SetShapeStyle
    Case 6
        '���ߍ��ݕ\��ON/OFF
        Call InvertFillVisible
        '�I�u�W�F�N�g�̏�����
        'Call DefaultShapeSetting
    Case 8
    Case 9
    End Select
End Sub

'----------------------------------------
'2x. �c�[���@�\
'----------------------------------------

'�}�`�A�C�e��
Private Sub Designer2_onAction(ByVal control As IRibbonControl)
    On Error Resume Next
    Call DrawGraphItem(RibbonID(control), SelectRange)
    On Error GoTo 0
End Sub

'���i�I��
Private Sub Designer2_getItemCount(control As IRibbonControl, ByRef returnedVal)
    Dim ws As Worksheet
    Set ws = TargetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    returnedVal = ws.Shapes.Count
End Sub

Private Sub Designer2_getItemID(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = index
End Sub

Private Sub Designer2_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim ws As Worksheet
    Set ws = TargetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    returnedVal = ws.Shapes(1 + index).name
End Sub

Private Sub Designer2_getSelectedItemID(control As IRibbonControl, ByRef returnedVal)
    returnedVal = g_select
End Sub

Private Sub Designer2_onActionDropDown(control As IRibbonControl, id As String, index As Integer)
    Dim ws As Worksheet
    Set ws = TargetSheet("#shapes")
    If ws Is Nothing Then Exit Sub
    Call SetDrawParam(10, ws.Shapes(index + 1).name)
    g_select = index
    If Not g_ribbon Is Nothing Then g_ribbon.InvalidateControl control.id
End Sub

'�^�[�Q�b�g�V�[�g�擾
Private Function TargetSheet(s As String) As Worksheet
    Dim v As Variant
    Dim ws As Worksheet
    For Each v In ActiveWorkbook.Worksheets
        If v.name = s Then Set ws = v
    Next v
    If ws Is Nothing Then
        For Each v In ThisWorkbook.Worksheets
            If v.name = s Then ws = v
        Next v
    End If
    If ws Is Nothing Then Exit Function
    Set TargetSheet = ws
End Function

'�_�C�i�~�b�N���j���[
Private Sub Designer2_getMenuContent(ByVal control As IRibbonControl, ByRef returnedVal)
    Dim xml As String

    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
          "<button id=""but1"" imageMso=""Help"" label=""Help"" onAction=""HelpMacro""/>" & _
          "<button id=""but2"" imageMso=""FindDialog"" label=""Find"" onAction=""FindMacro""/>" & _
          "</menu>"

    returnedVal = xml
End Sub

Sub HelpMacro(control As IRibbonControl)
    MsgBox "Help macro"
End Sub

Sub FindMacro(control As IRibbonControl)
    MsgBox "Find macro"
End Sub

'----------------------------------------
'3x. ��}�@�\(IDF)
'----------------------------------------

Private Sub Designer3_onAction(ByVal control As IRibbonControl)
    Dim ce As Range
    Set ce = ActiveCell
    '
    Select Case RibbonID(control)
    Case 1
        'IDF�t�@�C���ǂݍ���
        Call ImportIDF
    Case 2
        'IDF�t�@�C�������o��
        Call ExportIDF(ActiveSheet)
    Case 3
        'IDF��}
        Call DrawIDF(ce.Worksheet, ce.Left, ce.Top)
    Case 4
        'IDF��}
        Call DrawIDF(ce.Worksheet, ce.Left, ce.Top, sheet_load:=True)
    Case 5
        'IDF��}
        Call DrawIDF(ce.Worksheet, ce.Left, ce.Top, sheet_load:=True)
    End Select
End Sub

