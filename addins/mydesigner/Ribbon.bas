Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI

Private Sub RD_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
End Sub

Private Sub RefreshAddInsRibbon()
    g_ribbon.Invalidate
    DoEvents
End Sub

Private Function RD_TAG(control As IRibbonControl) As Integer
    RD_TAG = CInt("0" & control.tag)
End Function

Private Function RD_ID(control As IRibbonControl) As Integer
    If control.tag = "" Then
        RD_ID = CInt("0" & Right(control.id, 1))
    Else
        RD_ID = CInt("0" & control.tag)
    End If
End Function

Private Sub RD_onChange(ByRef control As IRibbonControl, ByRef Text As String)
    Call SetDrawParam(RD_ID(control), Text)
End Sub

Private Sub RD_getText(ByRef control As IRibbonControl, ByRef Text As Variant)
    Text = GetDrawParam(RD_ID(control))
End Sub

Private Sub RD_onAction(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    Dim v As Integer
    If pressed Then v = 1
    Call SetDrawParam(RD_ID(control), v)
End Sub

'----------------------------------------
'�}�`����@�\
'----------------------------------------

'�}�`�̃��X�g�A�b�v
Private Sub RD12_onAction(ByVal control As IRibbonControl)
    Call ListShape(ActiveCell, ActiveSheet, "")
End Sub

'�}�`�̍X�V
Private Sub RD13_onAction(ByVal control As IRibbonControl)
    Call UpdateShape(ActiveCell)
End Sub

'�}�`��S�č폜
Private Sub RD14_onAction(ByVal control As IRibbonControl)
    Call RemoveSharp
End Sub

'�}�`���G�ɕϊ�
Private Sub RD15_onAction(ByVal control As IRibbonControl)
    Call ConvToPic
End Sub

'�}�`��{�ݒ�
Private Sub RD16_onAction(ByVal control As IRibbonControl)
    Call SetShapeStyle
End Sub

'�I�u�W�F�N�g�̏�����
Private Sub RD17_onAction(ByVal control As IRibbonControl)
    Call DefaultShapeSetting
End Sub

Private Sub RD18_onAction(ByVal control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim cnt As Integer
    cnt = ws.Shapes.Count
    Dim v As Variant
    For Each v In ws.Shapes
        Dim sh As Shape
        Set sh = v
        Dim u As Variant
        For Each u In sh.GroupItems
            Dim s As String
            s = u.name
        Next u
    Next v
End Sub

'���E���]
Private Sub RD19_onAction(ByVal control As IRibbonControl)
    Dim obj As Object
    Set obj = Selection
    obj.ShapeRange.Flip msoFlipHorizontal
End Sub

'----------------------------------------
'��}�@�\
'----------------------------------------

'�}�`�A�C�e��
Private Sub RD21_onAction(ByVal control As IRibbonControl)
    Dim ra As Range
    If TypeName(Selection) = "Range" Then
        Dim s As String
        s = TypeName(Selection)
        Set ra = Selection
    Else
        Set ra = Range(Selection.TopLeftCell, Selection.BottomRightCell)
        ra.Select
    End If
    On Error Resume Next
    Call DrawGraphItem(RD_ID(control), ra)
    On Error GoTo 0
End Sub

'----------------------------------------
'��}�@�\(IDF)
'----------------------------------------

'IDF�t�@�C���ǂݍ���
Private Sub RD31_onAction(ByVal control As IRibbonControl)
    Select Case RD_ID(control)
    Case 1
        Call ImportIDF
    Case 2
        Call ExportIDF(ActiveSheet)
    End Select
End Sub

'IDF��}
Private Sub RD32_onAction(ByVal control As IRibbonControl)
    Dim ce As Range
    Set ce = ActiveCell
    Select Case RD_ID(control)
    Case 1
        Call DrawIDF(ce.Worksheet, ce.Left, ce.Top)
    Case 2
        Call DrawIDF2(ce.Worksheet, ce.Left, ce.Top)
    End Select
End Sub

'IDF��}
Private Sub RD33_onAction(ByVal control As IRibbonControl)
    Dim ce As Range
    Set ce = ActiveCell
    Call DrawIDF2(ce.Worksheet, ce.Left, ce.Top)
End Sub

'IDF��}
Private Sub RD34_onAction(ByVal control As IRibbonControl)

End Sub


