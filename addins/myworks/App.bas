Attribute VB_Name = "App"
'==================================
'�A�v���P�[�V����
'==================================

'[�Q�Ɛݒ�]
'�uMicrosoft Scripting Runtime�v

Option Explicit
Option Private Module

'----------------------------------------
'��{
'----------------------------------------

'�A�v���P�[�V�������
Public Function app_name() As String
    app_name = "Works"
End Function

'��ʍX�V�L����
Public Sub eof()
    ScreenUpdateOn
End Sub

'----------------------------------------
'
'----------------------------------------

Sub Auto_Close()
    If ThisWorkbook.saved = False Then
        ThisWorkbook.Save
    End If
End Sub

'���Z����I�����Ă��Ȃ��ꍇ�̓}�N���I��
'If TypeName(Selection) <> "Range" Then Exit Sub
'

'Private Function SelectRange() As Range
'    If TypeName(Selection) = "Range" Then
'        Set SelectRange = Selection
'        Exit Function
'    End If
'    Set SelectRange = Range(Selection.TopLeftCell, Selection.BottomRightCell)
'End Function

