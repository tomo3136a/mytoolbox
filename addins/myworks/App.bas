Attribute VB_Name = "App"
'==================================
'�A�v���P�[�V����
'==================================

'[�Q�Ɛݒ�]
'�uMicrosoft Scripting Runtime�v

Option Explicit
Option Private Module

'----------------------------------------
'common
'----------------------------------------

Private Function SelectRange() As Range
    If TypeName(Selection) = "Range" Then
        Set SelectRange = Selection
        Exit Function
    End If
    Set SelectRange = Range(Selection.TopLeftCell, Selection.BottomRightCell)
End Function


