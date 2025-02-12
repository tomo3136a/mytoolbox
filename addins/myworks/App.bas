Attribute VB_Name = "App"
'==================================
'アプリケーション
'==================================

'[参照設定]
'「Microsoft Scripting Runtime」

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


