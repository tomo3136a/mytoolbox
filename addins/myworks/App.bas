Attribute VB_Name = "App"
'==================================
'アプリケーション
'==================================

'[参照設定]
'「Microsoft Scripting Runtime」

Option Explicit
Option Private Module

Public Sub eof()
    ScreenUpdateOn
End Sub

'----------------------------------------
'基本
'----------------------------------------

'アプリケーション情報
Public Function app_name() As String
    app_name = "Works"
End Function

'----------------------------------------
'
'----------------------------------------

'Private Function SelectRange() As Range
'    If TypeName(Selection) = "Range" Then
'        Set SelectRange = Selection
'        Exit Function
'    End If
'    Set SelectRange = Range(Selection.TopLeftCell, Selection.BottomRightCell)
'End Function

