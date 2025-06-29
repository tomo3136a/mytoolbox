Attribute VB_Name = "App"
'==================================
'アプリケーション
'==================================

'[参照設定]
'「Microsoft Scripting Runtime」
'「Microsoft VBScript Regular Expressions」

Option Explicit
Option Private Module

'----------------------------------------
'基本処理
'----------------------------------------

'アプリケーション情報
Public Function app_name() As String
    app_name = "Works"
End Function

'画面更新有効化
Public Sub eof()
    ScreenUpdateOn
End Sub

'----------------------------------------
'自動処理
'----------------------------------------

'クローズ時自動保存
Sub Auto_Close()
    If ThisWorkbook.saved = False Then
        ThisWorkbook.Save
    End If
End Sub

'----------------------------------------
'設計メモ
'----------------------------------------

'■セルを選択していない場合はマクロ終了
'If TypeName(Selection) <> "Range" Then Exit Sub
'

'Private Function SelectRange() As Range
'    If TypeName(Selection) = "Range" Then
'        Set SelectRange = Selection
'        Exit Function
'    End If
'    Set SelectRange = Range(Selection.TopLeftCell, Selection.BottomRightCell)
'End Function

