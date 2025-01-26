Attribute VB_Name = "Shell"
Option Explicit
Option Private Module

'========================================
'シェルモジュール
'========================================

'----------------------------------------
'シェル実行
'----------------------------------------

Private Sub ShellExectue(cmdline As String, Optional mode As Integer = 1)
    Select Case mode
    Case 1
        With CreateObject("Wscript.Shell")
            .Run cmdline
        End With
    Case 2
        Dim shell As Object
        Set shell = CreateObject("Shell.Application")
        shell.ShellExecute cmdline, "", "", "open", 1
        If Not shell Is Nothing Then Set shell = Nothing
    End Select
End Sub

'----------------------------------------
'シェル複製
'----------------------------------------

Private Sub ShellCopy(src As String, dst As String)
    Dim shell As Object
    Set shell = CreateObject("Shell.Application")
    shell.Namespace(CVar(dst)).CopyHere shell.Namespace(CVar(src)).Items
    If Not shell Is Nothing Then
        Set shell = Nothing
    End If
End Sub

