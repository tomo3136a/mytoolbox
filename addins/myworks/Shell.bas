Attribute VB_Name = "Shell"
Option Explicit
'Option Private Module

'========================================
'�V�F�����W���[��
'========================================

'----------------------------------------
'�V�F�����s
'----------------------------------------

Private Sub ShellExectue_1(cmdline As String)
    With CreateObject("Wscript.Shell")
        .Run cmdline
    End With
End Sub

Private Sub ShellExectue_2(cmdline As String)
    Dim shell As Object
    Set shell = CreateObject("Shell.Application")
    shell.ShellExecute cmdline, "", "", "open", 1
    If Not shell Is Nothing Then
        Set shell = Nothing
    End If
End Sub

'----------------------------------------
'�V�F������
'----------------------------------------

Private Sub ShellCopy(src As String, dst As String)
    Dim shell As Object
    Set shell = CreateObject("Shell.Application")
    shell.Namespace(CVar(dst)).CopyHere shell.Namespace(CVar(src)).Items
    If Not shell Is Nothing Then
        Set shell = Nothing
    End If
End Sub

