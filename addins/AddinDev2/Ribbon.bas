Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI

'----------------------------------------
'ribbon helper
'----------------------------------------

Private Sub RefreshRibbon(Optional id As String)
    If Not g_ribbon Is Nothing Then
        If id = "" Then
            g_ribbon.Invalidate
        Else
            g_ribbon.InvalidateControl id
        End If
    End If
    DoEvents
End Sub

Private Function RibbonID(control As IRibbonControl, Optional n As Long) As Long
    Dim vs As Variant
    vs = Split(control.id, ".")
    If UBound(vs) < n Then Exit Function
    RibbonID = Val("0" & vs(UBound(vs) - n))
End Function

'----------------------------------------
'module event
'----------------------------------------

Private Sub AddinDev2_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
End Sub

'----------------------------------------
'function event
'----------------------------------------

Private Sub AddinDev2_onAction(ByVal control As IRibbonControl)
    ReloadAddin "AddinDev.xlam"
End Sub

'----------------------------------------
'function
'----------------------------------------

Private Sub ReloadAddin(name As String)
    If ThisWorkbook.name Like name Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim src As String, dst As String
    src = fso.BuildPath(ThisWorkbook.Path, "tmp")
    src = fso.BuildPath(src, fso.GetBaseName(name) & ".zip")
    dst = fso.BuildPath(ThisWorkbook.Path, name)
    
    Dim ai As AddIn
    For Each ai In AddIns
        If ai.name Like name Then Exit For
    Next ai
    If ai Is Nothing Then Exit Sub
    
    ai.Installed = False
    If fso.FileExists(src) Then
        If fso.FileExists(dst) Then fso.DeleteFile dst
        fso.MoveFile src, dst
    End If
    ai.Installed = True
End Sub

