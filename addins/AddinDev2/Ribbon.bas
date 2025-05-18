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
'イベント
'----------------------------------------

'起動時実行
Private Sub AddinDev2_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
    MsgBox "AddinDev2_onLoad"
End Sub

'----------------------------------------
'ハンドラ
'----------------------------------------

Private Sub AddinDev2_onAction(ByVal control As IRibbonControl)
    DeployAddin "AddinDev.xlam"
End Sub

'----------------------------------------
'機能
'----------------------------------------

Private Sub DeployAddin(name As String)
    If ThisWorkbook.name Like name Then Exit Sub
    MsgBox "DeployAddin( " & name & " )"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim src As String
    src = fso.BuildPath(ThisWorkbook.Path, "tmp")
    src = fso.BuildPath(src, fso.GetBaseName(name) & ".zip")
    If Not fso.FileExists(src) Then Exit Sub
    '
    Dim dst As String
    dst = fso.BuildPath(ThisWorkbook.Path, name)
    '
    Dim ai As AddIn
    For Each ai In AddIns
        If ai.name Like name Then Exit For
    Next ai
    If ai Is Nothing Then Exit Sub
    
    ai.Installed = False
    If fso.FileExists(dst) Then fso.DeleteFile dst
    fso.MoveFile src, dst
    ai.Installed = True
    
    Set fso = Nothing
    MsgBox name & "を更新しました"
End Sub

