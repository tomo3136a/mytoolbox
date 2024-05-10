Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI

Private Sub AddinDev_onLoad(ByVal ribbon As IRibbonUI)
    Set g_ribbon = ribbon
End Sub

'----------------------------------
'アドイン開発
'----------------------------------

'アドイン選択
Private Sub AddinDevSel_getItemCount(control As IRibbonControl, ByRef returnedVal)
    returnedVal = UserAddinCount
End Sub

Private Sub AddinDevSel_getItemID(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = index
End Sub

Private Sub AddinDevSel_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = UserAddinName(index)
End Sub

Private Sub AddinDevSel_getSelectedItemID(control As IRibbonControl, ByRef returnedVal)
    Dim i As Integer
    i = CurrentAddinID
    SetAddinName UserAddinName(i)
    returnedVal = i
End Sub

Private Sub AddinDevSel_onAction(control As IRibbonControl, id As String, index As Integer)
    SetAddinName UserAddinName(index)
End Sub

'アドイン操作
Private Sub AddinDev_onAction(ByVal control As IRibbonControl)
    AddinDevApp Val("0" & control.Tag)
    If g_ribbon Is Nothing Then
        MsgBox "g_ribbon"
        Exit Sub
    End If
    g_ribbon.Invalidate
    DoEvents
End Sub

Private Sub AddinDev_getEnabled(ByVal control As IRibbonControl, ByRef enable As Variant)
    Select Case Val("0" & control.Tag)
    Case 4
        enable = (LCase(Right(ActiveWorkbook.name, 5)) = ".xlam")
    Case 34
        enable = (LCase(Right(ActiveWorkbook.name, 5)) <> ".xlam")
    End Select
End Sub

