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
'アドイン選択
'----------------------------------

Private Sub AddinDevSel_getItemCount(Control As IRibbonControl, ByRef returnedVal)
    returnedVal = UserAddinCount
End Sub

Private Sub AddinDevSel_getItemID(Control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = index
End Sub

Private Sub AddinDevSel_getItemLabel(Control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = UserAddinName(index)
End Sub

Private Sub AddinDevSel_getSelectedItemID(Control As IRibbonControl, ByRef returnedVal)
    Dim i As Integer
    i = CurrentAddinID
    SetAddinName UserAddinName(i)
    returnedVal = i
End Sub

Private Sub AddinDevSel_onAction(Control As IRibbonControl, id As String, index As Integer)
    SetAddinName UserAddinName(index)
End Sub

'----------------------------------
'アドイン操作
'----------------------------------

Private Sub AddinDev_onAction(ByVal Control As IRibbonControl)
    AddinDevApp Val("0" & Control.Tag)
    If g_ribbon Is Nothing Then Exit Sub
    If Val("0" & Control.Tag) <> 52 Then Exit Sub
    g_ribbon.Invalidate
    DoEvents
End Sub

Private Sub AddinDev_getEnabled(ByVal Control As IRibbonControl, ByRef enable As Variant)
    Select Case Val("0" & Control.Tag)
    Case 4
        enable = (LCase(Right(ActiveWorkbook.name, 5)) = ".xlam")
    Case 34
        enable = (LCase(Right(ActiveWorkbook.name, 5)) <> ".xlam")
    End Select
End Sub

Private Sub AddinDev_getImage(ByVal Control As IRibbonControl, ByRef image As Variant)
    Select Case Val("0" & Control.Tag)
    Case 52
        Dim id As Integer
        id = Val("0" + ActiveCell.Value)
        id = 2
        Dim v1 As Variant
        Set v1 = Application.CommandBars.FindControl(id:=id)
        Set image = Application.CommandBars.FindControl(id:=id).Picture
    Case 53
        If ActiveCell.Value = "" Then Exit Sub
        image = ActiveCell.Value
    Case 54
        If ActiveCell.Value = "" Then Exit Sub
        Set image = CommandBars.GetImageMso(ActiveCell.Value, 48, 48)
    Case Else
        'Do Nothing
    End Select
End Sub

