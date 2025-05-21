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
'ribbon initialize
'----------------------------------------

Private Sub AddinDev_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
End Sub

'----------------------------------------
'select addin
'----------------------------------------

Private Sub AddinDev_getItemCount(control As IRibbonControl, ByRef returnedVal)
    returnedVal = UserAddinCount
End Sub

Private Sub AddinDev_getItemID(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = index
End Sub

Private Sub AddinDev_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = UserAddinName(index)
End Sub

Private Sub AddinDev_getSelectedItemID(control As IRibbonControl, ByRef returnedVal)
    Dim i As Integer
    i = CurrentAddinID
    SetAddinName UserAddinName(i)
    returnedVal = i
End Sub

Private Sub AddinDev_onActionDropDown(control As IRibbonControl, id As String, index As Integer)
    SetAddinName UserAddinName(index)
End Sub

'----------------------------------
'AddinDevProc application
'----------------------------------

Private Sub AddinDev_onAction(ByVal control As IRibbonControl)
    Select Case RibbonID(control, 1)
    Case 1: AddinDevFolderProc RibbonID(control): Exit Sub
    Case 3: AddinDevEditProc RibbonID(control)
    Case 5: AddinDevCallProc RibbonID(control)
    End Select
    
    If g_ribbon Is Nothing Then Exit Sub
    g_ribbon.Invalidate
    DoEvents
End Sub

Private Sub AddinDev_getEnabled(ByVal control As IRibbonControl, ByRef enable As Variant)
    Select Case RibbonID(control)
    Case 4: enable = Not (LCase(Right(ActiveWorkbook.name, 5)) Like ".xlam")
    Case 9: enable = (LCase(Right(ActiveWorkbook.name, 5)) Like ".xlam")
    End Select
End Sub

Private Sub AddinDev_getImage(ByVal control As IRibbonControl, ByRef image As Variant)
    Select Case RibbonID(control)
    Case 2
        If GetButtonImage = "" Then
            image = "About"
        Else
            image = GetButtonImage
        End If
    Case 3
        Dim id As Integer
        id = Val("0" + GetButtonImage)
        Dim v1 As Variant
        Set v1 = Application.CommandBars.FindControl(id:=id)
        Set image = Application.CommandBars.FindControl(id:=id).Picture
    Case 4
        Dim v As Integer
        v = Val(ActiveCell.Value)
        If v = 0 Then Exit Sub
        Dim pic As IPictureDisp
        Set pic = Application.CommandBars.GetImageMso(v, 32, 32)
        Set image = pic
    End Select
End Sub
    
