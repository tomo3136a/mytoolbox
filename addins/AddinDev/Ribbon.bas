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

Private Function RibbonID(control As IRibbonControl) As Integer
    Dim s As String
    s = control.Tag
    If s = "" Then s = control.id
    Dim vs As Variant
    vs = Split(s, ".")
    If UBound(vs) >= 0 Then
        RibbonID = Val("0" & vs(UBound(vs)))
        Exit Function
    End If
    vs = Split(s, "_")
    If UBound(vs) >= 0 Then
        RibbonID = Val("0" & vs(UBound(vs)))
        Exit Function
    End If
End Function

Private Sub RefreshRibbon()
    If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
    DoEvents
End Sub

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
'invoke application
'----------------------------------

Private Sub AddinDev_onAction(ByVal control As IRibbonControl)
    AppAddinDev.App RibbonID(control)
    If g_ribbon Is Nothing Then Exit Sub
    Select Case RibbonID(control)
    Case 52
        g_ribbon.Invalidate
        DoEvents
    Case 1
        g_ribbon.Invalidate
        DoEvents
    Case Else
    End Select
End Sub

Private Sub AddinDev_getEnabled(ByVal control As IRibbonControl, ByRef enable As Variant)
    Select Case RibbonID(control)
    Case 39
        enable = (LCase(Right(ActiveWorkbook.name, 5)) = ".xlam")
    Case 34
        enable = (LCase(Right(ActiveWorkbook.name, 5)) <> ".xlam")
    End Select
End Sub

Private Sub AddinDev_getImage(ByVal control As IRibbonControl, ByRef image As Variant)
    On Error Resume Next
    Select Case RibbonID(control)
    Case 53
        Dim id As Integer
        id = Val("0" + GetButtonImage)
        Dim v1 As Variant
        Set v1 = Application.CommandBars.FindControl(id:=id)
        Set image = Application.CommandBars.FindControl(id:=id).Picture
    Case 52
        If GetButtonImage = "" Then
            image = "About"
        Else
            image = GetButtonImage
        End If
    Case 54
        Dim v As Integer
        v = Val(ActiveCell.Value)
        If v = 0 Then Exit Sub
        Dim pic As IPictureDisp
        Set pic = Application.CommandBars.GetImageMso(v, 32, 32)
        Set image = pic
    Case Else
        'Do Nothing
    End Select
    On Error GoTo 0
End Sub
