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

Private Sub RefreshRibbon()
    If Not g_ribbon Is Nothing Then g_ribbon.Invalidate
    DoEvents
End Sub

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

'----------------------------------------
'ribbon initialize
'----------------------------------------

Private Sub onLoad(ByVal ribbon As IRibbonUI)
    Set g_ribbon = ribbon
End Sub

'----------------------------------
'アドイン選択
'----------------------------------

Private Sub getItemCount(control As IRibbonControl, ByRef returnedVal)
    returnedVal = UserAddinCount
End Sub

Private Sub getItemID(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = index
End Sub

Private Sub getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = UserAddinName(index)
End Sub

Private Sub getSelectedItemID(control As IRibbonControl, ByRef returnedVal)
    Dim i As Integer
    i = CurrentAddinID
    SetAddinName UserAddinName(i)
    returnedVal = i
End Sub

Private Sub onActionDropDown(control As IRibbonControl, id As String, index As Integer)
    SetAddinName UserAddinName(index)
End Sub

'----------------------------------
'アドイン操作
'----------------------------------

Private Sub onAction(ByVal control As IRibbonControl)
    AddinDevApp RibbonID(control)
    If RibbonID(control) <> 52 Then Exit Sub
    If g_ribbon Is Nothing Then Exit Sub
    g_ribbon.Invalidate
    DoEvents
End Sub

Private Sub getEnabled(ByVal control As IRibbonControl, ByRef enable As Variant)
    Select Case RibbonID(control)
    Case 4
        enable = (LCase(Right(ActiveWorkbook.name, 5)) = ".xlam")
    Case 34
        enable = (LCase(Right(ActiveWorkbook.name, 5)) <> ".xlam")
    End Select
End Sub

Private Sub getImage(ByVal control As IRibbonControl, ByRef image As Variant)
    On Error Resume Next
    Select Case RibbonID(control)
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

