Attribute VB_Name = "RibbonTable"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private Type RBInfo
    label As String
    Image As String
    Size As Boolean
End Type

Private g_info(10) As RBInfo

'----------------------------------
'ƒe[ƒuƒ‹
'----------------------------------

'----------------------------------------

'Private Sub RB1_onAction(ByVal control As IRibbonControl)
'    Call ReportDsp(RibbonID(control), CInt("0" & control.tag))
'End Sub


'----------------------------------
'“Ç‚Ýž‚Ý
'----------------------------------

Public Sub RBTable_Init()
    With g_info(1)
        .label = "AAaac"
        .Image = "DatabaseInsert"
        .Size = True
    End With
    With g_info(2)
        .label = "bbXXX"
        .Image = "DatabaseInsert"
        .Size = True
    End With
    With g_info(3)
        .label = "cc"
        .Image = "DatabaseInsert"
        .Size = True
    End With
    With g_info(4)
        .label = "dd"
        .Image = "DatabaseInsert"
    End With
End Sub

Public Sub RBTable_Update()
    With g_info(1)
        .label = "AAaac"
        .Image = "DatabaseInsert"
        .Size = True
    End With
    With g_info(2)
        .label = "bbXXX"
        .Image = "DatabaseInsert"
        .Size = True
    End With
    With g_info(3)
        .label = "cc"
        .Image = "DatabaseInsert"
        .Size = True
    End With
    With g_info(4)
        .label = "dd"
        .Image = "DatabaseInsert"
    End With
End Sub

Public Sub RBTable_onAction(id As Integer)
    Select Case id
    Case 1
    Case 2
    Case 3
    Case 4
    Case 5
    Case 6
    Case 7
    Case 8
    Case 9
    End Select
End Sub

Public Sub RBTable_getVisible(id As Integer, ByRef Visible As Variant)
    If Len(g_info(id).label) > 0 Then Visible = True
End Sub

Public Sub RBTable_getLabel(id As Integer, ByRef label As Variant)
    label = g_info(id).label
End Sub

Public Sub RBTable_onGetImage(id As Integer, ByRef bitmap As Variant)
    Dim s As String
    s = g_info(id).Image
    If Not s = "" Then Set bitmap = Application.CommandBars.GetImageMso(s, 80, 80)
End Sub

Public Sub RBTable_getSize(id As Integer, ByRef Size As Variant)
    If g_info(id).Size Then Size = 1
End Sub

