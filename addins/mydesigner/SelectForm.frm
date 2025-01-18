VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectForm 
   Caption         =   "選択"
   ClientHeight    =   4680
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3630
   OleObjectBlob   =   "SelectForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private sValue As String
Private sMatch As Object

Public Sub reset(Optional s As String, Optional ptn As String)
    Title.Caption = s & "一覧："
    ListBox1.Clear
    sValue = ""
    Set sMatch = Nothing
    If ptn <> "" Then Set sMatch = regex(ptn)
End Sub

Public Sub SetTitle(s As String)
    Title.Caption = s
End Sub

Public Sub SetFilter(ptn As String)
    If ptn <> "" Then Set sMatch = regex(ptn)
End Sub

Public Function Result() As String
    Result = sValue
End Function

Public Function index() As Integer
    index = ListBox1.ListIndex
End Function

Public Sub AddItem(s As String)
    If s = "" Then Exit Sub
    If sMatch Is Nothing Then
        ListBox1.AddItem s
    ElseIf sMatch.Test(s) Then
        ListBox1.AddItem s
    End If
    If ListBox1.ListCount > 0 Then ListBox1.ListIndex = 0
End Sub

Public Function ItemCount() As Integer
    ItemCount = ListBox1.ListCount
End Function

Public Function ItemValue(idx As Integer) As String
    ItemValue = ListBox1.List(idx)
End Function

Public Sub AddNames(obj As Object)
    Dim v As Variant
    For Each v In obj
        AddItem v.name
    Next v
End Sub

Public Sub AddValues(obj As Object)
    Dim v As Variant
    For Each v In obj
        AddItem v.Value
    Next v
End Sub

Public Sub SetRange(ra As Range)
    Dim v As Variant
    For Each v In ra.Value
        Dim s As String
        s = v
        AddItem s
    Next v
End Sub

Private Sub OkButton_Click()
    Dim i As Integer
    i = ListBox1.ListIndex
    If i >= 0 Then sValue = ListBox1.List(i)
    Set sMatch = Nothing
    Hide
End Sub

Private Sub CancelButton_Click()
    sValue = ""
    Set sMatch = Nothing
    Hide
End Sub

Private Sub OkButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    OkButton_Click
End Sub
