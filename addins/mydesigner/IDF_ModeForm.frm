VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IDF_ModeForm 
   Caption         =   "読み込み設定"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   12
   ClientWidth     =   3156
   OleObjectBlob   =   "IDF_ModeForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "IDF_ModeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private sValue As String

Public Function Result() As String
    Result = sValue
End Function

Private Sub UserForm_Initialize()
    sValue = False
    TextBox1.Value = IDF_GetParam(1)
    CheckBox1.Value = IDF_IsFlag(1)
    CheckBox2.Value = IDF_IsFlag(2)
    CheckBox3.Value = IDF_IsFlag(3)
    CheckBox4.Value = IDF_IsFlag(4)
    CheckBox5.Value = IDF_IsFlag(5)
    CheckBox6.Value = IDF_IsFlag(6)
    CheckBox7.Value = IDF_IsFlag(7)
    CheckBox8.Value = IDF_IsFlag(8)
End Sub

Private Sub CommandButton2_Click()
    Call IDF_SetParam(1, TextBox1.Value)
    Call IDF_SetFlag(1, CheckBox1.Value)
    Call IDF_SetFlag(2, CheckBox2.Value)
    Call IDF_SetFlag(3, CheckBox3.Value)
    Call IDF_SetFlag(4, CheckBox4.Value)
    Call IDF_SetFlag(5, CheckBox5.Value)
    Call IDF_SetFlag(6, CheckBox6.Value)
    Call IDF_SetFlag(7, CheckBox7.Value)
    Call IDF_SetFlag(8, CheckBox8.Value)
    Call IDF_SetFlag(0, True)
    sValue = True
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    Call IDF_SetFlag(0, False)
    Unload Me
End Sub

Private Sub SpinButton1_SpinUp()
    Dim v As Double
    Dim n As Long
    v = Abs(val(TextBox1))
    If v = 0 Then v = 1
    n = -5
    If v >= 0.0001 Then n = -4
    If v >= 0.001 Then n = -3
    If v >= 0.01 Then n = -2
    If v >= 0.1 Then n = -1
    If v >= 1 Then n = 0
    If v >= 10 Then n = 1
    If v >= 100 Then n = 2
    If v >= 1000 Then n = 3
    If v >= 10000 Then n = 4
    v = 10 ^ n
    v = (wsf.Floor_Math(val(TextBox1) / v) + 1) * v
    If v > 100000 Then Exit Sub
    If v < 0.0001 Then Exit Sub
    TextBox1 = v
End Sub

Private Sub SpinButton1_SpinDown()
    Dim v As Double
    Dim n As Long
    v = Abs(val(TextBox1))
    If v = 0 Then v = 1
    n = -5
    If v > 0.0001 Then n = -4
    If v > 0.001 Then n = -3
    If v > 0.01 Then n = -2
    If v > 0.1 Then n = -1
    If v > 1 Then n = 0
    If v > 10 Then n = 1
    If v > 100 Then n = 2
    If v > 1000 Then n = 3
    If v > 10000 Then n = 4
    v = 10 ^ n
    v = (wsf.Ceiling_Math(val(TextBox1) / v) - 1) * v
    If v > 100000 Then Exit Sub
    If v < 0.0001 Then Exit Sub
    TextBox1 = v
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then KeyAscii = 0
End Sub

