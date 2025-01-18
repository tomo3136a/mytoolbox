VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IDFModeForm 
   Caption         =   "読み込み設定"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3375
   OleObjectBlob   =   "IDFModeForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "IDFModeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub

Private Sub SpinButton1_SpinDown()
    TextBox1 = val(TextBox1) - 1
End Sub

Private Sub SpinButton1_SpinUp()
    TextBox1 = val(TextBox1) + 1
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Initialize()
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
