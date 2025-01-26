VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IDF_PartForm 
   Caption         =   "���i�ǉ�"
   ClientHeight    =   4065
   ClientLeft      =   -45
   ClientTop       =   -150
   ClientWidth     =   3525
   OleObjectBlob   =   "IDF_PartForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "IDF_PartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private sFileName As String
Private sTool As String
Private sDate As String
Private iVer As Long

Private Sub UserForm_Initialize()
    Call ComboBoxUnit.AddItem("MM")
    Call ComboBoxUnit.AddItem("THOU")
    sFileName = "-"
    sTool = "-"
    sDate = "10/22/96.16:41:37"
    iVer = 1
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OkButton_Click()
    Dim ce As Range
    Set ce = TableLeftTop(ActiveCell)
    If ce.Value = "" Then
        Set ce = ce.Parent.Cells(1, 1)
        Call AddHeader(ce)
    End If
    Call WriteData(LeftBottom(TableRange(ce)))
    Unload Me
End Sub

Private Sub WriteData(ce As Range)

    Dim s As String
    Dim w As Double, l As Double
    s = Trim(TextBoxW.Value)
    If s = "" Then Exit Sub
    w = s
    s = Trim(TextBoxL.Value)
    If s = "" Then Exit Sub
    l = s
    
    If Trim(TextBoxGeo.Value) = "" Then Exit Sub
    If Trim(TextBoxNum.Value) = "" Then Exit Sub
    
    Call AddRecord(ce, 0, -w / 2, -l / 2)
    Call AddRecord(ce, 1, w / 2, -l / 2)
    Call AddRecord(ce, 2, w / 2, l / 2)
    Call AddRecord(ce, 3, -w / 2, l / 2)
    Call AddRecord(ce, 4, -w / 2, -l / 2)
    
End Sub

Private Sub AddHeader(ce As Range)
        ce.Offset(, 0).Value = "�t�@�C����"
        ce.Offset(, 1).Value = "�t�@�C���^�C�v"
        ce.Offset(, 2).Value = "�d�l"
        ce.Offset(, 3).Value = "�쐬�c�[��"
        ce.Offset(, 4).Value = "�쐬��"
        ce.Offset(, 5).Value = "�Ő�"
        ce.Offset(, 6).Value = "����"
        ce.Offset(, 7).Value = "�P��"
        ce.Offset(, 8).Value = "�I�[�i�["
        ce.Offset(, 9).Value = "�Z�N�V����"
        ce.Offset(, 10).Value = "�`��"
        ce.Offset(, 11).Value = "���i�ԍ�"
        ce.Offset(, 12).Value = "����"
        ce.Offset(, 13).Value = "����"
        ce.Offset(, 14).Value = "�z�u"
        ce.Offset(, 15).Value = "�֘A"
        ce.Offset(, 16).Value = "���"
        ce.Offset(, 17).Value = "���x��"
        ce.Offset(, 18).Value = "����"
        ce.Offset(, 19).Value = "X���W"
        ce.Offset(, 20).Value = "Y���W"
        ce.Offset(, 21).Value = "�p�x"
        ce.Offset(, 22).Value = "������"
        ce.Offset(, 23).Value = "�����l"

End Sub

Private Sub AddRecord(ce As Range, i As Long, x As Double, y As Double)

    Set ce = ce.Offset(1)
    ce.Offset(, 0).Value = sFileName
    ce.Offset(, 1).Value = "LIBRARY_FILE"
    ce.Offset(, 2).Value = 3#
    ce.Offset(, 3).Value = sTool
    ce.Offset(, 4).Value = sDate
    ce.Offset(, 5).Value = iVer
    ce.Offset(, 6).Value = ""
    ce.Offset(, 7).Value = ComboBoxUnit.Value
    ce.Offset(, 8).Value = ""
    ce.Offset(, 9).Value = "ELECTRICAL"
    If CheckBoxMecanical.Value = True Then ce.Offset(, 9).Value = "MECANICAL"
    ce.Offset(, 10).Value = Trim(TextBoxGeo.Value)
    ce.Offset(, 11).Value = Trim(TextBoxNum.Value)
    ce.Offset(, 12).Value = val(TextBoxH.Value)
    ce.Offset(, 13).Value = ""
    ce.Offset(, 14).Value = ""
    ce.Offset(, 15).Value = ""
    ce.Offset(, 16).Value = ""
    ce.Offset(, 17).Value = 0
    ce.Offset(, 18).Value = i
    ce.Offset(, 19).Value = x
    ce.Offset(, 20).Value = y
    ce.Offset(, 21).Value = 0
    ce.Offset(, 22).Value = ""
    ce.Offset(, 23).Value = ""
    
End Sub

Private Sub TextBoxL_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then KeyAscii = 0
End Sub

Private Sub TextBoxW_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then KeyAscii = 0
End Sub

Private Sub TextBoxH_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then KeyAscii = 0
End Sub

