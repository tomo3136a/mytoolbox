VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IDF_PlaceForm 
   Caption         =   "配置"
   ClientHeight    =   5235
   ClientLeft      =   -15
   ClientTop       =   -150
   ClientWidth     =   3510
   OleObjectBlob   =   "IDF_PlaceForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "IDF_PlaceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private sFileName As String
Private sTool As String
Private sDate As String
Private iVer As Long

Private Sub CommandButtonLib_Click()
    Dim sht As Worksheet
    Set sht = SelectSheet
    If sht Is Nothing Then Exit Sub
    If sht Is ActiveSheet Then Exit Sub
    TextBoxLib.Value = sht.name
    Call SetRtParam("IDF", "lib", sht.name)
    AddItemFromRange ComboBoxGeo, 11
    AddItemFromRange ComboBoxNum, 12
    ComboBoxNum.SetFocus
End Sub

Private Sub ComboBoxNum_Change()
    UpdateComboBox ComboBoxNum, ComboBoxGeo
End Sub

Private Sub ComboBoxGeo_Change()
    UpdateComboBox ComboBoxGeo, ComboBoxNum
End Sub

Private Sub TextBoxX_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then KeyAscii = 0
End Sub

Private Sub TextBoxY_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then KeyAscii = 0
End Sub

Private Sub TextBoxZ_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then KeyAscii = 0
End Sub

Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub AddButton_Click()
    If TestData() Then Exit Sub
    
    Dim ce As Range
    Set ce = TableLeftTop(ActiveCell)
    If ce.Value = "" Then
        Set ce = ce.Parent.Cells(1, 1)
        Call AddHeader(ce)
    End If
    Set ce = LeftBottom(ce)
    ce.Offset(1).Select
    Call WriteData(ce)
    ComboBoxRef.Value = ""
    ComboBoxRef.SetFocus
End Sub

Private Sub UserForm_Initialize()
    TextBoxZ.Value = 0
    TextBoxA.Value = 0
    sFileName = "-"
    sTool = "-"
    sDate = "10/22/96.16:41:37"
    iVer = 1
    
    Call ComboBoxRef.AddItem("")
    Call ComboBoxRef.AddItem("NOREFDES")
    Call ComboBoxRef.AddItem("BOARD")
    
    Call ComboBoxSide.AddItem("TOP")
    Call ComboBoxSide.AddItem("BOTTOM")
    
    Call ComboBoxStatus.AddItem("PLACED")
    Call ComboBoxStatus.AddItem("UNPLACED")
    Call ComboBoxStatus.AddItem("MCAD")
    Call ComboBoxStatus.AddItem("ECAD")
    
    Call ComboBoxUnit.AddItem("MM")
    Call ComboBoxUnit.AddItem("THOU")
    
    Dim s As String
    s = GetRtParam("IDF", "lib")
    If Not s = "" Then TextBoxLib.Value = s
End Sub

Private Sub NextRowSelect()
    Dim ce As Range
    Set ce = TableLeftTop(ActiveCell)
    If ce.Value = "" Then
        Set ce = ce.Parent.Cells(1, 1)
        Call AddHeader(ce)
    End If
    Set ce = LeftBottom(ce)
    ce.Offset(1).Select
End Sub


Private Sub UpdateComboBox(src As Object, dst As Object)
    If src.ListIndex < 0 Then Exit Sub
    If dst.Value = dst.List(src.ListIndex) Then Exit Sub
    dst.Value = dst.List(src.ListIndex)
End Sub

Private Sub AddItemFromRange(o As Object, c As Long)
    If TextBoxLib.Value = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(TextBoxLib.Value)
    
    Dim ra As Range
    Set ra = TableDataRange(ws.Cells(2, 1))
    
    Dim dic As Dictionary
    Set dic = New Dictionary
    
    Dim v As Variant
    Dim i As Long
    For i = 2 To o.ListCount
        v = CStr(o.List(i))
        If Not dic.Exists(v) Then
            dic.Add v, 0
        End If
    Next i
    
    For Each v In ra.Columns(c).Value
        If Not dic.Exists(CStr(v)) Then
            dic.Add CStr(v), 1
            o.AddItem CStr(v)
        End If
    Next v
End Sub

Private Function TestRef() As Boolean
    

End Function

Private Function TestData() As Boolean
    Dim e As Boolean
    TestBlank e, ComboBoxGeo
    TestBlank e, ComboBoxNum
    TestBlank e, ComboBoxRef
    TestBlank e, TextBoxX
    TestBlank e, TextBoxY
    TestBlank e, TextBoxZ
    TestBlank e, TextBoxA
    TestData = e
End Function

Private Sub TestBlank(b As Boolean, o As Object)
    o.BackColor = &H80000005
    If Trim(o.Value) <> "" Then Exit Sub
    o.BackColor = &H80000002
    If Not b Then o.SetFocus
    b = True
End Sub

Private Sub WriteData(ce As Range)
    AddRecord ce
End Sub

Private Sub AddHeader(ce As Range)
    Dim s As String
    s = "ファイル名,ファイルタイプ,仕様,作成ツール,作成日,版数," & _
        "名称,単位,オーナー," & _
        "セクション,形状,部品番号,高さ,長さ,配置,関連,状態," & _
        "ラベル,順番,X座標,Y座標,角度,属性名,属性値"
    ce.Resize(1, 24).Value = Split(s, ",")
End Sub

Private Sub AddRecord(ce As Range)
    Dim rec(0 To 23) As Variant
    rec(0) = sFileName
    rec(1) = "BOARD_FILE"
    If CheckBoxPanel = True Then rec(1) = "PANEL_FILE"
    rec(2) = 3#
    rec(3) = sTool
    rec(4) = sDate
    rec(5) = iVer
    rec(6) = ""
    rec(7) = ComboBoxUnit.Value
    rec(8) = ""
    rec(9) = "PLACEMENT"
    rec(10) = Trim(ComboBoxGeo.Value)
    rec(11) = Trim(ComboBoxNum.Value)
    rec(12) = CDbl(TextBoxZ.Value)
    rec(13) = ""
    rec(14) = ComboBoxSide.Value
    rec(15) = Trim(ComboBoxRef.Value)
    rec(16) = Trim(ComboBoxStatus.Value)
    rec(17) = ""
    rec(18) = ""
    rec(19) = CDbl(TextBoxX.Value)
    rec(20) = CDbl(TextBoxY.Value)
    rec(21) = CDbl(TextBoxA.Value)
    rec(22) = ""
    rec(23) = ""
    
    Set ce = ce.Offset(1)
    ce.Resize(1, 24).Value = rec
End Sub

