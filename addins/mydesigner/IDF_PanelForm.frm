VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IDF_PanelForm 
   Caption         =   "パネル追加"
   ClientHeight    =   4065
   ClientLeft      =   -48
   ClientTop       =   -156
   ClientWidth     =   3528
   OleObjectBlob   =   "IDF_PanelForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "IDF_PanelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================
'ダイアログ
'==================================

Option Explicit

Private sFileName As String
Private sTool As String
Private sDate As String
Private iVer As Long

'==================================
'イベント
'==================================

Private Sub TextBoxL_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then KeyAscii = 0
End Sub

Private Sub TextBoxW_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then KeyAscii = 0
End Sub

Private Sub TextBoxH_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   If Not Chr(KeyAscii) Like "[0-9.]" Then KeyAscii = 0
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub AddButton_Click()
    If TestData() Then Exit Sub
    Call NextRowSelect
    Call WriteData(Selection)
    TextBoxGeo.Value = ""
    TextBoxGeo.SetFocus
End Sub

Private Sub UserForm_Initialize()
    Call ComboBoxType.AddItem("PANEL")
    Call ComboBoxType.AddItem("BOARD")
    Call ComboBoxUnit.AddItem("MM")
    Call ComboBoxUnit.AddItem("THOU")
    Call ComboBoxSect.AddItem("BOARD_OUTLINE")
    Call ComboBoxSect.AddItem("PANEL_OUTLINE")
    Call ComboBoxSect.AddItem("OTHER_OUTLINE")
    Call ComboBoxSect.AddItem("ROUTE_OUTLINE")
    Call ComboBoxSect.AddItem("PLACE_OUTLINE")
    Call ComboBoxSect.AddItem("ROUTE_KEEPOUT")
    Call ComboBoxSect.AddItem("PLACE_KEEPOUT")
    Call ComboBoxSect.AddItem("VIA_KEEPOUT")
    TextBoxXPos = 0
    TextBoxYPos = 0
    TextBoxW.Value = 100
    TextBoxL.Value = 100
    TextBoxH.Value = 1.6
    sFileName = ActiveSheet.name
    sTool = "designer"
    sDate = Format(Now(), "MM/dd/yy.hh:mm:ss")
    iVer = 1
End Sub

'==================================
'テスト
'==================================

Private Function TestRef() As Boolean
    
End Function

Private Function TestData() As Boolean
    Dim e As Boolean
    TestBlank e, TextBoxGeo
    TestBlank e, TextBoxXPos
    TestBlank e, TextBoxYPos
    TestBlank e, TextBoxH
    TestBlank e, TextBoxW
    TestBlank e, TextBoxL
    TestData = e
End Function

Private Sub TestBlank(b As Boolean, o As Object)
    o.BackColor = &H80000005
    If Trim(o.Value) <> "" Then Exit Sub
    o.BackColor = &H80000002
    If Not b Then o.SetFocus
    b = True
End Sub

'==================================
'検索
'==================================

Private Sub NextRowSelect()
    Dim ce As Range
    Set ce = TableLeftTop(ActiveCell)
    If ce.Value = "" Then
        Set ce = ce.Parent.Cells(1, 1)
        Call AddHeader(ce)
    End If
    Set ce = LeftBottom(TableRange(ce))
    ce.Offset(1).Select
End Sub

'==================================
'書き出し
'==================================

Private Sub WriteData(ce As Range)
    Dim s As String
    Dim w As Double, l As Double, x As Double, y As Double
    s = Trim(TextBoxW.Value)
    If s = "" Then Exit Sub
    w = s
    s = Trim(TextBoxL.Value)
    If s = "" Then Exit Sub
    l = s
    s = Trim(TextBoxXPos.Value)
    If s = "" Then Exit Sub
    x = s
    s = Trim(TextBoxYPos.Value)
    If s = "" Then Exit Sub
    y = s
    
    If Trim(TextBoxGeo.Value) = "" Then Exit Sub
    
    Call AddRecord(ce, 0, x, y)
    Call AddRecord(ce, 1, x + w, y)
    Call AddRecord(ce, 2, x + w, y + l)
    Call AddRecord(ce, 3, x, y + l)
    Call AddRecord(ce, 4, x, y)
    ce.Select
End Sub

Private Sub AddHeader(ce As Range)
    Dim s As String
    s = "ファイル名,ファイルタイプ,仕様,作成ツール,作成日,版数," & _
        "名称,単位,オーナー," & _
        "セクション,形状,部品番号,高さ,長さ,配置,関連,状態," & _
        "ラベル,順番,X座標,Y座標,角度,属性名,属性値"
    ce.Resize(1, 24).Value = Split(s, ",")
End Sub

Private Sub AddRecord(ce As Range, i As Long, x As Double, y As Double)
    Dim mode As Long
    mode = IIf(ComboBoxType = "LIBRARY", 0, 1)
    
    Dim rec(0 To 23) As Variant
    rec(0) = sFileName
    rec(1) = ComboBoxType.Value & "_FILE"
    rec(2) = 3#
    rec(3) = sTool
    rec(4) = sDate
    rec(5) = iVer
    rec(6) = Trim(TextBoxGeo.Value)
    rec(7) = ComboBoxUnit.Value
    rec(8) = IIf(CheckBoxMecanical.Value = True, "MCAD", "ECAD")
    rec(9) = Trim(ComboBoxSect.Value)
    rec(10) = ""
    rec(11) = ""
    rec(12) = val(TextBoxH.Value)
    rec(13) = ""
    rec(14) = ""
    rec(15) = ""
    rec(16) = ""
    rec(17) = 0
    rec(18) = i
    rec(19) = x
    rec(20) = y
    rec(21) = 0
    rec(22) = ""
    rec(23) = ""
    ce.Resize(1, 24).Value = rec
    Set ce = ce.Offset(1)
End Sub

