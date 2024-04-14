VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIDFPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================
'placement class
'========================================

Private pname As String     '名前
Private pnumb As String     '番号
Private prefs As String     'デザイナ: any,REFDES,BOARD
Private px As Double        'x point
Private py As Double        'y point
Private pm As Double        'mounting offset
Private pa As Double        'angle
Private pside As String     '実装面: TOP,BOTTOM
Private psts As String      'status: PLACED, UNPLACED, ECAD, MCAD

'----------------------------------------

Public Sub SetData(no As Integer, vs As Object)
    Select Case no
    Case 1
    Case 2
        pname = vs(0)
        pnumb = vs(1)
        prefs = vs(2)
    Case Else
        px = vs(0)
        py = vs(1)
        pm = vs(2)
        pa = vs(3)
        pside = vs(4)
        psts = vs(5)
    End Select
End Sub

'----------------------------------------

Public Function Draw(ws As Worksheet, x As Double, y As Double, a As Double, Optional g As Double = 1#, Optional sym As Variant) As Shape
    Dim s As String
    Dim sh As Shape
    If pside = "TOP" Then
        Set sh = sym.Draw(ws, x + g * px, y - g * py, a + pa, g)
    ElseIf pside = "BOTTOM" Then
        Set sh = sym.Draw(ws, x + g * px, y - g * py, a - pa, g, True)
    End If
    s = prefs & " " & sh.id
    sh.name = s
    Set Draw = sh
End Function

'----------------------------------------

Public Property Get name() As String
    name = pname
End Property

Public Property Get PartNumber() As String
    PartNumber = pnumb
End Property

Public Property Get Refarence() As String
    Refarence = prefs
End Property

Public Property Get x() As Double
    x = px
End Property

Public Property Get y() As Double
    y = py
End Property

Public Property Get MountingOffset() As Double
    MountingOffset = pm
End Property

Public Property Get Angle() As Double
    Angle = pa
End Property

Public Property Get side() As String
    side = pside
End Property

Public Property Get Status() As Integer
    Status = psts
End Property

