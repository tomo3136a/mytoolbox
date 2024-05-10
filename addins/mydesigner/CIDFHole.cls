VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIDFHole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================
'hole class
'========================================

Private pd As Double        'diameter
Private px As Double        'x point
Private py As Double        'y point
Private pstyle As String    'plating style:     PTH, NPTH
Private passoc As String    'association part:  BOARD, NOREFDES, PANEL
Private phtype As String    'hole type:         PIN, VIA, MTG, TOOL, Other
Private powner As String    'hole owner:        ECAD, MCAD, UNOWNED

'----------------------------------------

Public Sub SetData(va As Variant)
    pd = va(0)
    px = va(1)
    py = va(2)
    pstyle = va(3)
    passoc = va(4)
    phtype = va(5)
    powner = va(6)
End Sub

'----------------------------------------

'Public Property Get Diameter() As Integer
'    Diameter = pd
'End Property

'Public Property Get x() As Integer
'    x = px
'End Property

'Public Property Get y() As Integer
'    y = py
'End Property

'Public Property Get PlatedStyle() As Integer
'    PlatedStyle = pstyle
'End Property

'Public Property Get AssocPart() As Integer
'    AssocPart = passoc
'End Property

'Public Property Get HoleType() As Integer
'    HoleType = phtype
'End Property

'Public Property Get Owner() As Integer
'    Owner = powner
'End Property

Public Property Get id() As String
    id = phtype & "_" & passoc
End Property

'----------------------------------------

Public Function Draw(ws As Worksheet, x As Double, y As Double, a As Double, Optional g As Double = 1#) As Shape
    Dim r As Double
    Dim xa As Double
    Dim ya As Double
    r = g * pd / 2
    xa = x + g * px - r
    ya = y - g * py - r
    
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeOval, xa, ya, 2 * r, 2 * r)
    sh.name = pstyle & "-" & pd & " " & sh.id
    Set Draw = sh
End Function

