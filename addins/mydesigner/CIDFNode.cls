VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIDFNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================
'node class
'========================================

Private plabel As String    'Loop Label
Private px As Double        'x point
Private py As Double        'y point
Private pa As Double        'angle(0=lne, 360=circle, othrs=arc)

'----------------------------------------

Public Sub SetData(va As Object)
    plabel = va(0)
    px = va(1)
    py = va(2)
    pa = va(3)
End Sub

'----------------------------------------

Public Property Get Label() As String
    Label = plabel
End Property

Public Property Get x() As Double
    x = px
End Property

Public Property Get y() As Double
    y = py
End Property

Public Property Get Angle() As Double
    Angle = pa
End Property

'----------------------------------------

Public Function AX(a As Double, v As Double, g As Double, f As Boolean) As Double
    Dim p As Double
    If a = 0 Then
        p = g * px
    Else
        p = g * (px * Cos(2 * 3.14 * a / 360) - py * Sin(2 * 3.14 * a / 360))
    End If
    If f Then
        AX = v - p
    Else
        AX = v + p
    End If
End Function

Public Function AY(a As Double, v As Double, g As Double, f As Boolean) As Double
    Dim p As Double
    If a = 0 Then
        p = g * py
    Else
        p = g * (px * Sin(2 * 3.14 * a / 360) + py * Cos(2 * 3.14 * a / 360))
    End If
    If f Then
        AY = v - p
    Else
        AY = v + p
    End If
End Function

