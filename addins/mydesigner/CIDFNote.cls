VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIDFNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================
'note class
'========================================

Private px As Double        'x point
Private py As Double        'y point
Private ph As Double        'height
Private pw As Double        'length
Private ps As String        'text

'----------------------------------------

Public Sub SetData(va As Variant)
    px = va(0)
    py = va(1)
    ph = va(2)
    pw = va(3)
    ps = Mid(va(4), 2, Len(va(4)) - 2)
End Sub

'----------------------------------------

'Public Property Get x() As Double
'    x = px
'End Property

'Public Property Get y() As Double
'    y = py
'End Property

'Public Property Get Height() As Double
'    Height = ph
'End Property

'Public Property Get Length() As Double
'    Length = pw
'End Property

'Public Property Get Text() As String
'    Text = ps
'End Property

'----------------------------------------

Private Function AX(a As Double, x As Double, g As Double) As Double
    If a = 0 Then
        AX = x + g * px
        Exit Function
    End If
    AX = x + g * (Cos(a) * px - Sin(a) * py)
End Function

Private Function AY(a As Double, y As Double, g As Double) As Double
    If a = 0 Then
        AY = y - g * py
        Exit Function
    End If
    AY = y - g * (Sin(a) * px + Cos(a) * py)
End Function

'----------------------------------------

Public Function Draw(ws As Worksheet, x As Double, y As Double, a As Double, Optional g As Double = 1#) As Shape
    Dim tx As Double
    Dim ty As Double
    Dim tw As Double
    Dim th As Double
    tx = AX(a, x, g)
    ty = AY(a, y, g)
    tw = g * pw
    th = g * ph * 2
    
    Dim sh As Shape
    Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, tx, ty, tw, th)
    With sh
        .TextFrame.Characters.Font.Size = ph * g * 2
        .TextFrame.Characters.text = ps
        With .TextFrame2
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorBottom
            .HorizontalAnchor = msoAnchorNone
            .WordWrap = msoFalse
        End With
        .TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
    End With
    Dim s As String
    s = "NOTE " & sh.id
    sh.name = s
    Set Draw = sh
End Function

