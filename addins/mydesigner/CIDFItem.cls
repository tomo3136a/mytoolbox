VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIDFItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================
'item class
'========================================

Private psect As String     'section type

Private powner As String    'オーナー ECAD MCAD UNOWNED
Private pname As String     '名前
Private pnumb As String     '番号
Private pside As String     '実装面: TOP,BOTTOM,BOTH,INNER
Private punits As String    '
Private pthickness As String   '高さ/厚み

Private pcol As Collection  '座標データ

'----------------------------------------

Public Sub SetData(no As Integer, vs As Variant)
    Select Case no
    Case 1
        Set pcol = New Collection
        powner = vs(1)
    Case 2
        pname = vs(0)
        pnumb = vs(1)
        pside = ""
        punits = vs(2)
        pthickness = vs(3)
    Case Else
        Dim pos As CIDFNode
        Set pos = New CIDFNode
        pos.SetData vs
        pcol.Add pos
    End Select
End Sub

Public Sub SetOutLineData(vs As Variant)
    Select Case psect
    Case "BOARD_OUTLINE"
        pname = "BOARD_OUTLINE"
        pnumb = ""
        pside = ""
        punits = ""
        pthickness = vs(0)
    Case "PANEL_OUTLINE"
        pname = "PANEL_OUTLINE"
        pnumb = ""
        pside = ""
        punits = ""
        pthickness = vs(0)
    Case "OTHER_OUTLINE"
        pname = "OTHER_OUTLINE_" & vs(2)
        pnumb = vs(0)
        pside = vs(2)
        punits = ""
        pthickness = vs(1)
    Case "ROUTE_OUTLINE"
        pname = "ROUTE_OUTLINE_" & vs(0)
        pnumb = ""
        pside = vs(0)
        punits = ""
        pthickness = ""
    Case "PLACE_OUTLINE"
        pname = "PLACE_OUTLINE_" & vs(0)
        pnumb = ""
        pside = vs(0)
        punits = ""
        pthickness = vs(1)
    End Select
End Sub

Public Sub SetKeepOutData(vs As Variant)
    Select Case psect
    Case "ROUTE_KEEPOUT"
        pname = "ROUTE_KEEPOUT_" & vs(0)
        pnumb = ""
        pside = vs(0)
        punits = ""
        pthickness = ""
    Case "VIA_KEEPOUT"
        pname = "VIA_KEEPOUT"
        pnumb = ""
        pside = ""
        punits = ""
        pthickness = ""
    Case "PLACE_KEEPOUT"
        pname = "PLACE_KEEPOUT_" & vs(0)
        pnumb = ""
        pside = vs(0)
        punits = ""
        pthickness = vs(1)
    End Select
End Sub

Public Sub SetRegionData(no As Integer, vs As Variant)
    Select Case psect
    Case "PLACE_REGION"
        pname = "PLACE_REGION_" & vs(0)
        pnumb = ""
        pside = vs(0)
        punits = ""
        pthickness = vs(1)
    End Select
End Sub

Public Sub SetDorillHoleData(no As Integer, vs As Variant)
    Select Case psect
    Case "BOARD_OUTLINE"
        pname = "BOARD_OUTLINE"
        pnumb = ""
        pside = ""
        punits = ""
        pthickness = vs(0)
    End Select
End Sub

'----------------------------------------

Public Function Draw(ws As Worksheet, x As Double, y As Double, a As Double, Optional g As Double = 1#, Optional f As Boolean) As Shape
    Dim ns As Collection
    Set ns = New Collection
    Dim sh As Shape
    Dim s As String
    
    Dim px As Double
    Dim py As Double
    Dim pos As CIDFNode
    Dim fb As FreeformBuilder
    
    Dim i As Long
    For i = 1 To pcol.Count
        Set pos = pcol(i)
        If fb Is Nothing Then
            px = pos.AX(a, x, g, f)
            py = pos.AY(a, y, g, True)
            Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, px, py)
        ElseIf pcol(i - 1).Label <> pos.Label Then
            Set sh = fb.ConvertToShape
            s = sh.name
            ns.Add s
            Set fb = Nothing
            px = pos.AX(a, x, g, f)
            py = pos.AY(a, y, g, True)
            Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, px, py)
        Else
            If pos.Angle = 0 Then
                px = pos.AX(a, x, g, f)
                py = pos.AY(a, y, g, True)
                fb.AddNodes msoSegmentLine, msoEditingAuto, px, py
            ElseIf -360 < pos.Angle And pos.Angle < 360 Then
                px = pos.AX(a, x, g, f)
                py = pos.AY(a, y, g, True)
                'TODO: 線で代用、線から弧に変更
                fb.AddNodes msoSegmentLine, msoEditingAuto, px, py
            Else
                Dim tx As Double
                Dim ty As Double
                tx = px
                ty = py
                px = pos.AX(a, x, g, f)
                py = pos.AY(a, y, g, True)
                Dim mx As Double
                Dim my As Double
                mx = (px + tx) / 2
                my = (py + ty) / 2
                Dim dx As Double
                Dim dy As Double
                dx = px - tx
                dy = py - ty
                Dim r As Double
                r = Sqr(dx * dx + dy * dy)
                Set sh = ws.Shapes.AddShape(msoShapeOval, tx - r, ty - r, 2 * r, 2 * r)
                s = sh.name
                ns.Add s
                Set fb = Nothing
            End If
        End If
    Next i
    If Not fb Is Nothing Then
        Set sh = fb.ConvertToShape
        s = sh.name
        ns.Add s
        Set fb = Nothing
    End If
    If ns.Count = 0 Then Exit Function
    If ns.Count > 1 Then Set sh = ws.Shapes.Range(ToArray(ns)).Group
    s = pname & " " & sh.id
    sh.name = s
    Set ns = Nothing
    Set Draw = sh
End Function

'----------------------------------------

Public Property Get Left() As Double
    Dim x As Double
    Dim v As Variant
    x = pcol(1).x
    For Each v In pcol
        If x > v.x Then x = v.x
    Next v
    Left = x
End Property

Public Property Get Right() As Double
    Dim x As Double
    Dim v As Variant
    x = pcol(1).x
    For Each v In pcol
        If x < v.x Then x = v.x
    Next v
    Right = x
End Property

Public Property Get Top() As Double
    Dim y As Double
    Dim v As Variant
    y = pcol(1).y
    For Each v In pcol
        If y > v.y Then y = v.y
    Next v
    Top = y
End Property

Public Property Get Bottom() As Double
    Dim y As Double
    Dim v As Variant
    y = pcol(1).y
    For Each v In pcol
        If y < v.y Then y = v.y
    Next v
    Bottom = y
End Property

'----------------------------------------

Public Property Get Section() As String
    Section = psect
End Property

Public Property Let Section(v As String)
    psect = v
End Property

Public Property Get Owner() As String
    Owner = powner
End Property

Public Property Let Owner(v As String)
    powner = v
End Property

Public Property Get name() As String
    name = pname
End Property

Public Property Get PartNumber() As String
    PartNumber = pnumb
End Property

Public Property Get Units() As String
    Units = punits
End Property

Public Property Get Thickness() As String
    Thickness = pthickness
End Property

Public Property Get Items() As Collection
    Set Items = pcol
End Property
