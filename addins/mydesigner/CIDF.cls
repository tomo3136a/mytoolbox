VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================
'IDF class
'========================================

Private pfiles As Collection    'file collection
Private pholes As Collection    'hole collection
Private pnotes As Collection    'note collection
Private pparts As Collection    'part collection
Private plib As Dictionary      'lib dictionary

Private pitems As Dictionary    'item dictionary

Private pcfile As Integer       'current file id
Private pcitem As Object        'current item CIDFItem, CIDFHole, CIDFNote, CIDFPart

Private psect As String         'section
Private precno As Integer       'record number
Private plineno As Integer      'line number

Private g_px As Double
Private g_py As Double
Private g_pa As Double
Private g_pg As Double

'----------------------------------------

Public Sub Class_Initialize()
    Set pfiles = New Collection
    Set pholes = New Collection
    Set pnotes = New Collection
    Set pparts = New Collection
    Set plib = New Dictionary
    Set pitems = New Dictionary
End Sub

'----------------------------------------

'アイテム数
Public Property Get Count() As Integer
    Count = pitems.Count
End Property

'アイテムディクショナリ
Public Property Get Items() As Dictionary
    Set Items = pitems
End Property

'----------------------------------------

Public Property Get Left() As Double
    Dim x As Double
    Dim v As Variant
    x = pitems.Items(0).Left
    For Each v In pitems.Items
        If x > v.Left Then x = v.Left
    Next v
    Left = x
End Property

Public Property Get Right() As Double
    Dim x As Double
    Dim v As Variant
    x = pitems.Items(0).Right
    For Each v In pitems.Items
        If x < v.Right Then x = v.Right
    Next v
    Right = x
End Property

Public Property Get Top() As Double
    Dim y As Double
    Dim v As Variant
    y = pitems.Items(0).Top
    For Each v In pitems.Items
        If y > v.Top Then y = v.Top
    Next v
    Top = y
End Property

Public Property Get Bottom() As Double
    Dim y As Double
    Dim v As Variant
    y = pitems.Items(0).Bottom
    For Each v In pitems.Items
        If y < v.Bottom Then y = v.Bottom
    Next v
    Bottom = y
End Property

Public Property Get Height() As Double
    Height = Bottom - Top
End Property

Public Property Get Width() As Double
    Width = Right - Left
End Property

'----------------------------------------

Public Sub Parse(va As Variant)
    Dim s As String
    s = va(0)
    precno = precno + 1
    If Left(s, 1) = "." Then
        If Left(s, 5) = ".END_" Then
            psect = ""
            Set pcitem = Nothing
        Else
            psect = UCase(Replace(s, ".", ""))
            precno = 1
        End If
    End If
    If Not psect = "" Then
        Select Case psect
        Case "ELECTRICAL"
            Call SetSymbol(precno, va)
        Case "MECHANICAL"
            Call SetSymbol(precno, va)
        Case "HEADER"
            Call SetHeader(precno, va)
        Case "BOARD_OUTLINE"
            Call SetOutLine(precno, va)
        Case "PANEL_OUTLINE"
            Call SetOutLine(precno, va)
        Case "OTHER_OUTLINE"
            Call SetOutLine(precno, va)
        Case "ROUTE_OUTLINE"
            Call SetOutLine(precno, va)
        Case "PLACE_OUTLINE"
            Call SetOutLine(precno, va)
        Case "ROUTE_KEEPOUT"
            Call SetKeepOut(precno, va)
        Case "VIA_KEEPOUT"
            Call SetKeepOut(precno, va)
        Case "PLACE_KEEPOUT"
            Call SetKeepOut(precno, va)
        Case "PLACE_REGION"
            Call SetRegion(precno, va)
        Case "DRILLED_HOLES"
            Call SetHole(precno, va)
        Case "NOTES"
            Call SetNote(precno, va)
        Case "PLACEMENT"
            Call SetPlacement(precno, va)
        End Select
    End If
End Sub

'----------------------------------------

Private Sub SetHeader(no As Integer, va As Variant)
    Select Case no
    Case 1
        Set pcitem = New CIDFFile
        Call pfiles.Add(pcitem)
        pcfile = pfiles.Count
    Case Else
        Call pfiles(pcfile).SetData(no, va)
    End Select
End Sub

Private Sub SetOutLine(no As Integer, va As Variant)
    Dim s As String
    Select Case no
    Case 1
        Set pcitem = New CIDFItem
        pcitem.Section = psect
        Call pcitem.SetData(no, va)
    Case 2
        Call pcitem.SetOutLineData(va)
        s = pcitem.name & pitems.Count
        If pitems.Exists(s) Then pitems.Remove s
        Call pitems.Add(s, pcitem)
    Case Else
        Call pcitem.SetData(no, va)
    End Select
End Sub

Private Sub SetKeepOut(no As Integer, va As Variant)
    Dim s As String
    Select Case no
    Case 1
        Set pcitem = New CIDFItem
        pcitem.Section = psect
        Call pcitem.SetData(no, va)
    Case 2
        Call pcitem.SetKeepOutData(va)
        s = pcitem.name & pitems.Count
        If pitems.Exists(s) Then pitems.Remove s
        Call pitems.Add(s, pcitem)
    Case Else
        Call pcitem.SetData(no, va)
    End Select
End Sub

Private Sub SetRegion(no As Integer, va As Variant)
    Dim s As String
    Select Case no
    Case 1
        Set pcitem = New CIDFItem
        pcitem.Section = psect
        Call pcitem.SetData(no, va)
    Case 2
        Call pcitem.SetRegionData(no, va)
        s = pcitem.name & pitems.Count
        If pitems.Exists(s) Then pitems.Remove s
        Call pitems.Add(s, pcitem)
    Case Else
        Call pcitem.SetData(no, va)
    End Select
End Sub

Private Sub SetHole(no As Integer, va As Variant)
    Select Case no
    Case 1
    Case Else
        Set pcitem = New CIDFHole
        Call pcitem.SetData(va)
        Call pholes.Add(pcitem)
    End Select
End Sub

Private Sub SetNote(no As Integer, va As Variant)
    Select Case no
    Case 1
    Case Else
        Set pcitem = New CIDFNote
        Call pcitem.SetData(va)
        Call pnotes.Add(pcitem)
    End Select
End Sub

Private Sub SetPlacement(no As Integer, va As Variant)
    Select Case no
    Case 1
    Case 2
        Set pcitem = New CIDFPart
        Call pcitem.SetData(no, va)
    Case Else
        Call pcitem.SetData(no, va)
        Call pparts.Add(pcitem)
        precno = 1
    End Select
End Sub

Private Sub SetSymbol(no As Integer, va As Variant)
    Dim s As String
    Select Case no
    Case 1
        Set pcitem = New CIDFItem
        pcitem.Section = psect
        Call pcitem.SetData(no, va)
    Case 2
        Call pcitem.SetData(no, va)
        s = pcitem.name
        If plib.Exists(s) Then plib.Remove s
        Call plib.Add(s, pcitem)
    Case Else
        If va(0) = "PROP" Then
            'TODO: PROP を設定
        Else
            Call pcitem.SetData(no, va)
        End If
    End Select
End Sub

'----------------------------------------
'読み込み
'----------------------------------------

Private Function Readline(line As String) As Boolean
    Dim re As Object
    Set re = regex("(\""[^""]*\""|\S+)+")
    With fso.GetFile(path).OpenAsTextStream
        Do Until .AtEndOfStream = True
            Dim mc As Object
            Set mc = re.Execute(.Readline)
            If mc.Count > 0 Then Parse mc
        Loop
        .Close
    End With
    Readline = True
End Function

Public Function ReadRange(ra As Range) As Boolean
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Set ra = ws.UsedRange
    Dim arr As Variant
    arr = ra.Range(ra, ra.Offset(0, 8)).Value
    Dim r As Long
    Dim c As Integer
    Dim vs As Variant
    vs = Array("", "", "", "", "", "", "", "", "")
    For r = LBound(arr, 1) To UBound(arr, 1)
        For c = 1 To 8
            vs(c - 1) = arr(r, c)
        Next c
        Parse vs
    Next r
    ReadRange = True
End Function

Public Function LoadSheet(Optional path As String, Optional emp As Boolean) As Boolean
    Dim ws As Worksheet
    Set ws = SelectSheet()
    If ws Is Nothing Then Exit Function
    LoadSheet = ReadRange(ws.UsedRange)
End Function

Public Function ReadFile(path As String) As Boolean
    Dim re As Object
    Set re = regex("(\""[^""]*\""|\S+)+")
    With fso.GetFile(path).OpenAsTextStream
        Do Until .AtEndOfStream = True
            Dim mc As Object
            Set mc = re.Execute(.Readline)
            If mc.Count > 0 Then
                Dim vs As Variant
                vs = Array("", "", "", "", "", "", "", "")
                Dim i As Integer
                For i = 0 To mc.Count - 1
                    vs(i) = mc(i)
                Next i
                Parse vs
            End If
        Loop
        .Close
    End With
    ReadFile = True
End Function

Public Function LoadFile(Optional path As String, Optional emp As Boolean) As Boolean
    If path = "" Then
        Dim s As String
        s = "IDFファイル;*.emn,すべてのファイル"
        If emp Then s = "IDFファイル;*.emp;*.emn,すべてのファイル"
        path = SelectFile(path, "IDFファイル選択", s)
        If path = "" Then Exit Function
    End If
    LoadFile = ReadFile(path)
End Function

'----------------------------------------
'描画
'----------------------------------------

Private Function IDs(dic As Object) As Variant()
    Dim ks As Dictionary
    Set ks = New Dictionary
    Dim obj As Variant
    For Each obj In dic
        If Not ks.Exists(obj.id) Then ks.Add obj.id, 0
    Next obj
    IDs = ks.Keys
End Function

Public Function Draw(ws As Worksheet, _
        x As Double, y As Double, _
        a As Double, g As Double) As String
    g_px = x
    g_py = y
    g_pa = a
    g_pg = g
    '
    Dim ns As Collection
    Set ns = New Collection
    Dim v As Variant
    Dim sh As Object
    '
    'items
    Set sh = DrawItem(ws)
    If Not sh Is Nothing Then ns.Add sh.name
    '
    'holes
    For Each v In IDs(pholes)
        Set sh = DrawHole(ws, "" & v)
        If Not sh Is Nothing Then ns.Add sh.name
    Next v
    '
    'parts
    Set sh = DrawPart(ws, "BOTTOM")
    If Not sh Is Nothing Then ns.Add sh.name
    Set sh = DrawPart(ws, "BOTH")
    If Not sh Is Nothing Then ns.Add sh.name
    Set sh = DrawPart(ws, "TOP")
    If Not sh Is Nothing Then ns.Add sh.name
    '
    'notes
    Set sh = DrawNote(ws)
    If Not sh Is Nothing Then ns.Add sh.name
    '
    Dim s As String
    s = pfiles(1).name
    If ns.Count > 1 Then
        Set sh = ws.Shapes.Range(ToArray(ns)).Group
        s = pfiles(1).name & " " & sh.id
        sh.name = s
        SetShapeSetting
    End If
    Draw = s
End Function

Private Sub SetSetting(obj As Object, sel As String)
    Select Case sel
    Case "ELECTRICAL"
    Case "MECHANICAL"
    Case "HEADER"
    Case "BOARD_OUTLINE"
        obj.Fill.Visible = msoTrue
        obj.Fill.ForeColor.RGB = RGB(0, 255, 0)
        obj.Fill.Transparency = 0
    Case "PANEL_OUTLINE"
        obj.Fill.Visible = msoTrue
        obj.Fill.ForeColor.RGB = RGB(255, 0, 0)
        obj.Fill.Transparency = 0
    Case "OTHER_OUTLINE"
        obj.Fill.Visible = msoTrue
        obj.Fill.ForeColor.RGB = RGB(0, 0, 255)
        obj.Fill.Transparency = 0
        obj.Fill.Visible = msoFalse
    Case "ROUTE_OUTLINE"
        obj.Fill.Visible = msoTrue
        obj.Fill.ForeColor.RGB = RGB(255, 255, 0)
        obj.Fill.Transparency = 0
        obj.Fill.Visible = msoFalse
    Case "PLACE_OUTLINE"
        obj.Fill.Visible = msoFalse
        obj.Fill.ForeColor.RGB = RGB(0, 255, 255)
        obj.Fill.Transparency = 0
        obj.Fill.Visible = msoFalse
    Case "ROUTE_KEEPOUT"
        obj.Fill.Visible = msoTrue
        obj.Fill.Transparency = 0.8
        obj.Fill.Visible = msoFalse
    Case "VIA_KEEPOUT"
        obj.Fill.Visible = msoTrue
        obj.Fill.Transparency = 0.8
        obj.Fill.Visible = msoFalse
    Case "PLACE_KEEPOUT"
        obj.Fill.Visible = msoTrue
        obj.Fill.Transparency = 0.8
        obj.Fill.Visible = msoFalse
    Case "PLACE_REGION"
        obj.Fill.Visible = msoTrue
        obj.Fill.Transparency = 0.8
        obj.Fill.Visible = msoFalse
    Case "DRILLED_HOLES"
    Case "NOTES"
    Case "PLACEMENT"
    Case Else
    End Select
End Sub

Private Function DrawItem(ws As Worksheet) As Shape
    Dim ns As Collection
    Set ns = New Collection
    Dim sh As Shape
    Dim s As String
    Dim v As Variant
    For Each v In pitems.Items
        Set sh = v.Draw(ws, g_px, g_py, g_pa, g_pg)
        If Not sh Is Nothing Then
            Call SetSetting(sh, v.Section)
            ns.Add sh.name
        End If
    Next v
    If ns.Count = 0 Then Exit Function
    If ns.Count > 1 Then Set sh = ws.Shapes.Range(ToArray(ns)).Group
    s = "OUTLINE " & sh.id
    sh.name = s
    SetShapeSetting
    Set DrawItem = sh
End Function

Private Function DrawHole(ws As Worksheet, id As String) As Shape
    Dim ns As Collection
    Set ns = New Collection
    Dim sh As Shape
    Dim s As String
    Dim hole As CIDFHole
    For Each hole In pholes
        If hole.id = id Then
            Set sh = hole.Draw(ws, g_px, g_py, g_pa, g_pg)
            If Not sh Is Nothing Then ns.Add sh.name
        End If
    Next hole
    If ns.Count = 0 Then Exit Function
    If ns.Count > 1 Then Set sh = ws.Shapes.Range(ToArray(ns)).Group
    s = "DRILLED_HOLES_" & id & " " & sh.id
    sh.name = s
    SetShapeSetting
    Set DrawHole = sh
End Function

Private Function DrawPart(ws As Worksheet, id As String) As Shape
    Dim ns As Collection
    Set ns = New Collection
    Dim sh As Shape
    Dim s As String
    Dim part As CIDFPart
    For Each part In pparts
        If part.side = id Then
            s = part.name
            Dim sym As Variant
            Set sym = Nothing
            Do While sym Is Nothing
                If plib.Exists(s) Then
                    Set sym = plib(s)
                    Set sh = part.Draw(ws, g_px, g_py, g_pa, g_pg, sym)
                ElseIf pitems.Exists(s) Then
                    Set sym = pitems(s)
                    Set sh = sym.Draw(ws, g_px, g_py, g_pa, g_pg, sym)
                ElseIf Not LoadSheet(emp:=True) Then
                    If Not LoadFile(emp:=True) Then Exit Function
                End If
            Loop
            If Not sh Is Nothing Then ns.Add sh.name
        End If
        Set sh = Nothing
    Next part
    If ns.Count = 0 Then Exit Function
    If ns.Count > 1 Then Set sh = ws.Shapes.Range(ToArray(ns)).Group
    s = "PARTS_" & id & " " & sh.id
    sh.name = s
    SetShapeSetting
    Set DrawPart = sh
End Function

Private Function DrawNote(ws As Worksheet) As Shape
    Dim ns As Collection
    Set ns = New Collection
    Dim sh As Shape
    Dim s As String
    Dim note As CIDFNote
    For Each note In pnotes
        Set sh = note.Draw(ws, g_px, g_py, g_pa, g_pg)
        If Not sh Is Nothing Then ns.Add sh.name
    Next note
    If ns.Count = 0 Then Exit Function
    If ns.Count > 1 Then Set sh = ws.Shapes.Range(ToArray(ns)).Group
    s = "NOTES " & sh.id
    sh.name = s
    SetShapeSetting
    Set DrawNote = sh
End Function
