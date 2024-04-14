VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIDFFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================
'file class
'========================================

Private ptype As String         'file type:     BOARD_FILE, PANEL_FILE, LIBRARY_FILE
Private fmtver As String        'IDF version:   1.0, 2.0, 3.0
Private pbuilder As String      'builder
Private pdate As String         'build date:    format = yyyy/mm/dd.hh:mm:ss
Private pver As String          'file version
Private pname As String         'file name
Private punits As String        'units:         MM, THOU(=0.0254mm)

'----------------------------------------

Public Sub SetData(no As Integer, va As Object)
    Select Case no
    Case 1
    Case 2
        ptype = UCase(va(0))
        fmtver = va(1)
        pbuilder = va(2)
        pdate = va(3)
        pver = va(4)
    Case 3
        pname = va(0)
        punits = UCase(va(1))
    End Select
End Sub

'----------------------------------------

Public Property Get Text() As String
    Dim s As String
    s = "File Type   :  " & ptype
    s = s & Chr(10) & "IDF vesion   : " & fmtver
    s = s & Chr(10) & "Source system: " & pbuilder
    s = s & Chr(10) & "Date         : " & pdate
    s = s & Chr(10) & "File version : " & pver
    s = s & Chr(10) & "Board name   : " & pname
    s = s & Chr(10) & "Units        : " & punits
    Select Case punit
    Case "MM"
        s = s + "(mm)"
    Case ""
        s = s + "(mils=1/1000inch)"
    End Select
    Text = s
End Property

'Public Property Get FileType() As String
'    FileType = ptype
'End Property

'Public Property Get TypeVersion() As String
'    TypeVersion = fmtver
'End Property

'Public Property Get Builder() As String
'    Builder = pbuilder
'End Property

'Public Property Get BuildDate() As String
'    FileType = pdate
'End Property

'Public Property Get Version() As String
'    Version = pver
'End Property

Public Property Get name() As String
    name = pname
End Property

'Public Property Get Units() As String
'    Units = punits
'End Property

