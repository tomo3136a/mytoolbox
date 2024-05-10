Attribute VB_Name = "Common"
Option Explicit
Option Private Module

'----------------------------------------
'�I�u�W�F�N�g�Ăяo��
'----------------------------------------

'worksheet.function
Function wsf() As WorksheetFunction
    Set wsf = WorksheetFunction
End Function


'----------------------------------------
'���K�\��
'----------------------------------------

'regex
Function regex( _
        ptn As String, _
        Optional g As Boolean = True, _
        Optional ic As Boolean = True) As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = g
        .IgnoreCase = ic
        .Pattern = ptn
    End With
End Function

'������L��
Public Function re_test(s As String, ptn As String) As Boolean
    On Error Resume Next
    re_test = regex(ptn).test(s)
    On Error GoTo 0
End Function

'�����񒊏o
Public Function re_match(s As String, ptn As String, _
        Optional idx As Integer = 0) As String
    Dim mc As Object
    Set mc = regex(ptn).Execute(s)
    If idx < 0 Or idx >= mc.Count Then Exit Function
    re_match = mc(idx).Value
End Function

'������u������
Public Function re_replace(s As String, ptn As String, rep As String) As String
    re_replace = regex(ptn).Replace(s, rep)
End Function


'----------------------------------------
'�̈�̒l������擾
'----------------------------------------

Function StrRange(s As String) As String
    Dim ra As Range
    Set ra = Range(s)
    If ra.Count = 1 Then
        StrRange = s
        Exit Function
    End If
    Dim n As Integer
    n = ra.Column + ra.Columns.Count - 1
    Dim ce As Range
    Dim ss As String
    For Each ce In ra
        ss = ss & Chr(34) & ce.Value & Chr(34)
        If n = ce.Column Then
            ss = ss & vbLf
        Else
            ss = ss & ","
        End If
    Next ce
    StrRange = Left(ss, Len(ss) - 1)
End Function


'----------------------------------------
'�R���N�V������z��ɕϊ�
'----------------------------------------

Function ToArray(col As Collection) As Variant()
    Dim arr() As Variant
    ReDim arr(0 To col.Count - 1)
    Dim i As Integer
    For i = 1 To col.Count
        arr(i - 1) = col.Item(i)
    Next i
    ToArray = arr
End Function


'----------------------------------------
'�̈�̊p�擾
'----------------------------------------

Function LeftTop(ra As Range) As Range
    Set LeftTop = ra.Cells(1, 1)
End Function

Function RightTop(ra As Range) As Range
    Set RightTop = ra.Cells(1, ra.Columns.Count)
End Function

Function LeftBottom(ra As Range) As Range
    Set LeftBottom = ra.Cells(ra.Rows.Count, 1)
End Function

Function RightBottom(ra As Range) As Range
    Set RightBottom = ra.Cells(ra.Rows.Count, ra.Columns.Count)
End Function


'----------------------------------------
'�Z���T��
'----------------------------------------

'��[�擾
Public Function FarTop(ra As Range, Optional margin As Integer) As Range
    If margin < 1 Then margin = 1
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Row
    Dim cnt As Integer
    Dim rs As Range
    Set rs = ce
    Do While ce.Row > p And cnt < margin
        If ce.Offset(-1).Value = "" Then
            Set ce = ce.Offset(-1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlUp)
            Set rs = ce
            cnt = 0
        End If
    Loop
    Set FarTop = rs
End Function

'���[�擾
Public Function FarBottom(ra As Range, Optional margin As Integer) As Range
    If margin < 1 Then margin = 1
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Row
    p = p + ra.Worksheet.UsedRange.Rows.Count
    Dim cnt As Integer
    Dim re As Range
    Set re = ce
    Do While ce.Row < p And cnt < margin
        If ce.Offset(1).Value = "" Then
            Set ce = ce.Offset(1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlDown)
            Set re = ce
            cnt = 0
        End If
    Loop
    Set FarBottom = re
End Function

'���[�擾
Public Function FarLeft(ra As Range, Optional margin As Integer) As Range
    If margin < 1 Then margin = 1
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Column
    Dim cnt As Integer
    Dim rs As Range
    Set rs = ce
    Do While ce.Column > p And cnt < margin
        If ce.Offset(0, -1).Value = "" Then
            Set ce = ce.Offset(0, -1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlToLeft)
            Set rs = ce
            cnt = 0
        End If
    Loop
    Set FarLeft = rs
End Function

'�E�[�擾
Public Function FarRight(ra As Range, Optional margin As Integer) As Range
    If margin < 1 Then margin = 1
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Column
    p = p + ra.Worksheet.UsedRange.Columns.Count
    Dim cnt As Integer
    Dim re As Range
    Set re = ce
   Do While ce.Column < p And cnt < margin
        If ce.Offset(0, 1).Value = "" Then
            Set ce = ce.Offset(0, 1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlToRight)
            Set re = ce
            cnt = 0
        End If
    Loop
    Set FarRight = re
End Function

'�����񂪈�v����cell��T��
Public Function FindCell(s As String, ra As Range) As Range
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Dim r As Long
    Dim c As Integer
    For r = ra.Row To ws.UsedRange.Rows.Count
        For c = ra.Column To ws.UsedRange.Columns.Count
            Dim ce As Range
            Set ce = ws.Cells(r, c)
            If ce.Value = s Then
                Set FindCell = ce
                Exit Function
            End If
        Next c
    Next r
End Function

'�u�����N���X�L�b�v����
Public Function SkipBlankCell(ra As Range) As Range
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Dim r As Long
    Dim c As Integer
    For r = ra.Row To ws.UsedRange.Rows.Count
        For c = ra.Column To ws.UsedRange.Columns.Count
            Dim ce As Range
            Set ce = ws.Cells(r, c)
            If ce.Value <> "" Then
                Set SkipBlankCell = ce
                Exit Function
            End If
        Next c
    Next r
End Function


'----------------------------------------
'�p�X����
'----------------------------------------

'filesystemobject
Function fso() As Object
    Static obj As Object
    If obj Is Nothing Then
        Set obj = CreateObject("Scripting.FileSystemObject")
    End If
    Set fso = obj
End Function

'��{���擾
'  �p�X�r���A�g���q�r��
'  �������r��
Function BaseName(s As String) As String
    Dim re As Object
    Set re = regex("[\(�i]\d+[\)�j]|\s*-\s*�R�s�[")
    BaseName = re.Replace(fso.GetBaseName(s), "")
End Function

Function CanonicalPath(path As String)
    Dim arr As Variant
    arr = Array("Box", "OneDrive", "LOCALAPPDATA", "APPDATA", "USERPROFILE")
    '
    Dim p As String
    p = path
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        Dim name As String
        name = arr(i)
        Dim Base As String
        Base = Environ(name)
        If Mid(p & "\", 1, Len(Base & "\")) = Base & "\" Then
            p = Replace(p, Base, "(" & name & ")", compare:=vbTextCompare)
            Exit For
        End If
    Next i
    '
    CanonicalPath = p
End Function

Function EnvironmentPath(path As String)
    EnvironmentPath = re_replace(path, "^\((\w+)\)", "%$1%")
End Function


'----------------------------------------
'�V�[�g������
'----------------------------------------

'�V�[�g���L���̃`�F�b�N
Function HasSheetName(wb As Workbook, name As String) As Boolean
    Dim i As Integer
    For i = 1 To wb.Worksheets.Count
        If wb.Worksheets(i).name = name Then
            HasSheetName = True
            Exit Function
        End If
    Next i
End Function

'�d�����Ȃ��V�[�g���擾
Function UniqueSheetName(wb As Workbook, name As String) As String
    Dim i As Integer: i = 1
    Dim s As String: s = name
    Do While HasSheetName(wb, s)
        s = name & " (" & i & ")"
        i = i + 1
    Loop
    UniqueSheetName = s
End Function

'�V�[�g�����l�[���_�C�A���O
Sub SheetRenameDialog()
    CommandBars.ExecuteMso "SheetRename"
End Sub


'----------------------------------------
'�V�[�g�v���p�e�B����
'----------------------------------------

'�V�[�g�v���p�e�B�����X�g���擾
Function GetSheetPropertyNames(ws As Worksheet) As String()
    Dim lst() As String
    ReDim Preserve lst(ws.CustomProperties.Count)
    Dim i As Integer
    For i = 1 To ws.CustomProperties.Count
        lst(i) = ws.CustomProperties(i).name
    Next i
    GetSheetPropertyNames = lst
End Function

'�V�[�g�v���p�e�B������ԍ��擾
Function SheetPropertyIndex(ws As Worksheet, name As String) As Integer
    Dim i As Integer
    For i = 1 To ws.CustomProperties.Count
        If ws.CustomProperties(i).name = name Then
            SheetPropertyIndex = i
            Exit Function
        End If
    Next i
End Function

'�V�[�g�v���p�e�B������v���p�e�B�擾
Function GetSheetProperty(ws As Worksheet, name As String) As CustomProperty
    Dim idx As Integer
    idx = SheetPropertyIndex(ws, name)
    If idx > 0 Then
        Set GetSheetProperty = ws.CustomProperties(idx)
        Exit Function
    End If
    Set GetSheetProperty = ws.CustomProperties.Add(name, "")
End Function

'�V�[�g�v���p�e�B������l�擾
Function StrSheetProperty(ws As Worksheet, name As String) As String
    Dim idx As Integer
    idx = SheetPropertyIndex(ws, name)
    If idx > 0 Then StrSheetProperty = ws.CustomProperties(idx).Value
End Function


'----------------------------------------
'��ʃ`�����h�~
'----------------------------------------

Public Sub ScreenUpdateOff()
    '��ʃ`�����h�~���u
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Interactive = False
    Application.Cursor = xlWait
End Sub

Public Sub ScreenUpdateOn()
    '��ʃ`�����h�~���u����
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


'----------------------------------------
'�i�s�󋵕\��(status-bar)
'----------------------------------------

Sub ProgressStatusBar(Optional i As Long = 1, Optional cnt As Long = 1)
    Static tm_start As Double
    If i < 1 Then
        tm_start = Timer
        Application.StatusBar = "�i����(0%)"
        Exit Sub
    End If
    If i >= cnt Then
        Application.StatusBar = False
        Exit Sub
    End If
    Dim p As Double: p = i / cnt
    Dim s As String: s = "�i����(" & Int(p * 100) & "%)"
    s = s & " : " & ProgressBar(p)
    Dim tm As Double: tm = (Timer - tm_start) / p * (1 - p)
    Application.StatusBar = s & " : �c��" & Int(tm) & "�b"
End Sub

Private Function ProgressBar(p As Double) As String
    If p < 0.2 Then
        ProgressBar = "����������"
    ElseIf p < 0.4 Then
        ProgressBar = "����������"
    ElseIf p < 0.6 Then
        ProgressBar = "����������"
    ElseIf p < 0.8 Then
        ProgressBar = "����������"
    ElseIf p < 1 Then
        ProgressBar = "����������"
    Else
        ProgressBar = "����������"
    End If
End Function


