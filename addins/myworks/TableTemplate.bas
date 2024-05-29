Attribute VB_Name = "TableTemplate"
'==================================
'�e���v���[�g�@�\
'==================================

Option Explicit
Option Private Module

Sub TemplateMenu(id As Integer)
    Select Case id
    Case 1 '�V�[�g����
        CopyAddinSheet
    Case 2 '�V�[�g�X�V
        UpdateAddinSheet ActiveSheet
    Case 3 'addins: toggle visible/hidden sheets
        ToggleAddin
    Case 4 '�e�[�u���쐬
        AddTable
    Case 5 '�e�[�u���ǂݍ���
        LoadCsvTable
    Case 6 '�e�X�g�@�\
        'TestTable
    Case 7
        'BuildAddin
    End Select
End Sub

'----------------------------------
'�@�\
'----------------------------------

'�A�h�C���u�b�N����e���v���[�g�V�[�g�𕡐�
Private Sub CopyAddinSheet()
    Dim ws As Worksheet
    Set ws = SelectSheet(ThisWorkbook, "^[^#]")
    If ws Is Nothing Then Exit Sub
    ws.Copy after:=ActiveSheet
End Sub

'�A�h�C���u�b�N�̃e���v���[�g�V�[�g�X�V
Private Function UpdateAddinSheet(ws As Worksheet)
    Dim asu As Boolean
    asu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim ws2 As Worksheet
    For Each ws2 In ThisWorkbook.Sheets
        If ws2.name = ws.name Then Exit For
    Next ws2
    If ws2 Is Nothing Then
        ThisWorkbook.IsAddin = False
        ws.Copy after:=ThisWorkbook.Sheets(1)
        ThisWorkbook.IsAddin = True
    Else
        Dim old As Range
        Set old = Selection
        ws.Cells.Select
        Selection.Copy
        Set ws2 = ThisWorkbook.Sheets(ws.name)
        ws2.Paste ws2.Cells(1, 1)
        Application.CutCopyMode = False
        old.Select
    End If
    '
    Application.ScreenUpdating = asu
End Function

'�A�h�C���u�b�N�\���g�O��
Private Sub ToggleAddin()
    If ThisWorkbook.IsAddin Then
        ThisWorkbook.IsAddin = False
        ThisWorkbook.Activate
    Else
        ThisWorkbook.IsAddin = True
        ThisWorkbook.Save
    End If
End Sub

'�e�[�u���쐬�@�\
Private Sub AddTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Item("#header")
    If ws Is Nothing Then Exit Sub
    '
    Dim ra As Range
    Set ra = ws.UsedRange
    Set ra = ws.Range(ra.Cells(1, 1), ra.Cells(ra.Rows.Count, 1))
    Set ra = SelectCell(SectionRange(ra))
    If ra Is Nothing Then Exit Sub
    Set ra = ra.Offset(0, 1)
    '
    Dim asu As Boolean
    asu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim cm As Integer
    cm = ws.UsedRange.Column + ws.UsedRange.Columns.Count
    Dim c As Integer
    '
    Dim rm As Integer
    rm = ra.Row
    If ra.Offset(1).Value <> "" Then rm = ra.End(xlDown).Row
    Dim r As Integer
    '
    Dim r2 As Integer
    r2 = ra.Row
    For r = r2 To rm
        If Not IsNumeric(ws.Cells(r, 2).Value) Then Exit For
        If ws.Cells(r, 2).Value < 2 Then r2 = r
        Dim c2 As Integer
        c2 = ws.Cells(r, cm).End(xlToLeft).Column
        If c2 > c Then c = c2
    Next r
    r = r - 1
    Set ra = ra.Offset(0, 1)
    '
    
    Dim tbl As Range
    Set tbl = ws.Range(ra, ws.Cells(r, c))
    tbl.Copy Destination:=Selection.Cells(1, 1)
    Selection.Offset(r2 - ra.Row + 1).Select
    '
    Application.ScreenUpdating = asu
End Sub

'�e�[�u���ǂݍ��݋@�\
Sub LoadCsvTable(Optional path As String, Optional utf8 As Boolean)
    If path = "" Then path = SelectCsvFile()
    If path = "" Then Exit Sub
    '
    Dim asu As Boolean
    asu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim ra As Range
    Set ra = ActiveCell
    '
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    '
    Dim arrDataType(255) As Long
    Dim i As Integer
    For i = 0 To 255
        arrDataType(i) = xlTextFormat
    Next i
    '
    With ws.QueryTables.Add( _
            Connection:="TEXT;" + path, _
            Destination:=ra)
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        If utf8 Then
            .TextFilePlatform = 65001 'UTF-8
        Else
            .TextFilePlatform = 932   'SJIS
        End If
        .TextFileStartRow = 1
        .TextFileColumnDataTypes = arrDataType
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    '
    Application.ScreenUpdating = asu
End Sub

'----------------------------------
'���̑�
'----------------------------------

Sub LoadTable()
    Dim Title As String
    Title = "�f�[�^"
    '
    Dim arrDataType(255) As Long
    Dim i As Integer
    For i = 0 To 255
        arrDataType(i) = xlTextFormat
    Next i
    '
    Dim asu As Boolean
    asu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim v As Variant
    For Each v In SelectFiles(, , "CSV�t�@�C��")
        Dim ws As Worksheet
        Set ws = Sheets.Add
        ws.Cells(1, 1) = Title
        ws.Cells(2, 1) = fso.GetBaseName(v)
        ws.name = UniqueSheetName(ws.Parent, fso.GetBaseName(v))
        With ws.QueryTables.Add( _
                Connection:="TEXT;" + v, _
                Destination:=ws.Cells(3, 1))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            '.TextFilePlatform = 932   'SJIS
            .TextFilePlatform = 65001 'UTF-8
            .TextFileStartRow = 1
            .TextFileColumnDataTypes = arrDataType
            .Refresh BackgroundQuery:=False
            '.name = "tmp_tbl"
            .Delete
        End With
        'Dim n As name
        'For Each n In ActiveWorkbook.Names
        '    If n.name = .name & "!" & "tmp_tbl" Then
        '        n.Delete
        '    End If
        'Next n
    Next v
    '
    Application.ScreenUpdating = asu
End Sub

