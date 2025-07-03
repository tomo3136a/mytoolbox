Attribute VB_Name = "Template"
'==================================
'�e���v���[�g�@�\
'==================================

Option Explicit
Option Private Module

' mode=1: �V�[�g�ǉ�
'      2: �V�[�g�o�^
'      3: �V�[�g�폜
'      4: �e�[�u���쐬
'      5: �e�[�u���ǂݍ���
'      6: �e�[�u���X�V
'      7: �w�b�_�[�V�[�g�擾
'      8: addins: toggle visible/hidden sheets
Sub TemplateProc(mode As Long, Optional id As Long)
    Select Case mode
    Case 1: CopyAddinSheet
    Case 2: UpdateAddinSheet ActiveSheet
    Case 3: RemoveAddinSheet
    
    Case 4: AddTable
    Case 5: LoadCsvTable
    Case 6: 'BuildAddin
    
    Case 7: CopyHeaderAddinSheet
    Case 8: ToggleAddinBook
    End Select
End Sub

'----------------------------------
'�e���v���[�g�V�[�g�@�\
'----------------------------------

'�A�h�C���u�b�N����e���v���[�g�V�[�g�𕡐�
Private Sub CopyAddinSheet(Optional src As String, Optional dst As String)
    Dim ws As Worksheet
    If src <> "" Then Set ws = SearchName(ThisWorkbook.Sheets, src)
    If ws Is Nothing Then Set ws = SelectSheet(ThisWorkbook, "^[^#]")
    If ws Is Nothing Then Exit Sub
    '
    Dim s As String
    s = dst
    If s = "" Then
        Dim msg As String
        msg = "�쐬����V�[�g�������Ă��������B"
        s = InputBox(msg, app_name, ws.name)
        If StrPtr(s) = 0 Then Exit Sub
        If s = "" Then s = ws.name
    End If
    '
    s = UniqueSheetName(ActiveWorkbook, s)
    ws.Copy After:=ActiveSheet
    ActiveSheet.name = s
End Sub

'�A�h�C���u�b�N�փe���v���[�g�V�[�g�X�V
Private Sub UpdateAddinSheet(ws As Worksheet)
    Dim ws2 As Worksheet
    For Each ws2 In ThisWorkbook.Sheets
        If ws2.name = ws.name Then Exit For
    Next ws2
    '
    '���O���o�^����Ă��Ȃ���ΐV�K�ɒǉ�
    If ws2 Is Nothing Then
        ScreenUpdateOff
        ThisWorkbook.IsAddin = False
        ws.Copy After:=ThisWorkbook.Sheets(1)
        ThisWorkbook.IsAddin = True
        ScreenUpdateOn
        Exit Sub
    End If
    '
    '�㏑���m�F
    Dim msg As String
    msg = "�����̓o�^������܂��B" & vbLf & ws.name
    msg = msg & vbLf & "�㏑�����܂����B"
    Dim res As VbMsgBoxResult
    res = MsgBox(msg, vbYesNo Or vbDefaultButton2, app_name)
    If res = vbNo Then Exit Sub
    '
    '���O���o�^����Ă���Ώ㏑���R�s�[
    Dim old As Range
    Set old = Selection
    ws.Cells.Copy
    With ThisWorkbook.Sheets(ws.name)
        .Paste .Cells(1, 1)
    End With
    Application.CutCopyMode = False
    old.Select
End Sub

'�A�h�C���u�b�N����e���v���[�g�V�[�g���폜
Private Sub RemoveAddinSheet()
    Dim ws As Worksheet
    Set ws = SelectSheet(ThisWorkbook)
    If ws Is Nothing Then Exit Sub
    '
    Dim msg As String
    msg = "�e���v���[�g���폜���܂����B" & vbLf & ws.name
    Dim res As VbMsgBoxResult
    res = MsgBox(msg, vbYesNo Or vbDefaultButton2, app_name)
    If res = vbNo Then Exit Sub
    '
    Dim f As Boolean
    f = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = f
End Sub

'----------------------------------
'�e���v���[�g�e�[�u���@�\
'----------------------------------

'�w�b�_�[�V�[�g�擾
Private Sub CopyHeaderAddinSheet()
    ActivateConfigSheet "#header"
End Sub

'�e�[�u���쐬�@�\
Private Sub AddTable()
    '�w�b�_��`�V�[�g�擾
    Dim ws As Worksheet
    Set ws = SearchName(ActiveWorkbook.Sheets, "#header")
    If ws Is Nothing Then Set ws = SearchName(ThisWorkbook.Sheets, "#header")
    If ws Is Nothing Then Exit Sub
    '
    '�e���v���[�g�J�n�ʒu�擾
    Dim ra As Range
    Set ra = SectionRange(ws.UsedRange.Columns(1))
    If ra Is Nothing Then Exit Sub
    Set ra = SelectCell(ra)
    If ra Is Nothing Then Exit Sub
    If ra.Count <> 1 Then Exit Sub
    Set ra = ra.Offset(0, 2)
    '
    ScreenUpdateOff
    '
    Dim c As Long
    Dim cm As Long
    Dim rb As Range, rc As Range
    Set rb = ActiveCell
    Dim t As String
    t = Left(UCase(ra.Offset(0, -2)), 1)
    Do Until t = ""
        cm = ra.End(xlToRight).Column - ra.Column + 1
        Select Case t
        Case "H"
            For c = cm To 1 Step -1
                Set rc = rb(1, c)
                Set rc = rc.EntireColumn
                rc.Hidden = ra(1, c)
            Next c
        Case "D"
            For c = cm To 1 Step -1
                Set rc = rb(1, c)
                Set rc = rc.EntireColumn
                If ra(1, c) Then rc.Delete
            Next c
        Case Else
            Set rc = ra.Resize(1, cm)
            rc.Copy Destination:=rb
            Set rb = rb.Offset(1)
        End Select
        Set ra = ra.Offset(1)
        t = Left(UCase(ra.Offset(0, -2)), 1)
        If t = "[" Then Exit Do
    Loop
    '
    ScreenUpdateOn
End Sub

'�A�h�C���u�b�N�\���g�O��
Private Sub ToggleAddinBook()
    If ThisWorkbook.IsAddin Then
        ThisWorkbook.IsAddin = False
        ThisWorkbook.Activate
    Else
        ThisWorkbook.IsAddin = True
        ThisWorkbook.Save
    End If
End Sub

'----------------------------------
'���̑�
'----------------------------------


'�e�[�u���ǂݍ��݋@�\
Private Sub LoadCsvTable(Optional path As String, Optional utf8 As Boolean)
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


Private Sub LoadTable()
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
