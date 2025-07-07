Attribute VB_Name = "Template"
'==================================
'�e���v���[�g�@�\
'==================================

Option Explicit
Option Private Module

'�e���v���[�g�@�\
' mode=1: �V�[�g�ǉ�
'      2: �V�[�g�o�^
'      3: �V�[�g�폜
'      4: �e�[�u���ǉ�
'      4: �e�[�u���o�^
'      6: �e�[�u���폜
'      7: �w�b�_�[�V�[�g�擾
'      8: �e�[�u���ǂݍ���
'      9: addins: toggle visible/hidden sheets
Sub TemplateProc(mode As Long, Optional id As Long)
    Select Case mode
    Case 1: CopyTemplateSheet
    Case 2: UpdateTemplateSheet ActiveSheet
    Case 3: RemoveTemplateSheet
    
    Case 4: CopyTemplateTable
    Case 5: UpdateTemplateTable Selection
    Case 6: RemoveTemplateTable
    Case 7: CopyHeaderSheet
    
    Case 8: LoadCsvTable
    Case 9: ToggleAddinBook
    End Select
End Sub

'----------------------------------
'�e���v���[�g�V�[�g�@�\
'----------------------------------

'�e���v���[�g�V�[�g����
Private Sub CopyTemplateSheet( _
    Optional src As String, Optional dst As String)
    Dim ws As Worksheet
    If src <> "" Then Set ws = TakeByName(ThisWorkbook.Sheets, src)
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

'�e���v���[�g�V�[�g�X�V
Private Sub UpdateTemplateSheet(ws As Worksheet)
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

'�e���v���[�g�V�[�g�폜
Private Sub RemoveTemplateSheet( _
    Optional sname As String, Optional bforce As Boolean)
    Dim ws As Worksheet
    Set ws = TakeByName(ThisWorkbook.Worksheets, sname)
    If ws Is Nothing Then Set ws = SelectSheet(ThisWorkbook)
    If ws Is Nothing Then Exit Sub
    '
    '�폜�m�F
    Dim res As VbMsgBoxResult
    Dim msg As String
    If Not bforce Then
        msg = "�e���v���[�g���폜���܂����B" & vbLf & ws.name
        res = MsgBox(msg, vbYesNo Or vbDefaultButton2, app_name)
        If res = vbNo Then Exit Sub
    End If
    '
    '�V�[�g�폜
    Dim f As Boolean
    f = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = f
End Sub

'----------------------------------
'�e���v���[�g�e�[�u���@�\
'----------------------------------

'�e���v���[�g�e�[�u������
Private Sub CopyTemplateTable( _
        Optional tname As String, _
        Optional bforce As Boolean)
    '
    '�e���v���[�g�e�[�u���擾
    Dim ra As Range, rb As Range
    Set ra = GetTemplateTable(tname)
    Set rb = GetTemplateTableRange(ra, eol:=True)
    If rb Is Nothing Then Exit Sub
    '
    '�o�͐�擾
    Dim rc As Range
    Set rc = ActiveCell
    '
    '�e�[�u���R�s�[
    ScreenUpdateOff
    rb.Copy Destination:=rc
    If Left(LeftBottom(rb).Offset(1, -2), 1) <> "#" Then
        ScreenUpdateOn
        Exit Sub
    End If
    '
    '�e�[�u���T�C�Y�擾
    Dim rm As Long, cm As Long
    rm = rb.Rows.Count
    cm = rb.Columns.Count
    '
    '���쒊�o
    Dim cmd As Variant, arr As Variant
    Set rb = rb(1, 1).Offset(rm)
    Set rb = rb.Resize(SectionRowCount(rb.Offset(0, -2)))
    cmd = rb.Offset(0, -2).Resize(, 1).Value
    arr = rb.Value
    '
    '���쏈��
    Dim rd As Range
    Dim r As Long, c As Long, i As Long
    For r = 1 To UBound(cmd, 1)
        Select Case LCase(Trim(cmd(r, 1)))
        Case "#continue"
            i = CLng(arr(r, 1)) + 1
            If i > 1 Then
                Set rb = rc(rm, 1).Resize(1, cm)
                Set rd = rc(rm, 1).Resize(i, cm)
                rb.AutoFill Destination:=rd, Type:=xlFillDefault
                rm = rm + i
            End If
        Case "#hide"
            For c = cm To 1 Step -1
                If arr(r, c) Then
                    rd(1, c).EntireColumn.Hidden = True
                End If
            Next c
        Case "#delete"
            For c = cm To 1 Step -1
                If arr(r, c) Then
                    Set rc = rd(1, c).Resize(rm, 1)
                    rc.Delete Shift:=xlToLeft
                    cm = cm - 1
                End If
            Next c
        End Select
    Next r
    '
    ScreenUpdateOn
End Sub

'�e���v���[�g�e�[�u���X�V
Private Sub UpdateTemplateTable( _
        ra As Range, _
        Optional ByVal tname As String)
    If ra Is Nothing Then Exit Sub
    '
    '���O�擾
    Dim rb As Range
    If tname = "" Then
        tname = InputBox("���O����͂��Ă��������B", app_name)
        tname = Trim(Replace(Replace(tname, "[", ""), "]", ""))
        If tname = "" Then
            Set rb = GetTemplateTable
            If rb Is Nothing Then Exit Sub
            tname = rb.Offset(0, -2)
            tname = Mid(tname, 2, Len(tname) - 2)
        End If
    End If
    '
    '�e���v���[�g�e�[�u���擾
    Set rb = GetTemplateTable(tname)
    '
    '�e�[�u���o�^
    Dim ws As Worksheet
    Set ws = ConfigSheet("#table", True)
    Dim rc As Range
    Set rc = ws.UsedRange
    If rc Is Nothing Then Set rc = ws.Cells(1, 1)
    Set rc = LeftBottom(rc).Offset(2)
    rc.Value = "[" & tname & "]"
    Set rc = rc.Offset(, 2)
    ra.Copy Destination:=rc
    Set rc = rc.Offset(ra.Rows.Count)
    '
    '�f�[�^�s�����o�^
    Dim s As String
    s = InputBox("�J��Ԃ���������Γ��͂��Ă��������B", app_name)
    Dim i As Long
    i = CLng("0" & s)
    If i > 0 Then
        rc.Offset(0, -2) = "#continue"
        rc.Offset(0, 0) = i
        Set rc = rc.Offset(1)
    End If
    '
    '�e�[�u���폜
    If rb Is Nothing Then Exit Sub
    GetTemplateTableRange(rb).EntireRow.Delete
End Sub

'�e���v���[�g�e�[�u���폜
Private Sub RemoveTemplateTable( _
    Optional ByVal tname As String, Optional bforce As Boolean)
    '
    '�e���v���[�g�e�[�u���擾
    Dim ra As Range, rb As Range
    Set ra = GetTemplateTable(tname)
    Set rb = GetTemplateTableRange(ra)
    If rb Is Nothing Then Exit Sub
    '
    '�폜�m�F
    Dim res As VbMsgBoxResult
    Dim msg As String
    If Not bforce Then
        msg = "�e���v���[�g���폜���܂����B" & vbLf & tname
        res = MsgBox(msg, vbYesNo Or vbDefaultButton2, app_name)
        If res = vbNo Then Exit Sub
    End If
    '
    '�e�[�u���폜
    rb.EntireRow.Delete
End Sub

'�e���v���[�g�e�[�u���̊J�n�ʒu���擾
Private Function GetTemplateTable( _
    Optional tname As String) As Range
    '�w�b�_��`�V�[�g�擾
    Dim ws As Worksheet
    Set ws = ConfigSheet("#table")
    If ws Is Nothing Then Exit Function
    '
    '�e���v���[�g�J�n�ʒu�擾
    Dim ra As Range
    Set ra = SectionTags(ws.UsedRange.Columns(1))
    Set ra = SectionCell(ra, tname)
    If ra Is Nothing Then Exit Function
    '
    Set GetTemplateTable = ra.Offset(0, 2)
End Function

'�e���v���[�g�e�[�u���͈͂��擾
Private Function GetTemplateTableRange( _
    ra As Range, Optional eol As Boolean) As Range
    If ra Is Nothing Then Exit Function
    '
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    '
    '�e���v���[�g�s�񐔎擾
    Dim t As String
    Dim i As Long
    Dim r As Long, rm As Long
    Dim c As Long, cm As Long
    rm = SectionRowCount(ra.Offset(0, -2), eol)
    cm = ws.UsedRange.Columns.Count + ws.UsedRange.Column - ra.Column
    For r = 1 To rm
        i = ra(r, cm + 1).End(xlToLeft).Column - ra.Column + 1
        If c < i Then c = i
    Next r
    If r < rm Then rm = r
    cm = c
    If rm < 1 Or cm < 1 Then Exit Function
    Set GetTemplateTableRange = ra.Resize(rm, cm)
    Exit Function
End Function

'�w�b�_�[�V�[�g�擾
Private Sub CopyHeaderSheet()
    ActivateConfigSheet "#table"
End Sub

'----------------------------------
'�e���v���[�g�Ńt�@�C���Ǎ�
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

'�e�X�g�p
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

'----------------------------------
'���̑�
'----------------------------------

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

