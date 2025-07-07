Attribute VB_Name = "BaseCfg"
'==================================
'����(�ݒ�V�[�g����)
'==================================

'----------------------------------------
'API:
'  �ݒ�V�[�g����
'  ConfigSheet([sname,bcreate]) �ݒ�V�[�g�擾
'  SectionTags(ra)              �Z�N�V�����^�O�Z���ꗗ���擾
'  SectionTagNames(ra)          �Z�N�V�����^�O���ꗗ���擾
'  SectionCell(ra,[tname])      �Z�N�V�����̈�擾
'
'  ActivateConfigSheet([s])     �ݒ�V�[�g�L��
'  SectionTags(ra)             �Z�N�V�����ꗗ���擾
'  LoadConfig([sec,sht,wb])     �ݒ�擾
'----------------------------------------

Option Explicit
Option Private Module

'----------------------------------------
'�ݒ�V�[�g����
'----------------------------------------

'�ݒ�V�[�g�擾
Function ConfigSheet( _
        Optional sname As String = "#config", _
        Optional bcreate As Boolean) As Worksheet
    Dim ws As Worksheet
    Set ws = TakeByName(ActiveWorkbook.Sheets, sname)
    If ws Is Nothing Then Set ws = TakeByName(ThisWorkbook.Sheets, sname)
    If ws Is Nothing Then
        If Not bcreate Then Exit Function
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = sname
    End If
    Set ConfigSheet = ws
End Function

'�Z�N�V�����Z���ꗗ���擾
Function SectionTags(ra As Range) As Range
    If ra Is Nothing Then Exit Function
    '
    Dim rb As Range, ce As Range, ce2 As Range
    Dim cnt As Long
    For Each ce In ra
        For Each ce2 In ce.Cells
            If Left(ce2.Value, 1) = "[" Then
                If cnt = 0 Then
                    Set rb = ce2
                Else
                    Set rb = Union(rb, ce2)
                End If
                cnt = cnt + 1
            End If
        Next ce2
    Next ce
    Set SectionTags = rb
End Function

'�Z�N�V�������z��ɕϊ�
Function SectionTagNames(ra As Range) As Variant
    If ra Is Nothing Then Exit Function
    '
    Dim ce As Range
    Dim i As Long
    Dim arr As Variant
    ReDim arr(1 To ra.Count)
    For Each ce In ra
        i = i + 1
        arr(i) = Replace(Replace(ce, "[", ""), "]", "")
    Next ce
    SectionTagNames = arr
End Function

'�Z�N�V�����^�O�擾
Function SectionCell(ra As Range, _
    Optional ByVal tname As String) As Range
    If ra Is Nothing Then Exit Function
    '
    Dim s As String
    s = tname
    If s = "" Then
        s = SelectArray(SectionTagNames(ra))
        If s = "" Then Exit Function
    End If
    '
    Dim rb As Range
    Set rb = TakeByValue(ra, "[" & s & "]")
    If rb Is Nothing Then Exit Function
    If rb.Count <> 1 Then Exit Function
    Set SectionCell = rb
End Function

'�Z�N�V�����̈�擾
Function SectionRange(ra As Range, Optional eol As Boolean) As Range
    If ra Is Nothing Then Exit Function
    '
    '�e���v���[�g�s�񐔎擾
    Dim t As String
    Dim i As Long
    Dim r As Long, rm As Long
    Dim c As Long, cm As Long
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    rm = ws.UsedRange.Rows.Count + ws.UsedRange.Row - ra.Row
    cm = ws.UsedRange.Columns.Count + ws.UsedRange.Column - ra.Column
    For r = 1 To rm
        i = ra(r, cm + 1).End(xlToLeft).Column - ra.Column + 1
        If c < i Then c = i
        t = LCase(Trim(ra.Offset(r, -2)))
        t = Left(t, 1)
        If t = "[" Then Exit For
        If eol And t = "#" Then Exit For
    Next r
    If r < rm Then rm = r
    cm = c
    If rm < 1 Or cm < 1 Then Exit Function
    Set SectionRange = ra.Resize(rm, cm)
End Function

'�Z�N�V�����̍ŏI�s���擾
Function SectionRowCount(ce As Range, Optional eol As Boolean) As Long
    If ce Is Nothing Then Exit Function
    '
    Dim s As String
    Dim i As Long, im As Long
    im = ce.Worksheet.UsedRange.Row + ce.Worksheet.UsedRange.Rows.Count - ce.Row + 1
    For i = 1 To im
        s = Left(ce.Offset(i, 0).Value, 1)
        If s = "[" Then Exit For
        If eol And s = "#" Then Exit For
    Next i
    SectionRowCount = i
End Function

'�Z�N�V�����J�n�s�ꗗ���擾
Function SectionRowsDict(ra As Range) As Dictionary
    Dim s As String
    Dim ce As Range
    Dim dict As Dictionary
    Set dict = New Dictionary
    For Each ce In ra
        s = ce.Value
        If Left(s, 1) = "[" Then
            s = Replace(Replace(s, "[", ""), "]", "")
            dict.Add s, ce.Row - ra.Row + 1
        End If
    Next ce
    Set SectionRowsDict = dict
End Function

'�ݒ�V�[�g�L��
Sub ActivateConfigSheet(Optional s As String = "#config")
    Dim ws As Worksheet
    Set ws = TakeByName(ActiveWorkbook.Sheets, s)
    If Not ws Is Nothing Then
        ws.Activate
        Exit Sub
    End If
    '
    Set ws = TakeByName(ThisWorkbook.Sheets, s)
    If Not ws Is Nothing Then
        ws.Copy After:=ActiveSheet
        Exit Sub
    End If
    '
    ActiveWorkbook.Sheets.Add
    ActiveSheet.name = s
End Sub

'�ݒ�擾
Private Function LoadConfig( _
        Optional sec As String, _
        Optional sht As String = "#config", _
        Optional wb As Workbook = Null) As Range
    '
    If wb Is Nothing Then Set wb = ThisWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sht)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = SelectSheet(wb)
    If ws Is Nothing Then Exit Function
    '
    Dim ra As Range
    If sec = "" Then
        Set ra = SectionTags(ws.UsedRange)
        Set ra = SelectCell(ra)
    Else
        For Each ra In SectionTags(ws.UsedRange)
            If ra.Value = sec Then Exit For
        Next ra
    End If
    If ra Is Nothing Then Exit Function
    Set ra = ra.Offset(1)
    Set LoadConfig = ra
End Function

