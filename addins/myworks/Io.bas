Attribute VB_Name = "Io"
'==================================
'IO����
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'�@�\�Ăяo��
'mode=1: �I��͈͂�\�`���ŃG�N�X�|�[�g
'     2: �I��͈͂����X�g�`���ŃG�N�X�|�[�g
'----------------------------------------

Sub MenuExport(ra As Range, mode As Integer)
    Dim enc As Boolean
    enc = GetRtParamBool("export", 1)
    '
    Select Case mode
    Case 1: Call ExportRangeToSpreadSheet(ra, enc)
    Case 2: Call ExportRangeToText(ra, enc)
    End Select
End Sub

'----------------------------------------
'�@�\
'----------------------------------------

'�I��͈͂�\�`���ŃG�N�X�|�[�g
Private Sub ExportRangeToSpreadSheet(ra As Range, utf8 As Boolean)
    Dim flt As String
    flt = "CSV �t�@�C��,*.csv"
    flt = flt & ",Excel �u�b�N,*.xlsx"
    flt = flt & ",Excel �}�N���L���u�b�N,*.xlsm"
    flt = flt & ",�e�L�X�g�t�@�C��,*.txt"
    flt = flt & ",XML �f�[�^,*.xml"
    '
    Dim path As String
    path = GetIoSaveAsFilename("", "csv", flt)
    If path = "" Then Exit Sub
    '
    ScreenUpdateOff
    On Error Resume Next
    '
    Dim n As Integer
    n = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    Application.SheetsInNewWorkbook = n
    '
    Dim rb As Range
    Set rb = ra.SpecialCells(xlCellTypeVisible)
    rb.Copy
    wb.Worksheets(1).Paste
    '
    Select Case LCase(fso.GetExtensionName(path))
    Case "txt": n = xlText: If utf8 Then n = xlUnicodeText
    Case "xml": n = xlXMLSpreadsheet
    Case "xlsx": n = xlOpenXMLWorkbook
    Case "xlsm": n = xlOpenXMLWorkbookMacroEnabled
    Case Else: n = xlCSV: If utf8 Then n = xlCSVUTF8
    End Select
    '
    wb.SaveAs Filename:=path, FileFormat:=n
    wb.Close
    '
    On Error GoTo 0
    ScreenUpdateOn
End Sub

'�I��͈͂����X�g�`���ŃG�N�X�|�[�g
Private Sub ExportRangeToText(ra As Range, utf8 As Boolean)
    If Not IsArray(ra.Value) Then Exit Sub
    '
    Dim flt As String
    flt = "�e�L�X�g�t�@�C��,*.txt"
    '
    Dim path As String
    path = GetIoSaveAsFilename(ActiveSheet.name, "txt", flt)
    If path = "" Then Exit Sub
    '
    Open path For Output As #1
    Dim rs As Variant
    For Each rs In wsf.Transpose(ra.Value)
        Dim line As String
        line = Trim(rs)
        If line <> "" Then Print #1, line
    Next rs
    Close #1
End Sub

'----------------------------------
'���ʋ@�\
'----------------------------------

'�ۑ��t�@�C�����I��
Private Function GetIoSaveAsFilename( _
        Optional path As String, _
        Optional ext As String, _
        Optional flt As String)
    If path = "" Then path = fso.GetBaseName(ActiveWorkbook.name) & "_" & ActiveSheet.name
    If fso.GetExtensionName(path) = "" Then path = path & "." & ext
    path = fso.BuildPath(ActiveWorkbook.path, path)
    If fso.GetExtensionName(path) <> ext Then
        ext = LCase(fso.GetExtensionName(path))
        flt = UCase(ext) & "�t�@�C��,*." & ext
    End If
    If flt = "" Then flt = UCase(ext) & "�t�@�C��,*." & ext
    flt = flt + ",���ׂẴt�@�C��,*.*"
    path = Application.GetSaveAsFilename(path, flt)
    If path = "False" Then Exit Function
    '
    GetIoSaveAsFilename = path
End Function

