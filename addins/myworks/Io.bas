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

Sub ExportProc(ra As Range, mode As Integer)
    Dim enc As Boolean
    enc = GetBookBool("export.1")
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
    Dim fmt As Integer
    Select Case LCase(fso.GetExtensionName(path))
    Case "txt": fmt = IIf(utf8, xlUnicodeText, xlText)
    Case "xml": fmt = xlXMLSpreadsheet
    Case "xlsx": fmt = xlOpenXMLWorkbook
    Case "xlsm": fmt = xlOpenXMLWorkbookMacroEnabled
    Case Else: fmt = IIf(utf8, xlCSVUTF8, xlCSV)
    End Select
    '
    ra.SpecialCells(xlCellTypeVisible).Copy
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
    wb.Worksheets(1).Paste
    wb.SaveAs Filename:=path, FileFormat:=fmt
    wb.Close
    '
    On Error GoTo 0
    ScreenUpdateOn
End Sub

'�I��͈͂����X�g�`���ŃG�N�X�|�[�g
' apnd:  �ǉ��̏ꍇ�� True
' frc:   �����㏑���̏ꍇ�� True
' enc:   �����R�[�h�w�� Shift_JIS, UTF-8, EUC-JP, ISO-2022-JP
' eol:   ���s�R�[�h�w�� -1:CRLF, 10:LF, 13:CR
      
Private Sub ExportRangeToText(ra As Range, _
        Optional ByVal apnd As Boolean, _
        Optional ByVal frc As Boolean, _
        Optional ByVal enc As String = "Shift_JIS", _
        Optional ByVal eol As Integer = -1)
    If Not IsArray(ra.Value) Then Exit Sub
    '
    Dim flt As String, pth As String
    flt = "�e�L�X�g�t�@�C��,*.txt"
    pth = GetIoSaveAsFilename(ActiveSheet.name, "txt", flt)
    If pth = "" Then Exit Sub
    '
    Dim res As Variant
    If fso.FileExists(pth) Then
        res = MsgBox("�����t�@�C��������܂��B�㏑�����܂����H", vbYesNoCancel Or vbDefaultButton2)
        If res = vbCancel Then Exit Sub
        If res = vbYes Then frc = True
        '
        res = MsgBox("�����t�@�C��������܂��B�����t�@�C���ɒǉ����܂����H", vbYesNoCancel Or vbDefaultButton2)
        If res = vbCancel Then Exit Sub
        If res = vbYes Then apnd = True
    End If
    '
    res = MsgBox("�����R�[�h��UTF-8�ł����H", vbYesNoCancel Or vbDefaultButton2)
    If res = vbCancel Then Exit Sub
    If res = vbYes Then enc = "UTF-8"
    '
    res = MsgBox("���s�R�[�h��LF�ł����H", vbYesNoCancel Or vbDefaultButton2)
    If res = vbCancel Then Exit Sub
    If res = vbYes Then eol = 10
    
    WriteText ra, pth, apnd, frc, enc, eol
End Sub

'----------------------------------
'���ʋ@�\
'----------------------------------

'�ۑ��t�@�C�����I��
Private Function GetIoSaveAsFilename( _
        Optional path As String, _
        Optional ext As String, _
        Optional flt As String)
    Dim p As String, e As String, n As String
    p = fso.GetFileName(path)
    e = ext
    If p = "" Then
        p = ActiveSheet.name
        n = fso.GetBaseName(ActiveWorkbook.name)
        If fso.GetExtensionName(p) <> "" Then
            e = fso.GetExtensionName(p)
            p = fso.GetBaseName(p)
        ElseIf LCase(n) <> LCase(p) Then
            p = n & "_" & p
        End If
    End If
    '
    If fso.GetExtensionName(p) = "" Then p = p & "." & e
    p = fso.BuildPath(ActiveWorkbook.path, p)
    e = LCase(e)
    '
    Dim f As String
    f = flt
    Dim v As Variant
    For Each v In Split(flt, ",", , vbTextCompare)
        If "*." & e = LCase(v) Then Exit For
    Next v
    If TypeName(v) <> "String" Then
        f = UCase(e) & " �t�@�C��,*." & e & "," & f
    End If
    f = f + ",���ׂẴt�@�C��,*.*"
    '
    p = Application.GetSaveAsFilename(p, f)
    If p = "False" Then Exit Function
    '
    GetIoSaveAsFilename = p
End Function

