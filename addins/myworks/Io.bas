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
    Dim enc As Boolean, apd As Boolean
    enc = GetBookBool("export.1")
    apd = GetBookBool("export.2")
    Dim enc_name As String
    enc_name = IIf(enc, "UTF-8", "Shift_JIS")
    '
    Select Case mode
    Case 1: Call ExportRangeToSpreadSheet(ra, enc)
    Case 2: Call ExportRangeToText(ra, apd, False, enc_name)
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
    Dim pth As String
    pth = GetIoSaveAsFilename("", "csv", flt)
    If pth = "" Then Exit Sub
    '
    Dim fmt As Integer
    Select Case LCase(fso.GetExtensionName(pth))
    Case "txt": fmt = IIf(utf8, xlUnicodeText, xlText)
    Case "xml": fmt = xlXMLSpreadsheet
    Case "xlsx": fmt = xlOpenXMLWorkbook
    Case "xlsm": fmt = xlOpenXMLWorkbookMacroEnabled
    Case Else: fmt = IIf(utf8, xlCSVUTF8, xlCSV)
    End Select
    '
    Dim ra2 As Range
    Set ra2 = ra
    If ra2.Count = 1 Then Set ra2 = ra.Worksheet.UsedRange
    '
    ScreenUpdateOff
    On Error Resume Next
    Application.Visible = False
    '
    Dim n As Integer
    n = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    Application.SheetsInNewWorkbook = n
    '
    ra2.Copy Destination:=wb.Worksheets(1).Cells(1, 1)
    '
    wb.SaveAs Filename:=pth, FileFormat:=fmt
    wb.Close
    '
    Application.Visible = True
    On Error GoTo 0
    ScreenUpdateOn
End Sub

'�I��͈͂����X�g�`���ŃG�N�X�|�[�g
' apd:   �ǉ��̏ꍇ�� True
' frc:   �����㏑���̏ꍇ�� True
' enc:   �����R�[�h�w�� Shift_JIS, UTF-8, EUC-JP, ISO-2022-JP
' eol:   ���s�R�[�h�w�� -1:CRLF, 10:LF, 13:CR
      
Private Sub ExportRangeToText(ra As Range, _
        Optional ByVal apd As Boolean, _
        Optional ByVal frc As Boolean, _
        Optional ByVal enc As String = "Shift_JIS", _
        Optional ByVal eol As Integer = -1)
    If ra.Count < 2 Then
        MsgBox "�����̃Z����I�����Ă��������B"
        Exit Sub
    End If
    If Not IsArray(ra.Value) Then Exit Sub
    '
    Dim flt As String, pth As String
    flt = "�e�L�X�g�t�@�C��,*.txt"
    pth = GetIoSaveAsFilename(ActiveSheet.name, "txt", flt)
    If pth = "" Then Exit Sub
    '
    Dim msg As String
    Dim res As Variant
    If fso.FileExists(pth) Then
        If (Not apd) And (Not frc) Then
            msg = "�����t�@�C��������܂��B�㏑�����܂����H"
            res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton1)
            If res = vbCancel Then Exit Sub
            If res = vbYes Then
                frc = True
            Else
                msg = "�����t�@�C���ɒǉ����܂����H"
                res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton1)
                If res = vbCancel Then Exit Sub
                If res = vbYes Then apd = True
            End If
        End If
    Else
        apd = False
        frc = False
    End If
    '
    msg = "�����R�[�h��UTF-8�ł����H"
    res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton2)
    If res = vbCancel Then Exit Sub
    enc = IIf(res = vbYes, "UTF-8", "Shift_JIS")
    '
    msg = "���s�R�[�h��LF�ł����H"
    res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton2)
    If res = vbCancel Then Exit Sub
    eol = IIf(res = vbYes, 10, -1)
    '
    Dim sep As String
    sep = " "
    msg = "����������TAB�ł����H"
    res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton1)
    If res = vbCancel Then Exit Sub
    If res = vbYes Then
        sep = Chr(9)
    Else
        msg = "���������͉��s�ł����H"
        res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton1)
        If res = vbCancel Then Exit Sub
        If res = vbYes Then sep = IIf(eol = 10, Chr(10), Chr(13) & Chr(10))
    End If
    '
    Dim blank As Boolean
    msg = "��s�͖������܂����H"
    res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton1)
    If res = vbCancel Then Exit Sub
    If res <> vbYes Then blank = True
    '
    WriteText ra, pth, apd, frc, enc, sep, blank, eol
End Sub

'----------------------------------
'���ʋ@�\
'----------------------------------

'�ۑ��t�@�C�����I��
Private Function GetIoSaveAsFilename( _
        Optional pth As String, _
        Optional ext As String, _
        Optional flt As String)
    Dim p As String, e As String, n As String
    p = fso.GetFileName(pth)
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

