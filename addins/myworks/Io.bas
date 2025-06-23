Attribute VB_Name = "Io"
'==================================
'IO操作
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'機能呼び出し
'mode=1: 選択範囲を表形式でエクスポート
'     2: 選択範囲をリスト形式でエクスポート
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
'機能
'----------------------------------------

'選択範囲を表形式でエクスポート
Private Sub ExportRangeToSpreadSheet(ra As Range, utf8 As Boolean)
    Dim flt As String
    flt = "CSV ファイル,*.csv"
    flt = flt & ",Excel ブック,*.xlsx"
    flt = flt & ",Excel マクロ有効ブック,*.xlsm"
    flt = flt & ",テキストファイル,*.txt"
    flt = flt & ",XML データ,*.xml"
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

'選択範囲をリスト形式でエクスポート
' apd:   追加の場合は True
' frc:   強制上書きの場合は True
' enc:   文字コード指定 Shift_JIS, UTF-8, EUC-JP, ISO-2022-JP
' eol:   改行コード指定 -1:CRLF, 10:LF, 13:CR
      
Private Sub ExportRangeToText(ra As Range, _
        Optional ByVal apd As Boolean, _
        Optional ByVal frc As Boolean, _
        Optional ByVal enc As String = "Shift_JIS", _
        Optional ByVal eol As Integer = -1)
    If ra.Count < 2 Then
        MsgBox "複数のセルを選択してください。"
        Exit Sub
    End If
    If Not IsArray(ra.Value) Then Exit Sub
    '
    Dim flt As String, pth As String
    flt = "テキストファイル,*.txt"
    pth = GetIoSaveAsFilename(ActiveSheet.name, "txt", flt)
    If pth = "" Then Exit Sub
    '
    Dim msg As String
    Dim res As Variant
    If fso.FileExists(pth) Then
        If (Not apd) And (Not frc) Then
            msg = "既存ファイルがあります。上書きしますか？"
            res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton1)
            If res = vbCancel Then Exit Sub
            If res = vbYes Then
                frc = True
            Else
                msg = "既存ファイルに追加しますか？"
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
    msg = "文字コードはUTF-8ですか？"
    res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton2)
    If res = vbCancel Then Exit Sub
    enc = IIf(res = vbYes, "UTF-8", "Shift_JIS")
    '
    msg = "改行コードはLFですか？"
    res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton2)
    If res = vbCancel Then Exit Sub
    eol = IIf(res = vbYes, 10, -1)
    '
    Dim sep As String
    sep = " "
    msg = "分割文字はTABですか？"
    res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton1)
    If res = vbCancel Then Exit Sub
    If res = vbYes Then
        sep = Chr(9)
    Else
        msg = "分割文字は改行ですか？"
        res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton1)
        If res = vbCancel Then Exit Sub
        If res = vbYes Then sep = IIf(eol = 10, Chr(10), Chr(13) & Chr(10))
    End If
    '
    Dim blank As Boolean
    msg = "空行は無視しますか？"
    res = MsgBox(msg, vbYesNoCancel Or vbDefaultButton1)
    If res = vbCancel Then Exit Sub
    If res <> vbYes Then blank = True
    '
    WriteText ra, pth, apd, frc, enc, sep, blank, eol
End Sub

'----------------------------------
'共通機能
'----------------------------------

'保存ファイル名選択
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
        f = UCase(e) & " ファイル,*." & e & "," & f
    End If
    f = f + ",すべてのファイル,*.*"
    '
    p = Application.GetSaveAsFilename(p, f)
    If p = "False" Then Exit Function
    '
    GetIoSaveAsFilename = p
End Function

