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
    Dim enc As Boolean
    enc = GetBookBool("export.1")
    '
    Select Case mode
    Case 1: Call ExportRangeToSpreadSheet(ra, enc)
    Case 2: Call ExportRangeToText(ra, enc)
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

'選択範囲をリスト形式でエクスポート
' apnd:  追加の場合は True
' frc:   強制上書きの場合は True
' enc:   文字コード指定 Shift_JIS, UTF-8, EUC-JP, ISO-2022-JP
' eol:   改行コード指定 -1:CRLF, 10:LF, 13:CR
      
Private Sub ExportRangeToText(ra As Range, _
        Optional ByVal apnd As Boolean, _
        Optional ByVal frc As Boolean, _
        Optional ByVal enc As String = "Shift_JIS", _
        Optional ByVal eol As Integer = -1)
    If Not IsArray(ra.Value) Then Exit Sub
    '
    Dim flt As String, pth As String
    flt = "テキストファイル,*.txt"
    pth = GetIoSaveAsFilename(ActiveSheet.name, "txt", flt)
    If pth = "" Then Exit Sub
    '
    Dim res As Variant
    If fso.FileExists(pth) Then
        res = MsgBox("既存ファイルがあります。上書きしますか？", vbYesNoCancel Or vbDefaultButton2)
        If res = vbCancel Then Exit Sub
        If res = vbYes Then frc = True
        '
        res = MsgBox("既存ファイルがあります。既存ファイルに追加しますか？", vbYesNoCancel Or vbDefaultButton2)
        If res = vbCancel Then Exit Sub
        If res = vbYes Then apnd = True
    End If
    '
    res = MsgBox("文字コードはUTF-8ですか？", vbYesNoCancel Or vbDefaultButton2)
    If res = vbCancel Then Exit Sub
    If res = vbYes Then enc = "UTF-8"
    '
    res = MsgBox("改行コードはLFですか？", vbYesNoCancel Or vbDefaultButton2)
    If res = vbCancel Then Exit Sub
    If res = vbYes Then eol = 10
    
    WriteText ra, pth, apnd, frc, enc, eol
End Sub

'----------------------------------
'共通機能
'----------------------------------

'保存ファイル名選択
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
        f = UCase(e) & " ファイル,*." & e & "," & f
    End If
    f = f + ",すべてのファイル,*.*"
    '
    p = Application.GetSaveAsFilename(p, f)
    If p = "False" Then Exit Function
    '
    GetIoSaveAsFilename = p
End Function

