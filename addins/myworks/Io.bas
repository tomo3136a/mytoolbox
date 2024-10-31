Attribute VB_Name = "Io"
'==================================
'IO操作
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'機能呼び出し
'----------------------------------------

Sub MenuExport(ra As Range, mode As Integer)
    Select Case mode
    Case 1
        '選択範囲をスープレットシートにエクスポート
        Call ExportRangeToSpreadSheet(ra, GetRtParamBool("export", 1))
    Case 2
        '選択範囲をテキストファイルにエクスポート
        Call ExportRangeToText(ra, GetRtParamBool("export", 1))
    End Select
End Sub

'----------------------------------------
'機能
'----------------------------------------

'選択範囲をスープレットシートにエクスポート
Private Sub ExportRangeToSpreadSheet(ra As Range, utf8 As Boolean)
    Dim flt As String
    flt = "CSV ファイル,*.csv"
    flt = flt & ",Excel ブック,*.xlsx"
    flt = flt & ",EXxcel マクロ有効ブック,*.xlsm"
    flt = flt & ",テキストファイル,*.txt"
    flt = flt & ",XML データ,*.xml"
    '
    Dim path As String
    path = GetIoSaveAsFilename("", "csv", flt)
    If path = "" Then Exit Sub
    '
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    '
    Dim n As Integer
    n = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    Application.SheetsInNewWorkbook = n
    '
    Set ra = ra.SpecialCells(xlCellTypeVisible)
    ra.Copy
    wb.Worksheets(1).Paste
    Select Case LCase(fso.GetExtensionName(path))
    Case "txt"
        n = xlText
        If utf8 Then n = xlUnicodeText
    Case "xml"
        n = xlXMLSpreadsheet
    Case "xlsx"
        n = xlOpenXMLWorkbook
    Case "xlsm"
        n = xlOpenXMLWorkbookMacroEnabled
    Case Else
        n = xlCSV
        If utf8 Then n = xlCSVUTF8
    End Select
    wb.SaveAs Filename:=path, FileFormat:=n
    wb.Close
    '
    On Error GoTo 0
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'選択範囲をテキストファイルにエクスポート
Private Sub ExportRangeToText(ra As Range, utf8 As Boolean)
    If Not IsArray(ra.Value) Then Exit Sub
    Dim path As String
    path = GetIoSaveAsFilename(ActiveSheet.name, "txt", "テキストファイル,*.txt")
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
'共通機能
'----------------------------------

'保存ファイル名選択
Private Function GetIoSaveAsFilename(Optional path As String, Optional ext As String, Optional flt As String)
    If path = "" Then path = fso.GetBaseName(ActiveWorkbook.name) & "_" & ActiveSheet.name
    If fso.GetExtensionName(path) = "" Then path = path & "." & ext
    path = fso.BuildPath(ActiveWorkbook.path, path)
    If fso.GetExtensionName(path) <> ext Then
        ext = LCase(fso.GetExtensionName(path))
        flt = UCase(ext) & "ファイル,*." & ext
    End If
    If flt = "" Then flt = UCase(ext) & "ファイル,*." & ext
    flt = flt + ",すべてのファイル,*.*"
    path = Application.GetSaveAsFilename(path, flt)
    If path = "False" Then Exit Function
    '
    GetIoSaveAsFilename = path
End Function

