Attribute VB_Name = "CommonIO"
Option Explicit
Option Private Module

'==================================
'IO共通
'==================================

'----------------------------------------
'定数
'----------------------------------------

'拡張子マップ
Const c_exts As String = _
    "ワークシート;*.xlsx;*.xlsm;*.xls," & _
    "テンプレート;*.xltx;*.xlt," & _
    "アドイン;*.xlam;*.xla," & _
    "Webページ;*.htm;*.html," & _
    "XMLファイル;*.xml," & _
    "テキストファイル;*.txt;*.prn," & _
    "CSVファイル;*.csv," & _
    "リストファイル;*.lst," & _
    "ログファイル;*.log," & _
    "PDFファイル;*.pdf;*.fdf;*.xfdf," & _
    "すべてのファイル;*.*"


'----------------------------------------
'ファイルフィルタ
'----------------------------------------

' 拡張子種別名リストアップ
Public Sub ListupExtNames(Optional exts As String)
    If exts = "" Then exts = c_exts
    Dim ce As Range
    Set ce = Selection.Cells(1, 1)
    Dim dic As Object
    Set dic = ExtNamesMap(ce.Value, ExtNamesMap())
    If ce.Value <> "" Then Set ce = ce.Offset(1)
    Dim s As Variant
    For Each s In dic.Keys
        ce.Value = s
        ce.Offset(, 1).Value = dic(s)
        Set ce = ce.Offset(1)
    Next s
End Sub

' 拡張子種別名マップ取得
Private Function ExtNamesMap( _
        Optional exts As String, _
        Optional nest As Object) As Object
    If exts = "" Then exts = c_exts
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Dim s As Variant
    For Each s In Split(exts, ",")
        Dim kv() As String
        If s <> "" Then
            kv = Split(s, ";", 2)
            Dim k As String
            k = Trim(kv(0))
            If UBound(kv) > 0 Then
                dic.Add k, Replace(kv(1), " ", "")
            ElseIf nest Is Nothing Then
            Else
                Dim v As Variant
                v = nest(k)
                If v = "" Then
                    k = Replace(k, ".", "")
                    k = Replace(k, "ファイル", "")
                    v = "*." & k
                    k = k & "ファイル"
                    dic.Add k, v
                Else
                    dic.Add k, nest(k)
                End If
            End If
        End If
    Next s
    Set ExtNamesMap = dic
End Function

'----------------------------------------
'ファイル選択
'----------------------------------------

'単体ファイル選択
Public Function SelectFile( _
        Optional path As String, _
        Optional Title As String, _
        Optional exts As String) As String
    Dim res As Variant
    Set res = SelectFiles(path, Title, exts, False)
    If res Is Nothing Then Exit Function
    If res.Count > 0 Then
        SelectFile = res(1)
    End If
End Function

'複数ファイル選択
Public Function SelectFiles( _
        Optional path As String, _
        Optional Title As String, _
        Optional exts As String, _
        Optional multi As Boolean = True) As Variant
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = multi
        If Title <> "" Then .Title = Title
        If exts = "" Then exts = c_exts
        .Filters.Clear
        Dim dic As Object
        Set dic = ExtNamesMap(exts, ExtNamesMap(c_exts))
        Dim s As Variant
        For Each s In dic.Keys
            .Filters.Add s, dic(s)
        Next s
        .FilterIndex = 0
        If path <> "" Then
            path = re_replace(path, "^\((\w+)\)", "%$1%")
            If Not fso.FileExists(path) Then path = path & "\"
            .InitialFileName = path
        End If
        .Show
        Set SelectFiles = .SelectedItems
    End With
End Function

'----------------------------------------
'ファイル選択(個別)
'----------------------------------------

'CSVファイル選択
Public Function SelectCsvFile( _
        Optional path As String, _
        Optional Title As String) As String
    Dim filter As String
    filter = "CSVファイル,すべてのファイル"
    SelectCsvFile = SelectFile(path, Title, filter)
End Function

'リストファイル選択
Public Function SelectListFile( _
        Optional path As String, _
        Optional Title As String) As String
    Dim filter As String
    filter = "リストファイル,CSVファイル,すべてのファイル"
    SelectListFile = SelectFile(path, Title, filter)
End Function

'----------------------------------------
'フォルダ選択
'----------------------------------------

Public Function SelectFolder( _
        Optional path As String = "", _
        Optional mode As Integer = 1) As String
    If path = "" Then
        path = ActiveWorkbook.path
    Else
        path = re_replace(path, "^\((\w+)\)", "%$1%")
    End If
    '
    If mode = 1 Then
        SelectFolder = SelectFolder1(path)
    Else
        SelectFolder = SelectFolder2(path)
    End If
End Function

Private Function SelectFolder1(path As String) As String
    If Not Right(path, 1) = "\" Then path = path + "\"
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = path
        .AllowMultiSelect = False
        .Title = "フォルダの選択"
        If Not .Show Then Exit Function
        SelectFolder1 = .SelectedItems(1)
    End With
End Function

Private Function SelectFolder2(path As String) As String
    With CreateObject("Shell.Application")
        Dim fo As Object
        Set fo = .BrowseForFolder(Application.Hwnd, _
                        "フォルダを選んでください", _
                        &H1 + &H10 + &H200, path)
        If fo Is Nothing Then Exit Function
        SelectFolder2 = fo.Items.Item.path
    End With
End Function


'----------------------------------------
'選択ダイアログ
'----------------------------------------

'ワークブック選択
Function SelectBook(Optional ptn As String) As Workbook
    SelectFormX.Reset "ブック", ptn
    SelectFormX.AddNames Workbooks
    SelectFormX.Show
    Dim s As String
    s = SelectFormX.Result
    Unload SelectFormX
    If s <> "" Then Set SelectBook = Workbooks(s)
End Function

'ワークシート選択
Function SelectSheet(Optional wb As Workbook, Optional ptn As String) As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    SelectFormX.Reset "シート", ptn
    SelectFormX.AddNames wb.Worksheets
    If SelectFormX.ItemCount > 0 Then SelectFormX.Show
    Dim s As String
    s = SelectFormX.Result
    Unload SelectFormX
    If s <> "" Then Set SelectSheet = wb.Worksheets(s)
End Function

'セル選択
Function SelectCell(Optional ra As Range, Optional s As String, Optional ptn As String) As Range
    If ra Is Nothing Then Set ra = Selection
    With SelectFormX
        .Reset s, ptn
        .AddValues ra
        .Show
        Dim i As Integer
        i = .index
        Dim v As Variant
        For Each v In ra
            If i = 0 Then Exit For
            i = i - 1
        Next v
        If Not v = Empty Then Set SelectCell = v
    End With
    Unload SelectFormX
End Function

'アドイン選択
Function SelectAddin(Optional ptn As String) As addin
    With SelectFormX
        .Reset "アドイン", ptn
        Dim v As Variant
        For Each v In AddIns
            .AddItem v.Title
        Next v
        '.AddNames AddIns
        .Show
        Dim s As String
        s = .Result
    End With
    Unload SelectFormX
    If s <> "" Then Set SelectAddin = AddIns(s)
End Function

Public Function SelectAddinsSheet() As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    Set ws = ActiveSheet
    Set wb = ActiveWorkbook
    ThisWorkbook.Activate
    With CommandBars.Add(temporary:=True)
        .Controls.Add(id:=957).Execute
        .Delete
    End With
    If Not ThisWorkbook.ActiveSheet Is ws Then
        Set SelectAddinsSheet = ThisWorkbook.ActiveSheet
    End If
    ThisWorkbook.IsAddin = False
    wb.Activate
    
    ws.Select
    Application.ScreenUpdating = True
End Function

'名前選択・移動
Public Function SelectName() As Range
    Application.Dialogs(63).Show
End Function


'----------------------------------------
'読み込み・保存
'----------------------------------------

'ワークシート取込
Public Sub OpenWorkbook()
    Dim path As Variant
    path = Application.GetOpenFilename( _
        FileFilter:="Excelファイル,*.xls*,Csvファイル,*.csv" _
        & ",テキストファイル,*.txt,全てのファイル,*.*" _
    )
    If path = False Then
        Exit Sub
    End If
    '
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    '
    Dim old As Worksheet
    Set old = ActiveSheet
    '
    Dim wb As Workbook
    Set wb = Workbooks.Open(path, ReadOnly:=True)
    '
    Dim ws As Worksheet
    Set ws = SelectSheet(wb)
    ws.name = UniqueSheetName(wb, ws.name)
    ws.Copy After:=old
    wb.Close
    Application.EnableEvents = True
    Application.ScreenUpdating = fsu
End Sub

'ワークブックセーブ選択
Public Function SaveWorkbook(Optional path As Variant = False) As String
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    path = Application.GetSaveAsFilename( _
        FileFilter:="Excelファイル,*.xlsx" _
    )
    If path <> False Then
        Application.EnableEvents = False
        ActiveWorkbook.SaveAs path
        Application.EnableEvents = True
        SaveWorkbook = path
    End If
    '
    Application.ScreenUpdating = fsu
End Function

