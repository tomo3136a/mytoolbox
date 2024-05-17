Attribute VB_Name = "CommonIO"
Option Explicit
Option Private Module

'==================================
'IO����
'==================================

'----------------------------------------
'�萔
'----------------------------------------

'�g���q�}�b�v
Const c_exts As String = _
    "���[�N�V�[�g;*.xlsx;*.xlsm;*.xls," & _
    "�e���v���[�g;*.xltx;*.xlt," & _
    "�A�h�C��;*.xlam;*.xla," & _
    "Web�y�[�W;*.htm;*.html," & _
    "XML�t�@�C��;*.xml," & _
    "�e�L�X�g�t�@�C��;*.txt;*.prn," & _
    "CSV�t�@�C��;*.csv," & _
    "���X�g�t�@�C��;*.lst," & _
    "���O�t�@�C��;*.log," & _
    "PDF�t�@�C��;*.pdf;*.fdf;*.xfdf," & _
    "���ׂẴt�@�C��;*.*"


'----------------------------------------
'�t�@�C���t�B���^
'----------------------------------------

' �g���q��ʖ����X�g�A�b�v
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

' �g���q��ʖ��}�b�v�擾
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
                    k = Replace(k, "�t�@�C��", "")
                    v = "*." & k
                    k = k & "�t�@�C��"
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
'�t�@�C���I��
'----------------------------------------

'�P�̃t�@�C���I��
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

'�����t�@�C���I��
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
'�t�@�C���I��(��)
'----------------------------------------

'CSV�t�@�C���I��
Public Function SelectCsvFile( _
        Optional path As String, _
        Optional Title As String) As String
    Dim filter As String
    filter = "CSV�t�@�C��,���ׂẴt�@�C��"
    SelectCsvFile = SelectFile(path, Title, filter)
End Function

'���X�g�t�@�C���I��
Public Function SelectListFile( _
        Optional path As String, _
        Optional Title As String) As String
    Dim filter As String
    filter = "���X�g�t�@�C��,CSV�t�@�C��,���ׂẴt�@�C��"
    SelectListFile = SelectFile(path, Title, filter)
End Function

'----------------------------------------
'�t�H���_�I��
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
        .Title = "�t�H���_�̑I��"
        If Not .Show Then Exit Function
        SelectFolder1 = .SelectedItems(1)
    End With
End Function

Private Function SelectFolder2(path As String) As String
    With CreateObject("Shell.Application")
        Dim fo As Object
        Set fo = .BrowseForFolder(Application.Hwnd, _
                        "�t�H���_��I��ł�������", _
                        &H1 + &H10 + &H200, path)
        If fo Is Nothing Then Exit Function
        SelectFolder2 = fo.Items.Item.path
    End With
End Function


'----------------------------------------
'�I���_�C�A���O
'----------------------------------------

'���[�N�u�b�N�I��
Function SelectBook(Optional ptn As String) As Workbook
    SelectFormX.Reset "�u�b�N", ptn
    SelectFormX.AddNames Workbooks
    SelectFormX.Show
    Dim s As String
    s = SelectFormX.Result
    Unload SelectFormX
    If s <> "" Then Set SelectBook = Workbooks(s)
End Function

'���[�N�V�[�g�I��
Function SelectSheet(Optional wb As Workbook, Optional ptn As String) As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    SelectFormX.Reset "�V�[�g", ptn
    SelectFormX.AddNames wb.Worksheets
    If SelectFormX.ItemCount > 0 Then SelectFormX.Show
    Dim s As String
    s = SelectFormX.Result
    Unload SelectFormX
    If s <> "" Then Set SelectSheet = wb.Worksheets(s)
End Function

'�Z���I��
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

'�A�h�C���I��
Function SelectAddin(Optional ptn As String) As addin
    With SelectFormX
        .Reset "�A�h�C��", ptn
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

'���O�I���E�ړ�
Public Function SelectName() As Range
    Application.Dialogs(63).Show
End Function


'----------------------------------------
'�ǂݍ��݁E�ۑ�
'----------------------------------------

'���[�N�V�[�g�捞
Public Sub OpenWorkbook()
    Dim path As Variant
    path = Application.GetOpenFilename( _
        FileFilter:="Excel�t�@�C��,*.xls*,Csv�t�@�C��,*.csv" _
        & ",�e�L�X�g�t�@�C��,*.txt,�S�Ẵt�@�C��,*.*" _
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

'���[�N�u�b�N�Z�[�u�I��
Public Function SaveWorkbook(Optional path As Variant = False) As String
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    path = Application.GetSaveAsFilename( _
        FileFilter:="Excel�t�@�C��,*.xlsx" _
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

