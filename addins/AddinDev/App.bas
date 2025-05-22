Attribute VB_Name = "App"
'==================================
'�A�h�C���J��
'==================================

'[�Q�Ɛݒ�]
'�uMicrosoft Visual Basic for Application Extensibility 5.3�v

Option Explicit
Option Private Module

Private g_addin As String
Private g_image As String

'----------------------------------
'�A�v���P�[�V����I/F
'----------------------------------

Sub AddinDevFolderProc(id As Long)
    Select Case id
        Case 1: OpenCurrentFolder       '�J�����g�t�H���_���J��
        Case 2: OpenAddinsFolder        '�A�h�C���t�H���_���J��
        Case 3: OpenImageMsoFolder      'ImageMso�t�@�C���t�H���_���J��
    End Select
End Sub

Sub AddinDevEditProc(id As Long)
    Select Case id
    Case 1: EditCustomUI g_addin        'CustomUI �ҏW
    Case 2: MargeCustomUI g_addin       'CustomUI �}�[�W
    Case 3: DeployAddin g_addin         '�A�h�C���z�u
    Case 4: ToggleAddin g_addin         '�A�h�C���u�b�N�\���E��\���g�O��
    Case 5: OpenAddinManager            '�A�h�C���}�l�[�W��
    Case 6: ExportModules g_addin       '�A�h�C���\�[�X�̃G�N�X�|�[�g
    Case 7: ImportModules g_addin       '�A�h�C���\�[�X�̃C���|�[�g
    Case 8: ReloadAddin g_addin         '�A�h�C���ēǂݍ���
    Case 9: ToggleAddin ActiveWorkbook.name     '����
    End Select
End Sub

Sub AddinDevCallProc(id As Long)
    Select Case id
    Case 1: CallDialog                  '�_�C�A���O�Ăяo��
    Case 2: If ActiveCell.Value <> "" Then g_image = ActiveCell.Value
    End Select
End Sub

'----------------------------------
'�ݒ�
'----------------------------------

Sub SetAddinName(s As String)
    g_addin = s
    g_image = "About"
End Sub

Function GetButtonImage() As String
    GetButtonImage = g_image
End Function

'----------------------------------
'�I�[�v���t�H���_
'----------------------------------

'�J�����g�t�H���_���J��
Private Sub OpenCurrentFolder()
    With CreateObject("Wscript.Shell")
        .Run ActiveWorkbook.path
    End With
End Sub

'�A�h�C���t�H���_���J��
Private Sub OpenAddinsFolder()
    With CreateObject("Wscript.Shell")
        .Run AddinsPath
    End With
End Sub

'ImageMso�t�@�C���t�H���_���J���A������΍쐬
Private Sub OpenImageMsoFolder()
    Dim name As String
    name = "ImageMso"
    '
    Dim parent As String
    parent = fso.BuildPath(Environ("USERPROFILE"), "Documents")
    parent = fso.BuildPath(parent, name)
    If Not fso.FolderExists(parent) Then
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(name)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox name & "�V�[�g���K�v�ł��B"
            Exit Sub
        End If
        '
        Dim ra As Range
        Set ra = ws.UsedRange
        If ra Is Nothing Then
            MsgBox "�f�[�^��������܂���ł����B"
            Exit Sub
        End If
        If ra.Cells(1, 1).Value = "" Then
            MsgBox "�f�[�^��������܂���ł����B"
            Exit Sub
        End If
        Dim cnt As Long
        cnt = ra.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
        Dim arr As Variant
        arr = ws.Range("A2:A" & cnt).Value
        '
        fso.CreateFolder parent
        If Not fso.FolderExists(parent) Then
            MsgBox parent & " �t�H���_���쐬�ł��܂���ł����B"
            Exit Sub
        End If
        '
        Dim cb As CommandBars
        Set cb = Application.CommandBars
        Dim i As Long
        On Error Resume Next
        ProgressStatusBar 0
        For i = LBound(arr, 1) To UBound(arr, 1)
            Dim path As String
            path = fso.BuildPath(parent, arr(i, 1) & ".png")
            Dim img As IPictureDisp
            Set img = cb.GetImageMso(arr(i, 1), 128, 128)
            Call stdole.SavePicture(img, path)
            '
            If i Mod 100 = 0 Then ProgressStatusBar i, cnt
        Next i
        ProgressStatusBar
        On Error GoTo 0
    End If
    '
    With CreateObject("Wscript.Shell")
        .Run parent
    End With
End Sub

'----------------------------------
'�A�h�C���ҏW
'----------------------------------

'CustomUI �ҏW
Private Sub EditCustomUI(name As String)
    Dim xlam As String
    xlam = fso.BuildPath(AddinsPath, name)
    If Not fso.FileExists(xlam) Then Exit Sub
    '
    Dim tmp As String
    tmp = fso.BuildPath(AddinsPath, "tmp")
    If Not fso.FolderExists(tmp) Then fso.CreateFolder tmp
    '
    Dim base As String
    base = fso.GetBaseName(name)
    '
    Dim zip As String
    zip = fso.BuildPath(tmp, base & ".zip")
    fso.CopyFile xlam, zip
    '
    Dim src As String
    src = fso.BuildPath(zip, "CustomUI")
    '
    Dim dst As String
    dst = fso.BuildPath(tmp, base)
    If Not fso.FolderExists(dst) Then fso.CreateFolder dst
    dst = fso.BuildPath(dst, "CustomUI")
    If Not fso.FolderExists(dst) Then fso.CreateFolder dst
    '
    Dim path As String
    path = fso.BuildPath(dst, "customUI.xml")
    If Not fso.FileExists(path) Then
        Dim shell As Object
        Set shell = CreateObject("Shell.Application")
        Dim fo As Object
        Set fo = shell.Namespace(CVar(dst))
        fo.CopyHere shell.Namespace(CVar(src)).Items
        If Not shell Is Nothing Then Set shell = Nothing
    End If
    '
    If fso.FileExists(zip) Then fso.DeleteFile zip
    '
    With CreateObject("Wscript.Shell")
        .Run path
    End With
End Sub

'CustomUI �}�[�W
Private Sub MargeCustomUI(name As String)
    On Error Resume Next
    Workbooks(g_addin).Save
    On Error GoTo 0
    '
    Dim xlam As String
    xlam = fso.BuildPath(AddinsPath, name)
    If Not fso.FileExists(xlam) Then Exit Sub
    '
    Dim tmp As String
    tmp = fso.BuildPath(AddinsPath, "tmp")
    If Not fso.FolderExists(tmp) Then fso.CreateFolder tmp
    '
    Dim base As String
    base = fso.GetBaseName(name)
    '
    Dim zip As String
    zip = fso.BuildPath(tmp, base & ".zip")
    fso.CopyFile xlam, zip
    '
    Dim src As String
    src = fso.BuildPath(tmp, base)
    src = fso.BuildPath(src, "CustomUI")
    '
    Dim dst As String
    dst = fso.BuildPath(zip, "CustomUI")
    '
    Dim path As String
    path = fso.BuildPath(src, "customUI.xml")
    If Not fso.FileExists(path) Then
        MsgBox "CustomUI.xml ������܂���B"
        Exit Sub
    End If
    '
    Dim shell As Object
    Set shell = CreateObject("Shell.Application")
    Dim fo As Object
    Set fo = shell.Namespace(CVar(dst))
    fo.CopyHere shell.Namespace(CVar(src)).Items
    '
    If Not shell Is Nothing Then Set shell = Nothing
End Sub

'�A�h�C���z�u
Private Sub DeployAddin(name As String)
    If name = "" Then Exit Sub
    If name Like ThisWorkbook.name Then
        MsgBox name & " �A�h�C���͔z�u�ł��܂���B"
        Application.EnableEvents = False
        Exit Sub
    End If
    '
    Dim base As String
    base = fso.GetBaseName(name)
    '
    Dim src As String
    src = fso.BuildPath(ThisWorkbook.path, "tmp")
    src = fso.BuildPath(src, base & ".zip")
    If Not fso.FileExists(src) Then
        MsgBox base & ".zip �t�@�C��������܂���B"
        Exit Sub
    End If
    '
    Dim dst As String
    dst = fso.BuildPath(ThisWorkbook.path, name)
    '
    Dim ai As addin
    For Each ai In AddIns
        If ai.name Like name Then Exit For
    Next ai
    If ai Is Nothing Then
        MsgBox name & " �A�h�C���̓o�^������܂���B"
        Exit Sub
    End If
    '
    Dim kw As String
    kw = ai.Title
    AddIns(kw).Installed = False
    If fso.FileExists(dst) Then fso.DeleteFile dst
    fso.MoveFile src, dst
    AddIns(kw).Installed = True
End Sub

'�A�h�C���}�l�[�W���I�[�v��
Private Sub OpenAddinManager()
    Application.Dialogs(xlDialogAddinManager).Show
End Sub

'���W���[���\�[�X�R�[�h�G�N�X�|�[�g
Private Sub ExportModules(name As String)
    If name = "" Then Exit Sub
    '
    Dim ai As addin
    For Each ai In AddIns
        If ai.name Like name Then Exit For
    Next ai
    If ai Is Nothing Then
        MsgBox name & " �A�h�C���̓o�^������܂���B"
        Exit Sub
    End If
    '
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Application.Workbooks(ai.name)
    On Error GoTo 0
    If wb Is Nothing Then
        MsgBox ai.name & "��L���ɂ��Ă��������B"
        Exit Sub
    End If
    '
    On Error Resume Next
    Dim col As Object
    Set col = wb.VBProject.VBComponents
    On Error GoTo 0
    If col Is Nothing Then
        MsgBox "�uVBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ�" & _
                "�A�N�Z�X��M������v�ɐݒ肵�Ă��������B"
        Exit Sub
    End If
    '
    Dim path As String
    path = ActiveWorkbook.path
    If Not Right(path, 1) = "\" Then path = path + "\"
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = path
        .AllowMultiSelect = False
        .Title = "�A�h�C���\�[�X���[�g�t�H���_�̑I��"
        If Not .Show Then Exit Sub
        path = .SelectedItems(1)
    End With
    path = fso.BuildPath(path, fso.GetBaseName(name))
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
    '
    Dim m As Object
    For Each m In col
        If m.CodeModule.CountOfLines > 0 Then
            '1:vbext_ct_StdModule
            '2:vbext_ct_ClassModule
            '3:vbext_ct_MSForm
            '100:vbext_ct_Document
            Select Case m.Type
            Case 1: Call m.Export(path & "\" & m.name & ".bas")
            Case 2: Call m.Export(path & "\" & m.name & ".cls")
            Case 3: Call m.Export(path & "\" & m.name & ".frm")
            Case 100: Call m.Export(path & "\" & m.name & ".cls")
            Case Else: MsgBox "�G�N�X�|�[�g�ΏۊO�F " & m.Type & ":" & m.name
            End Select
        End If
    Next m
    '
    Dim src As String
    src = fso.BuildPath(AddinsPath, "tmp")
    src = fso.BuildPath(src, fso.GetBaseName(name))
    src = fso.BuildPath(src, "CustomUI")
    If fso.FileExists(fso.BuildPath(src, "customUI.xml")) Then
        Dim dst As String
        dst = fso.BuildPath(path, "CustomUI")
        If Not fso.FolderExists(dst) Then fso.CreateFolder dst
        '
        Dim shell As Object
        Set shell = CreateObject("Shell.Application")
        Dim fo As Object
        Set fo = shell.Namespace(CVar(dst))
        fo.CopyHere shell.Namespace(CVar(src)).Items
        If Not shell Is Nothing Then Set shell = Nothing
    End If
    '
    With CreateObject("Wscript.Shell")
        .Run path
    End With
End Sub

'���W���[���\�[�X�R�[�h�C���|�[�g
Private Sub ImportModules(Optional name As String)
    If name = "" Then Exit Sub
    '
    Dim base As String
    base = fso.GetBaseName(name)
    '
    Dim wb As Workbook
    Set wb = Application.Workbooks(name)
    '
    On Error Resume Next
    Dim col As Object     'VBComponents
    Set col = wb.VBProject.VBComponents
    On Error GoTo 0
    If col Is Nothing Then
        MsgBox "�uVBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ�" & _
                "�A�N�Z�X��M������v�ɐݒ肵�Ă��������B"
        Exit Sub
    End If
    '
    On Error Resume Next
    Dim path As String
    path = fso.BuildPath(ActiveWorkbook.path, base)
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "�\�[�X�t�@�C��", "*.bas;*.cls;*.frm"
        .Filters.Add "���ׂẴt�@�C��", "*.*"
        .FilterIndex = 0
        .InitialFileName = path
        .Show
        Dim fi As Variant
        For Each fi In .SelectedItems
            Dim s As String
            s = fso.GetBaseName(fi)
            Select Case LCase(fso.GetExtensionName(fi))
            Case "bas": col.Remove col(s): col.Import fi
            Case "frm": col.Remove col(s): col.Import fi
            Case "cls": col.Remove col(s): col.Import fi
            Case "cls2": col(s).Item.CodeModule.AddFromFile fi
            End Select
        Next fi
    End With
    On Error GoTo 0
End Sub

'�A�h�C���ēǂݍ���
Private Sub ReloadAddin(name As String)
    If name = "" Then Exit Sub
    If name Like ThisWorkbook.name Then
        MsgBox name & " �A�h�C���͍ēǂݍ��݂ł��܂���B"
        Application.EnableEvents = False
        Exit Sub
    End If
    '
    Dim ai As addin
    For Each ai In AddIns
        If ai.name Like name Then Exit For
    Next ai
    If ai Is Nothing Then Exit Sub
    ai.Installed = False
    ai.Installed = True
End Sub

'�_�C�A���O�Ăяo��
Private Sub CallDialog()
    On Error Resume Next
    If Application.Dialogs(ActiveCell.Value).Show Then
        MsgBox True
    Else
        MsgBox False
    End If
    On Error GoTo 0
End Sub

'==================================
'����
'==================================

'----------------------------------
'�I�u�W�F�N�g�Ăяo��
'----------------------------------

'filesystemobject
Private Function fso() As Object
    Static obj As Object
    If obj Is Nothing Then
        Set obj = CreateObject("Scripting.FileSystemObject")
    End If
    Set fso = obj
End Function

'regex
Private Function regex( _
        ptn As String, _
        Optional g As Boolean = True, _
        Optional ic As Boolean = True) As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = g
        .IgnoreCase = ic
        .Pattern = ptn
    End With
End Function

'----------------------------------
'�i�s�󋵕\��(status-bar)
'----------------------------------

Private Sub ProgressStatusBar(Optional i As Long = 1, Optional cnt As Long = 1)
    Static tm_start As Double
    If i < 1 Then
        tm_start = Timer
        Application.StatusBar = "�i����(0%)"
        Exit Sub
    End If
    If i >= cnt Then
        Application.StatusBar = False
        Exit Sub
    End If
    Dim p As Double: p = i / cnt
    Dim s As String: s = "�i����(" & Int(p * 100) & "%)"
    s = s & " : " & ProgressBar(p)
    Dim tm As Double: tm = (Timer - tm_start) / p * (1 - p)
    Application.StatusBar = s & " : �c��" & Int(tm) & "�b"
End Sub

Private Function ProgressBar(p As Double) As String
    If p < 0.2 Then
        ProgressBar = "����������"
    ElseIf p < 0.4 Then
        ProgressBar = "����������"
    ElseIf p < 0.6 Then
        ProgressBar = "����������"
    ElseIf p < 0.8 Then
        ProgressBar = "����������"
    ElseIf p < 1 Then
        ProgressBar = "����������"
    Else
        ProgressBar = "����������"
    End If
End Function

'----------------------------------
'�A�h�C��
'----------------------------------

'���[�U�A�h�C���t�H���_�擾
Function AddinsPath() As String
    AddinsPath = ThisWorkbook.path
End Function

'�A�h�C�����擾
Function AddinName(Optional name As String) As String
    If name = "" Then name = ThisWorkbook.name
    AddinName = Replace(name, ".xlam", "")
End Function

'�A�h�C���u�b�N�\���E��\���g�O��
Sub ToggleAddin(Optional name As String)
    If name = "" Then name = ThisWorkbook.name
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks(name)
    If wb Is Nothing Then
        name = ThisWorkbook.path & "\" & name
        Set wb = Application.Workbooks.Open(name)
    End If
    On Error GoTo 0
    If wb.IsAddin Then
        wb.IsAddin = False
        wb.Activate
    Else
        wb.IsAddin = True
        wb.Save
    End If
End Sub

'==================================
'�A�h�C���J��
'==================================

'���[�U�A�h�C���t�H���_�̃A�h�C�����J�E���g
Function UserAddinCount() As Integer
    Dim cnt As Integer
    Dim obj As addin
    For Each obj In AddIns
        If obj.path = AddinsPath Then
            If LCase(Right(obj.name, 5)) = ".xlam" Then
                cnt = cnt + 1
            End If
        End If
    Next obj
    UserAddinCount = cnt
End Function

'���[�U�A�h�C���t�H���_�̃A�h�C�����擾
Function UserAddinName(index As Integer) As String
    Dim cnt As Integer
    Dim obj As addin
    For Each obj In AddIns
        If obj.path = AddinsPath Then
            If LCase(Right(obj.name, 5)) = ".xlam" Then
                cnt = cnt + 1
                If cnt = index + 1 Then UserAddinName = obj.name
            End If
        End If
    Next obj
End Function

'�J�����g���[�U�A�h�C��ID�擾
Function CurrentAddinID() As Integer
    Dim cnt As Integer
    Dim cnt2 As Integer
    Dim obj As addin
    For Each obj In AddIns
        If obj.path = AddinsPath Then
            If LCase(Right(obj.name, 5)) = ".xlam" Then
                If obj.name = g_addin Then
                    CurrentAddinID = cnt
                    Exit Function
                ElseIf Not obj.name = ThisWorkbook.name Then
                    cnt2 = cnt + 1
                End If
                cnt = cnt + 1
            End If
        End If
    Next obj
    If cnt2 > 0 Then cnt = cnt2 - 1
    CurrentAddinID = cnt
End Function

