Attribute VB_Name = "Dev"
'[�Q�Ɛݒ�]
'�uMicrosoft Visual Basic for Application Extensibility 5.3�v

Option Explicit
Option Private Module

Private g_addin As String

'==================================
'�A�h�C���J��
'==================================

Sub SetAddinName(s As String)
    g_addin = s
End Sub

'----------------------------------
'�A�v���P�[�V����I/F
'----------------------------------

Sub AddinDevApp(id As Integer)
    Select Case id
    Case 1
        '�A�h�C���t�H���_���J��
        OpenAddinsFolder
    Case 2
        'ImageMso�t�@�C���ۑ�
        SaveImageMso
    Case 31
        'CustomUI �ҏW
        EditCustomUI g_addin
    Case 32
        'CustomUI �}�[�W
        Workbooks(g_addin).Save
        MargeCustomUI g_addin
    Case 33
        '�A�h�C���z�u
        DeployAddin g_addin
    Case 34
        '�A�h�C���u�b�N�\���E��\���g�O��
        ToggleAddin g_addin
    Case 35
    Case 37
        '�A�h�C���\�[�X�̃G�N�X�|�[�g
        ExportModules g_addin
    Case 38
        '�A�h�C���\�[�X�̃C���|�[�g
        ImportModules g_addin
    Case 4
        '����
        ToggleAddin ActiveWorkbook.name
    Case Else
        MsgBox g_addin
    End Select
End Sub

'----------------------------------
'���\�[�X
'----------------------------------

'�A�h�C���t�H���_���J��
Private Sub OpenAddinsFolder()
    With CreateObject("Wscript.Shell")
        .Run AddinsPath
    End With
End Sub

'ImageMso�t�@�C���ۑ�
Private Sub SaveImageMso()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("AddinDev")
    If ws Is Nothing Then Exit Sub
    '
    Dim ra As Range
    Set ra = ws.UsedRange
    If ra Is Nothing Then Exit Sub
    If ra.Cells(1, 1).Value = "" Then Exit Sub
    '
    Dim cnt As Long
    cnt = ra.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    Dim arr
    arr = ws.Range("A2:A" & cnt).Value
    '
    Dim parent As String
    parent = Environ("USERPROFILE") & "\Documents"
    If parent = "" Then Exit Sub
    parent = fso.BuildPath(parent, "imageMso")
    If Not fso.FolderExists(parent) Then
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
    Dim zip As String
    zip = xlam & ".zip"
    fso.CopyFile xlam, zip
    '
    Dim src As String
    src = fso.BuildPath(zip, "CustomUI")
    '
    Dim dst As String
    dst = fso.BuildPath(AddinsPath, AddinName(name))
    If Not fso.FolderExists(dst) Then fso.CreateFolder dst
    dst = fso.BuildPath(dst, "CustomUI")
    If Not fso.FolderExists(dst) Then fso.CreateFolder dst
    '
    If Not fso.FileExists(fso.BuildPath(dst, "customUI.xml")) Then
        Dim shell As Object
        Set shell = CreateObject("Shell.Application")
        shell.Namespace(CVar(dst)).CopyHere shell.Namespace(CVar(src)).Items
        If Not shell Is Nothing Then
            Set shell = Nothing
        End If
    End If
    '
    Call CreateObject("Wscript.Shell").Run(dst & "\customUI.xml")
End Sub

'CustomUI �}�[�W
Private Sub MargeCustomUI(name As String)
    Dim xlam As String
    xlam = fso.BuildPath(AddinsPath, name)
    If Not fso.FileExists(xlam) Then Exit Sub
    '
    Dim zip As String
    zip = xlam & ".zip"
    fso.CopyFile xlam, zip
    '
    Dim src As String
    src = fso.BuildPath(AddinsPath, AddinName(name))
    src = fso.BuildPath(src, "CustomUI")
    If Not fso.FileExists(fso.BuildPath(src, "customUI.xml")) Then
        MsgBox "CustomUI.xml ������܂���B"
        Exit Sub
    End If
    '
    Dim dst As String
    dst = fso.BuildPath(zip, "CustomUI")
    '
    Dim shell As Object
    Set shell = CreateObject("Shell.Application")
    Dim fo As Object
    Set fo = shell.Namespace(CVar(dst))
    fo.CopyHere shell.Namespace(CVar(src)).Items
    If Not shell Is Nothing Then
        Set shell = Nothing
    End If
End Sub

'�A�h�C���z�u
Private Sub DeployAddin(name As String)
    If name = ThisWorkbook.name Then
        MsgBox name & " �A�h�C���͔z�u�ł��܂���B" & Chr(10) & _
        "zip �t�@�C���� xlam �t�@�C���ɕϊ�����" & _
        "�A�h�C�����ăC���X�g�[�����Ă��������B"
        OpenAddinsFolder
        AddIns(AddinName(name)).Installed = False
        Exit Sub
    End If
    
    Dim xlam As String
    xlam = fso.BuildPath(AddinsPath, name)
    '
    Dim zip As String
    zip = xlam & ".zip"
    If Not fso.FileExists(zip) Then Exit Sub
    '
    AddIns(AddinName(name)).Installed = False
    If fso.FileExists(xlam) Then fso.DeleteFile xlam
    fso.MoveFile zip, xlam
    AddIns(AddinName(name)).Installed = True
End Sub

'���W���[���\�[�X�R�[�h�G�N�X�|�[�g
Private Sub ExportModules(Optional name As String)
    If name = "" Then Exit Sub
    name = Replace(name, ".xlam", "")
    '
    On Error Resume Next
    Dim ai As addin
    Set ai = Application.AddIns(name)
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
        .Title = "�t�H���_�̑I��"
        If Not .Show Then Exit Sub
        path = .SelectedItems(1)
    End With
    path = fso.BuildPath(path, name)
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
            Case 1
                Call m.Export(path & "\" & m.name & ".bas")
            Case 2
                Call m.Export(path & "\" & m.name & ".cls")
            Case 3
                Call m.Export(path & "\" & m.name & ".frm")
            Case 100
                Call m.Export(path & "\" & m.name & ".cls")
            Case Else
                MsgBox "�G�N�X�|�[�g�ΏۊO�F " & m.Type & ":" & m.name
            End Select
        End If
    Next m
    '
    With CreateObject("Wscript.Shell")
        .Run path
    End With
End Sub

'���W���[���\�[�X�R�[�h�C���|�[�g
Private Sub ImportModules(Optional name As String)
    If name = "" Then Exit Sub
    name = fso.GetBaseName(name)
    name = Replace(name, ".xlam", "")
    '
    Dim wb As Workbook
    Set wb = Application.Workbooks(Application.AddIns(name).name)
    '
    On Error Resume Next
    Dim col As VBComponents
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
    path = fso.BuildPath(ActiveWorkbook.path, name)
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
            Case "bas"
                col.Remove col(s)
                col.Import fi
            Case "frm"
                col.Remove col(s)
                col.Import fi
            Case "cls"
                col.Remove col(s)
                col.Import fi
            Case "cls2"
                col(s).Item.CodeModule.AddFromFile fi
            End Select
        Next fi
    End With
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

