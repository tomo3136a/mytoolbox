Attribute VB_Name = "DataList"
'==================================
'�f�[�^���X�g�擾
'==================================

Option Explicit
Option Private Module

Private g_sheet As Boolean

'----------------------------------
'���ʋ@�\
'----------------------------------

Sub SetDataListParam(id As Integer, ByVal val As String)
    Select Case id
    Case 1
        g_sheet = val
    End Select
End Sub

Function GetDataListParam(id As Integer) As String
    Dim val As String
    Select Case id
    Case 1
        val = g_sheet
    End Select
    GetDataListParam = val
End Function

'---------------------------------------------
'����
'---------------------------------------------

'�v���p�e�B
Private Sub SetInfoSheet(Optional ws As Worksheet, Optional v As String = "")
    If ws Is Nothing Then Set ws = ActiveSheet
    ws.CustomProperties.Add "info", v
End Sub

Private Function TestInfoSheet(ws As Worksheet) As Boolean
    If SheetPropertyIndex(ws, "info") > 0 Then TestInfoSheet = True
End Function

'---------------------------------------------
'��񃊃X�g�A�b�v
'---------------------------------------------

Sub AddInfoSheet(mode As Integer)
    Application.ScreenUpdating = False
    '
    Dim Title As String
    Title = GetInfoTitle(mode)
    '
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim ra As Range
    If g_sheet Then
        Dim ws As Worksheet
        Set ws = wb.Worksheets.Add
        Call SetInfoSheet(ws, Title)
        If Title <> "" Then ws.name = UniqueSheetName(wb, Title)
        Set ra = ws.Cells(2, 2)
    Else
        Set ra = ActiveCell
    End If
    '
    Call AddInfoTable(mode, ra, Title, wb)
    '
    Application.ScreenUpdating = True
End Sub

Private Function GetInfoTitle(mode As Integer) As String
    Dim s As String
    Select Case mode
    Case 1
        s = "�ڎ�"
    Case 2
        s = "�V�[�g�ꗗ"
    Case 3
        s = "���O�ꗗ"
    Case 4
        s = "�����N�ꗗ"
    Case 5
        s = "�v���p�e�B�ꗗ"
    Case 6
        s = "�R�����g�ꗗ"
    Case 7
        s = "�t�@�C���ꗗ"
    End Select
    GetInfoTitle = s
End Function

Private Sub AddInfoTable(mode As Integer, ByRef ra As Range, _
        Optional Title As String, Optional wb As Workbook)
    Dim hd As Range
    Set hd = ra.Offset(0)
    If Title <> "" Then Set hd = hd.Offset(1)
    Dim ce As Range
    Set ce = hd.Offset(1)
    '
    Dim r As Long
    r = ce.Row
    '
    Dim h As Variant
    Select Case mode
    Case 1
        h = Array("�ԍ�", "���O", "�����N", "����")
        Call IndexList(ce, wb)
    Case 2
        h = Array("�ԍ�", "�V�[�g��", "���", "�g�p�͈�", _
            "�e�[�u����", "�O���t��", "�}�`��", "���O��", _
            "�����N��", "�R�����g��", "�v���p�e�B��", _
            "����(�W��)", "��(�W��)")
        Call SheetList(ce, wb)
    Case 3
        h = Array("�ԍ�", "���O", "���", "�Q�Ɣ͈�", _
            "�l", "���", "�͈�", "�R�����g")
        Call NameList(ce, wb)
    Case 4
        h = Array("�ԍ�", "�V�[�g", "���", "�����N��", _
            "�\��������", "�����N��", "�q���g")
        Call LinkList(ce, wb)
    Case 5
        h = Array("�ԍ�", "���O", "���", "�Q�Ɣ͈�", _
            "�l", "���", "�͈�", "�R�����g")
        Call PropList(ce, wb)
    Case 6
        h = Array("�ԍ�", "�V�[�g", "���", _
            "�Q�Ɣ͈�", "�l", "�\��")
        Call CommentList(ce, wb)
    Case 7
        Dim path As String
        path = ActiveWorkbook.path
        path = SelectFolder(path)
        If path = "" Then Exit Sub
        h = Array("�ԍ�", "���O", "���", "���", "�T�C�Y", "�ҏW��")
        Set hd = hd.Offset(1)
        Set ce = ce.Offset(1)
        Call FileList(ce, path)
        If r < ce.Row Then hd.Offset(-1).Value = GetShortPath(path)
    End Select
    '
    If r = ce.Row Then Exit Sub
    If Title <> "" Then ra.Value = Title
    Range(hd, hd.Offset(0, UBound(h))) = h
    Call Waku(hd, fit:=True)
End Sub

'---------------------------------------------
'�e�[�u���쐬
'---------------------------------------------

'�ڎ��ꗗ
Private Sub IndexList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If TestInfoSheet(ws) Then
        ElseIf ws.Visible Then
            no = no + 1
            ra.Value = no
            ra.Offset(0, 1).Value = ws.name
            ra.Worksheet.Hyperlinks.Add _
                Anchor:=ra.Offset(0, 2), _
                Address:="", _
                SubAddress:=(ws.name & "!A1"), _
                TextToDisplay:="�V�[�g", _
                ScreenTip:=ws.name
            Set ra = ra.Offset(1)
        End If
    Next ws
End Sub

'�V�[�g�ꗗ
Private Sub SheetList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    On Error Resume Next
    '
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.name = ActiveSheet.name Then
        Else
            no = no + 1
            ra.Value = no
            ra.Worksheet.Hyperlinks.Add _
                Anchor:=ra.Offset(0, 1), _
                Address:=wb.path & "\" & wb.name, _
                SubAddress:=ws.name & "!A1", _
                TextToDisplay:=ws.name
                ', _
                'ScreenTip:=ws.name
            If Not ws.Visible Then ra.Offset(0, 2).Value = "��\��"
            ra.Offset(0, 3).Value = ws.UsedRange.Address(False, False)
            ra.Offset(0, 4).Value = ws.Shapes.Count
            ra.Offset(0, 5).Value = ws.ChartObjects.Count
            ra.Offset(0, 6).Value = ws.Shapes.Count
            ra.Offset(0, 7).Value = ws.Names.Count
            ra.Offset(0, 8).Value = ws.Hyperlinks.Count
            ra.Offset(0, 9).Value = ws.Comments.Count
            ra.Offset(0, 10).Value = ws.CustomProperties.Count
            ra.Offset(0, 11).Value = ws.StandardHeight
            ra.Offset(0, 12).Value = ws.StandardWidth
            Set ra = ra.Offset(1)
        End If
    Next ws
    '
    On Error GoTo 0
End Sub

'���O�ꗗ
Private Sub NameList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    On Error Resume Next
    '
    Dim nm As name
    For Each nm In wb.Names
        Dim sts As String
        sts = ""
        no = no + 1
        ra.Value = no
        ra.Offset(0, 1).Value = nm.name
        If nm.Visible = False Then sts = "��\��"
        ra.Offset(0, 3).Value = "'" & nm.Value
        ra.Offset(0, 4).Value = StrRange(nm.Value)
        If err Then
            ra.Offset(0, 4).Value = "#REF!"
            sts = "�G���["
            err.Clear
        End If
        ra.Offset(0, 5).Value = TypeName(nm.Parent)
        ra.Offset(0, 6).Value = nm.Parent.name
        ra.Offset(0, 7).Value = nm.Comment
        If sts <> "" Then ra.Offset(0, 2).Value = sts
        Set ra = ra.Offset(1)
    Next nm
    '
    On Error GoTo 0
End Sub

'�����N�ꗗ
Private Sub LinkList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If TestInfoSheet(ws) Then
        Else
            Dim lnk As Hyperlink
            For Each lnk In ws.Hyperlinks
                no = no + 1
                ra.Value = no
                Call HyperlinkInfo(ra.Offset(0, 1), ws, lnk)
                Set ra = ra.Offset(1)
            Next lnk
        End If
    Next ws
End Sub

Private Sub HyperlinkInfo(ra As Range, ws As Worksheet, lnk As Hyperlink)
    Dim sts As String
    Dim src_addr As String
    Dim src_disp As String
    If lnk.Type = 0 Then
        src_addr = lnk.Range.Address
        src_disp = lnk.TextToDisplay
    Else
        src_addr = lnk.Shape.TopLeftCell.Address
        src_disp = "[" & lnk.Shape.name & "]"
    End If
    Dim src As String
    src = "'" & ws.name & "'!" & src_addr
    '
    Dim dst_addr As String
    Dim dst_disp As String
    '
    On Error Resume Next
    ra.Value = ws.name
    ActiveSheet.Hyperlinks.Add Anchor:=ra.Offset(0, 2), _
        Address:="", SubAddress:=src, TextToDisplay:=src_addr
    If err Then sts = "Error": err.Clear
    ra.Offset(0, 3).Value = src_disp
    '
    If lnk.SubAddress = "" Then
        ActiveSheet.Hyperlinks.Add Anchor:=ra.Offset(0, 4), _
            Address:=lnk.Address, TextToDisplay:=lnk.Address
    ElseIf lnk.Address = "" Then
        ActiveSheet.Hyperlinks.Add Anchor:=ra.Offset(0, 4), _
            Address:="", SubAddress:=lnk.SubAddress, TextToDisplay:=lnk.SubAddress
    Else
        ActiveSheet.Hyperlinks.Add Anchor:=ra.Offset(0, 4), _
            Address:=lnk.Address, SubAddress:=lnk.SubAddress, _
            TextToDisplay:="'" & lnk.Address & "'!" & lnk.SubAddress
    End If
    If err Then sts = "Error": err.Clear
    ra.Offset(0, 5).Value = lnk.ScreenTip
    '
    If lnk.Address <> "" Then
        Dim wb As Workbook: Set wb = ws.Parent
        Dim path As String: path = wb.path & "\" & lnk.Address
        If Dir(path) = "" Then sts = "�����N�؂�(" & lnk.Address & ")"
        If err Then sts = "�s��": err.Clear
    End If
    '
    If sts Then ra.Offset(0, 1).Value = sts
    On Error GoTo 0
End Sub

'�v���p�e�B�ꗗ
Private Sub PropList(ByRef ra As Range, wb As Workbook)
    Dim ts() As Variant
    ts = Array("", "����", "�_���l", "���t", "������", "����")
    '
    Dim no As Integer
    Dim sts As String
    Dim dp As DocumentProperty
    On Error Resume Next
    '
    'BuiltinDocumentProperties
    For Each dp In wb.BuiltinDocumentProperties
        If dp Is Nothing Then
        Else
            sts = ""
            no = no + 1
            ra.Value = no
            ra.Offset(0, 1).Value = dp.name
            ra.Offset(0, 3).Value = ts(dp.Type)
            ra.Offset(0, 4).Value = dp.Value
            ra.Offset(0, 5).Value = "�g�ݍ���"
            If err Then sts = "�G���[": err.Clear
            ra.Offset(0, 2).Value = sts
            Set ra = ra.Offset(1)
        End If
    Next dp
    '
    'customDocumentPropertoes
    For Each dp In wb.customDocumentPropertoes
        If dp Is Nothing Then
        Else
            sts = ""
            no = no + 1
            ra.Value = no
            ra.Offset(0, 1).Value = dp.name
            ra.Offset(0, 3).Value = ts(dp.Type)
            ra.Offset(0, 4).Value = dp.Value
            ra.Offset(0, 5).Value = "�J�X�^��"
            If err Then sts = "�G���[": err.Clear
            ra.Offset(0, 2).Value = sts
            Set ra = ra.Offset(1)
        End If
    Next dp
    '
    'CustomProperty
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If TestInfoSheet(ws) Then
        Else
            Dim cp As CustomProperty
            For Each cp In ws.CustomProperties
                sts = ""
                no = no + 1
                ra.Value = no
                ra.Offset(0, 1).Value = cp.name
                ra.Offset(0, 3).Value = TypeName(cp.Value)
                ra.Offset(0, 4).Value = cp.Value
                ra.Offset(0, 5).Value = ws.name
                If err Then sts = "�G���[": err.Clear
                ra.Offset(0, 2).Value = sts
                Set ra = ra.Offset(1)
            Next cp
        End If
    Next ws
    '
    On Error GoTo 0
End Sub

'�R�����g�ꗗ
Private Sub CommentList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    Dim cm As Comment
    On Error Resume Next
    '
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        For Each cm In ws.Comments
            no = no + 1
            ra.Value = no
            ra.Offset(0, 1).Value = ws.name
            If TypeName(cm.Parent) = "Range" Then
                ra.Offset(0, 3).Value = cm.Parent.Address
            End If
            ra.Offset(0, 4).Value = cm.text
            If cm.Visible Then
                ra.Offset(0, 5).Value = "�\��"
            End If
            Set ra = ra.Offset(1)
        Next cm
    Next ws
    '
    On Error GoTo 0
End Sub

'�t�@�C���ꗗ
Private Sub FileList(ByRef ra As Range, path As String)
    If path = "" Then Exit Sub
    Dim no As Integer
    On Error Resume Next
    '
    Dim obj As Variant
    For Each obj In fso.GetFolder(path).SubFolders
        no = no + 1
        ra.Value = no
        With obj
            ra.Offset(0, 1).Value = .name
            ra.Offset(0, 2).Value = ""
            ra.Offset(0, 3).Value = "�t�H���_"
            ra.Offset(0, 4).Value = .Size
            ra.Offset(0, 4).Style = "Comma [0]"
            ra.Offset(0, 5).Value = .DateLastModified
            ra.Offset(0, 5).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss;@"
        End With
        Set ra = ra.Offset(1)
    Next obj
    '
    For Each obj In fso.GetFolder(path).Files
        no = no + 1
        ra.Value = no
        With obj
            ra.Offset(0, 1).Value = .name
            ra.Offset(0, 2).Value = ""
            ra.Offset(0, 3).Value = "�t�@�C��"
            ra.Offset(0, 4).Value = .Size
            ra.Offset(0, 4).Style = "Comma [0]"
            ra.Offset(0, 5).Value = .DateLastModified
            ra.Offset(0, 5).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss;@"
        End With
        Set ra = ra.Offset(1)
    Next obj
    '
    On Error GoTo 0
End Sub

