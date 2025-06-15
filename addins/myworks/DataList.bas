Attribute VB_Name = "DataList"
'==================================
'�f�[�^���X�g�擾
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'�@�\�Ăяo��
'----------------------------------------

'��񃊃X�g�A�b�v
'[parameter]
' info.sheet �t���O���L���Ȃ�V�[�g��ǉ����ď�������
' info.info ���V�[�g�̓��e�����X�g�Ώ�
'[option]
' mode=1: �ڎ�
'      2: �V�[�g�ꗗ
'      3: ���O�ꗗ
'      4: �����N�ꗗ
'      5: �v���p�e�B�ꗗ
'      6: �m�[�g�ꗗ
'      7: �R�����g�ꗗ
'      8:
'      9: �t�@�C���ꗗ

Sub AddInfoTable(mode As Integer)
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    '
    ScreenUpdateOff
    '
    '�V�[�g�I��
    Dim ce As Range
    Set ce = ActiveCell
    If GetBookBool("info.sheet") Then
        Dim ws As Worksheet
        Set ws = wb.Worksheets.Add
        ws.name = UniqueSheetName(wb, GetInfoTitle(mode))
        Call SetInfoSheet(ws, "1")
        Set ce = ws.Cells(2, 2)
    End If
    '
    '���o��
    ce.Value = GetInfoTitle(mode)
    Set ce = ce.Offset(2)
    Dim tc As Range
    Set tc = ce
    '
    '�e�[�u���w�b�_�[
    Dim h As Variant
    h = GetInfoHeader(mode)
    Dim hd As Range
    Set hd = ce.Resize(1, UBound(h) - LBound(h) + 1)
    hd.Value = h
    Set ce = ce.Offset(1)
    '
    '�e�[�u���f�[�^
    Select Case mode
    Case 1: Call IndexList(ce, wb)      '�ڎ�
    Case 2: Call SheetList(ce, wb)      '�V�[�g�ꗗ
    Case 3: Call NameList(ce, wb)       '���O�ꗗ
    Case 4: Call LinkList(ce, wb)       '�����N�ꗗ
    Case 5: Call PropList(ce, wb)       '�v���p�e�B�ꗗ
    Case 6: Call NoteList(ce, wb)       '�m�[�g�ꗗ
    Case 7: Call CommentList(ce, wb)    '�R�����g�ꗗ
    Case 8:
    Case 9: Call FileList(ce, ActiveWorkbook.path)
    Case 10
    End Select
    '
    '�g�\��
    Call Waku(tc, fit:=True)
    '
    ScreenUpdateOn
End Sub

'��񃊃X�g�̃^�C�g��
Private Function GetInfoTitle(mode As Integer) As String
    Dim s As String
    Select Case mode
    Case 1: s = "�ڎ�"
    Case 2: s = "�V�[�g�ꗗ"
    Case 3: s = "���O�ꗗ"
    Case 4: s = "�����N�ꗗ"
    Case 5: s = "�v���p�e�B�ꗗ"
    Case 6: s = "�m�[�g�ꗗ"
    Case 7: s = "�R�����g�ꗗ"
    Case 8: s = "�ꗗ"
    Case 9: s = "�t�@�C���ꗗ"
    Case Else: s = "�ꗗ"
    End Select
    GetInfoTitle = s
End Function

Private Function GetInfoHeader(mode As Integer) As Variant
    Dim va As Variant
    Select Case mode
    Case 1: va = Array("�ԍ�", "���O", "�����N", "����")
    Case 2: va = Array("�ԍ�", "�V�[�g��", "���", "�g�p�͈�", _
                "�e�[�u����", "�O���t��", "�}�`��", "���O��", _
                "�����N��", "�R�����g��", "�v���p�e�B��", _
                "����(�W��)", "��(�W��)")
    Case 3: va = Array("�ԍ�", "���O", "���", "�Q�Ɣ͈�", "�l", "���", "�͈�", "���l")
    Case 4: va = Array("�ԍ�", "�V�[�g", "�����N��", "���", "�\��������", "�����N��", "�q���g")
    Case 5: va = Array("�ԍ�", "���O", "���", "�Q�Ɣ͈�", "�l", "���", "�͈�", "���l")
    Case 6: va = Array("�ԍ�", "�R�����g", "���", "���", "�V�[�g", "�Q�Ɣ͈�", "�l", "�͈�", "���l")
    Case 7: va = Array("�ԍ�", "�V�[�g", "�Q�Ɣ͈�", "�l", "���", "�R�����g", "�L����", "�L����", "���l")
    Case 8: va = Array("�ԍ�", "�V�[�g", "���", "�Q�Ɣ͈�", "�l", "�\��")
    Case 9: va = Array("�ԍ�", "���O", "�K�w", "���", "�T�C�Y", "�ҏW��")
    End Select
    GetInfoHeader = va
End Function

'----------------------------------
'�ڎ��ꗗ
'----------------------------------

Private Sub IndexList(ByRef ra As Range, wb As Workbook)
    Dim i As Long, j As Long, m As Long, n As Long
    Dim va(1 To 20, 1 To 4) As Variant
    m = UBound(va, 1)
    n = UBound(va, 2)
    '
    Dim ce As Range
    Set ce = ra(1, 1)
    '
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not TestInfoSheet(ws) Then
            If ws.Visible Then
                i = i + 1
                j = j + 1
                va(i, 1) = j
                va(i, 2) = ws.name
                va(i, 3) = "''" & ws.name & "'!A1"
                If i = m Then
                    ce.Resize(i, n).Value = va
                    Set ce = ce.Offset(i)
                    i = 0
                End If
            End If
        End If
    Next ws
    If i > 0 Then ce.Resize(i, n).Value = va
    '
    'hyperlink
    Set ws = ra.Worksheet
    If j > 0 Then
        For Each ce In ra.Offset(0, 2).Resize(j, 1)
            ws.Hyperlinks.Add _
                Anchor:=ce, _
                Address:="", _
                SubAddress:=ce.Value, _
                TextToDisplay:="�V�[�g", _
                ScreenTip:=ce.Offset(, -1).Value
        Next ce
    End If
    '
    Set ra = ra.Offset(j)
End Sub

'----------------------------------
'�V�[�g�ꗗ
'----------------------------------

Private Sub SheetList(ByRef ra As Range, wb As Workbook)
    Dim i As Long, j As Long, m As Long, n As Long
    Dim va(1 To 20, 1 To 13) As Variant
    m = UBound(va, 1)
    n = UBound(va, 2)
    '
    Dim ce As Range
    Set ce = ra(1, 1)
    '
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.name = ActiveSheet.name Then
        Else
            i = i + 1
            j = j + 1
            va(i, 1) = j
            va(i, 2) = ws.name
            va(i, 3) = IIf(ws.Visible, Empty, "��\��")
            va(i, 4) = ws.UsedRange.Address(False, False)
            va(i, 5) = ws.Shapes.Count
            va(i, 6) = ws.ChartObjects.Count
            va(i, 7) = ws.Shapes.Count
            va(i, 8) = ws.Names.Count
            va(i, 9) = ws.Hyperlinks.Count
            va(i, 10) = ws.Comments.Count
            va(i, 11) = ws.CustomProperties.Count
            va(i, 12) = ws.StandardHeight
            va(i, 13) = ws.StandardWidth
            If i = m Then
                ce.Resize(i, n).Value = va
                Set ce = ce.Offset(i)
                i = 0
            End If
        End If
    Next ws
    If i > 0 Then ce.Resize(i, n).Value = va
    '
    'hyperlink
    Set ws = ra.Worksheet
    If j > 0 Then
        For Each ce In ra.Offset(0, 1).Resize(j, 1)
            ws.Hyperlinks.Add _
                Anchor:=ce, _
                Address:=wb.path & "\" & wb.name, _
                SubAddress:="'" & ce.Value & "'!A1", _
                TextToDisplay:=ce.Value, _
                ScreenTip:=wb.path & Chr(10) & wb.name & Chr(10) & ce.Value & "!A1"
        Next ce
    End If
    '
    Set ra = ra.Offset(j)
End Sub

'----------------------------------
'���O�ꗗ
'----------------------------------

Private Sub NameList(ByRef ra As Range, wb As Workbook)
    Dim i As Long, j As Long, m As Long, n As Long
    Dim va(1 To 20, 1 To 8) As Variant
    m = UBound(va, 1)
    n = UBound(va, 2)
    '
    Dim ce As Range
    Set ce = ra(1, 1)
    '
    On Error Resume Next
    Dim nm As name
    For Each nm In wb.Names
        i = i + 1
        j = j + 1
        va(i, 1) = j
        va(i, 2) = nm.name
        va(i, 3) = IIf(nm.Visible, Empty, "��\��")
        va(i, 4) = "'" & nm.Value
        va(i, 5) = StrRange(nm.Value)
        If err Then
            va(i, 5) = "#REF!"
            va(i, 3) = "�G���["
            err.Clear
        End If
        va(i, 6) = TypeName(nm.Parent)
        va(i, 7) = nm.Parent.name
        va(i, 8) = nm.Comment
        If i = m Then
            ce.Resize(i, n).Value = va
            Set ce = ce.Offset(i)
            i = 0
        End If
    Next nm
    If i > 0 Then ce.Resize(i, n).Value = va
    On Error GoTo 0
    '
    Set ra = ra.Offset(j)
End Sub

'----------------------------------
'�����N�ꗗ
'----------------------------------

Private Sub LinkList(ByRef ra As Range, wb As Workbook)
    Dim no As Integer
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not TestInfoSheet(ws) Then
            Dim lnk As Hyperlink
            For Each lnk In ws.Hyperlinks
                no = no + 1
                ra.Value = no
                Call LinkListItem(ra.Offset(0, 1), lnk)
                Set ra = ra.Offset(1)
            Next lnk
        End If
    Next ws
    Set ra = ra.Offset(-no)
End Sub

Private Sub LinkListItem(ra As Range, lnk As Hyperlink)
    Dim sts As String
    sts = ""
    '
    Dim addr As String
    Dim disp As String
    Dim src As String
    If lnk.Type = msoHyperlinkRange Then
        addr = lnk.Range.Address
        disp = lnk.TextToDisplay
    Else
        addr = lnk.Shape.TopLeftCell.Address
        disp = "[" & lnk.Shape.name & "]"
    End If
    src = "'" & lnk.Parent.Parent.name & "'!" & addr
    '
    On Error Resume Next
    ra.Value = lnk.Parent.Parent.name
    ra.Worksheet.Hyperlinks.Add Anchor:=ra.Offset(0, 1), _
        Address:="", SubAddress:=src, TextToDisplay:=addr
    If err Then sts = "Error": err.Clear
    ra.Offset(0, 3).Value = disp
    If lnk.SubAddress = "" Then
        ra.Worksheet.Hyperlinks.Add Anchor:=ra.Offset(0, 4), _
            Address:=lnk.Address, TextToDisplay:=lnk.Address
    ElseIf lnk.Address = "" Then
        ra.Worksheet.Hyperlinks.Add Anchor:=ra.Offset(0, 4), _
            Address:="", SubAddress:=lnk.SubAddress, _
            TextToDisplay:=lnk.SubAddress
    Else
        ra.Worksheet.Hyperlinks.Add Anchor:=ra.Offset(0, 4), _
            Address:=lnk.Address, SubAddress:=lnk.SubAddress, _
            TextToDisplay:=lnk.Address & "#" & lnk.SubAddress
    End If
    If err Then sts = "Error": err.Clear
    ra.Offset(0, 5).Value = lnk.ScreenTip
    '
    If sts = "" And lnk.Address <> "" Then
        Dim path As String: path = lnk.Address
        If Not path Like "*:*" Then
            path = lnk.Parent.Parent.Parent.path & "\" & path
        End If
        If Dir(path) = "" Then sts = "�����N�؂�"
        If err Then sts = "�s��": err.Clear
    End If
    '
    If sts Then ra.Offset(0, 2).Value = sts
    On Error GoTo 0
End Sub

'----------------------------------
'�v���p�e�B�ꗗ
'----------------------------------

Private Sub PropList(ByRef ra As Range, wb As Workbook)
    Dim ts() As Variant
    ts = Array("", "����", "�_���l", "���t", "������", "����")
    '
    Dim i As Long, j As Long, m As Long, n As Long
    Dim va(1 To 20, 1 To 6) As Variant
    m = UBound(va, 1)
    n = UBound(va, 2)
    '
    Dim ce As Range
    Set ce = ra(1, 1)
    '
    Dim sts As String
    Dim dp As DocumentProperty
    Dim cp As CustomProperty
    On Error Resume Next
    '
    'BuiltinDocumentProperties
    For Each dp In wb.BuiltinDocumentProperties
        If Not dp Is Nothing Then
            sts = ""
            i = i + 1
            j = j + 1
            va(i, 1) = j
            va(i, 2) = dp.name
            va(i, 4) = ts(dp.Type)
            va(i, 5) = dp.Value
            va(i, 6) = "�g�ݍ���"
            If err Then sts = "�G���[": err.Clear
            va(i, 3) = sts
            If i = m Then
                ce.Resize(i, n).Value = va
                Set ce = ce.Offset(i)
                i = 0
            End If
        End If
    Next dp
    '
    'customDocumentPropertoes
    For Each dp In wb.customDocumentPropertoes
        If Not dp Is Nothing Then
            sts = ""
            i = i + 1
            j = j + 1
            va(i, 1) = j
            va(i, 2) = dp.name
            va(i, 4) = ts(dp.Type)
            va(i, 5) = dp.Value
            va(i, 6) = "�J�X�^��"
            If err Then sts = "�G���[": err.Clear
            va(i, 3) = sts
            If i = m Then
                ce.Resize(i, n).Value = va
                Set ce = ce.Offset(i)
                i = 0
            End If
        End If
    Next dp
    '
    'CustomProperty
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not TestInfoSheet(ws) Then
            For Each cp In ws.CustomProperties
                sts = ""
                i = i + 1
                j = j + 1
                va(i, 1) = j
                va(i, 2) = cp.name
                va(i, 4) = TypeName(cp.Value)
                va(i, 5) = cp.Value
                va(i, 6) = ws.name
                If err Then sts = "�G���[": err.Clear
                va(i, 3) = sts
                If i = m Then
                    ce.Resize(i, n).Value = va
                    Set ce = ce.Offset(i)
                    i = 0
                End If
            Next cp
        End If
    Next ws
    If i > 0 Then ce.Resize(i, n).Value = va
    '
    On Error GoTo 0
    Set ra = ra.Offset(j)
End Sub

'----------------------------------
'�m�[�g�ꗗ
'----------------------------------

Private Sub NoteList(ByRef ra As Range, wb As Workbook)
    Dim i As Long, j As Long, m As Long, n As Long
    Dim va(1 To 20, 1 To 8) As Variant
    m = UBound(va, 1)
    n = UBound(va, 2)
    '
    Dim ce As Range
    Set ce = ra(1, 1)
    '
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim cm As Comment
        For Each cm In ws.Comments
            i = i + 1
            j = j + 1
            va(i, 1) = j
            va(i, 2) = cm.Text
            va(i, 3) = IIf(cm.Visible, Empty, "��\��")
            va(i, 4) = TypeName(cm.Parent)
            va(i, 5) = ws.name
            If TypeName(cm.Parent) = "Range" Then
                Dim cc As Range
                Set cc = cm.Parent
                va(i, 6) = cc.Address
                va(i, 7) = cc.Value
            Else
                va(i, 6) = Empty
                va(i, 7) = Empty
            End If
            If i = m Then
                ce.Resize(i, n).Value = va
                Set ce = ce.Offset(i)
                i = 0
            End If
        Next cm
    Next ws
    If i > 0 Then ce.Resize(i, n).Value = va
    '
    'hyperlink
    Set ws = ra.Worksheet
    If j > 0 Then
        For Each ce In ra.Offset(0, 5).Resize(j, 1)
            ws.Hyperlinks.Add _
                Anchor:=ce, _
                Address:=wb.path & "\" & wb.name, _
                SubAddress:="'" & ce.Offset(0, -1).Value & "'!" & ce.Value, _
                TextToDisplay:=ce.Value
        Next ce
    End If
    '
    Set ra = ra.Offset(j)
End Sub

'----------------------------------
'�R�����g�ꗗ
'----------------------------------

Private Sub CommentList(ByRef ra As Range, wb As Workbook)
    Dim i As Long, j As Long, m As Long, n As Long
    Dim va(1 To 20, 1 To 9) As Variant
    m = UBound(va, 1)
    n = UBound(va, 2)
    '
    Dim ce As Range
    Set ce = ra(1, 1)
    '
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Call CommentListItem(ce, va, i, j, m, n, 1, ws, ws.CommentsThreaded)
    Next ws
    If i > 0 Then ce.Resize(i, n).Value = va
    '
    'hyperlink
    On Error Resume Next
    Set ws = ra.Worksheet
    If j > 0 Then
        For Each ce In ra.Offset(0, 2).Resize(j, 1)
            ws.Hyperlinks.Add _
                Anchor:=ce, _
                Address:=wb.path & "\" & wb.name, _
                SubAddress:="'" & ce.Offset(0, -1).Value & "'!" & ce.Value, _
                TextToDisplay:=ce.Value
        Next ce
    End If
    On Error GoTo 0
    '
    Set ra = ra.Offset(j)
End Sub

Private Sub CommentListItem( _
        ByRef ce As Range, ByRef va As Variant, ByRef i As Long, ByRef j As Long, _
        m As Long, n As Long, ByVal d As Long, ws As Worksheet, cts As CommentsThreaded)
    On Error Resume Next
    
    Dim ct As CommentThreaded
    For Each ct In cts
        i = i + 1
        j = j + 1
        va(i, 1) = j
        If TypeName(ct.Parent) = "Range" Then
            va(i, 2) = ws.name
            va(i, 3) = ct.Parent.Address
            va(i, 4) = ct.Parent.Value
        Else
            va(i, 2) = Empty
            va(i, 3) = Empty
            va(i, 4) = Empty
        End If
        va(i, 5) = IIf(ct.Resolved, "�ς�", Empty)
        va(i, 6) = ct.Text
        va(i, 7) = ct.Author.name
        va(i, 8) = ct.Date
        If i = m Then
            ce.Resize(i, n).Value = va
            Set ce = ce.Offset(i)
            i = 0
        End If
        If Not ct.Replies Is Nothing Then
            If ct.Replies.Count > 0 Then
                Call CommentListItem(ce, va, i, j, m, n, d + 1, ws, ct.Replies)
            End If
        End If
    Next ct
    
    On Error GoTo 0
End Sub

'----------------------------------
'�t�@�C���ꗗ
'----------------------------------
'0: �K�w�\��
'1: ��΃p�X
'2: ���΃p�X

Private Sub FileList(ByRef ra As Range, path As String)
    Dim v As Variant
    v = Application.InputBox("�^�C�v����͂��Ă��������B(0: �K�w�\��, 1: ��΃p�X, 2: ���΃p�X)", Type:=1)
    Call SetBookNum("path.5", CLng(v))
    '
    path = SelectFolder(path)
    If path = "" Then Exit Sub
    '
    Dim i As Long, j As Long, m As Long, n As Long
    Dim va(1 To 20, 1 To 9) As Variant
    m = UBound(va, 1)
    n = UBound(va, 2)
    '
    Dim ce As Range
    Set ce = ra(1, 1)
    '
    Dim p As String
    p = GetShortPath(path)
    path = GetAbstructPath(p, ce.Parent.Parent.path & "\")
    Dim sp As String
    Select Case GetBookNum("path.5")
    Case 1: sp = Replace(path, "\", "/") & "/"
    Case 2: sp = ""
    Case Else: sp = ""
    End Select
    Call FileListSubFolder(ce, path, va, i, j, m, n, 1, sp)
    If i > 0 Then ce.Resize(i, n).Value = va
    '
    ra.Offset(0, 4).Resize(j, 1).Style = "Comma [0]"
    ra.Offset(0, 5).Resize(j, 1).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss;@"
    '
    Set ra = ra.Offset(j)
End Sub

Private Sub FileListSubFolder( _
        ByRef ce As Range, path As String, _
        ByRef va As Variant, ByRef i As Long, ByRef j As Long, _
        m As Long, n As Long, ByVal d As Long, ByVal sp As String)
    Dim re As Object
    Set re = regex("^[~._]|bak$|tmp$")
    '
    Dim obj As Variant
    For Each obj In fso.GetFolder(path).SubFolders
        If GetBookBool("path.4") Or Not re.Test(obj.name) Then
            If GetBookBool("path.2") Then
                i = i + 1
                j = j + 1
                va(i, 1) = j
                va(i, 2) = sp & obj.name & "/"
                va(i, 3) = d
                va(i, 4) = "�t�H���_"
                va(i, 5) = obj.Size
                va(i, 6) = obj.DateLastModified
                If i = m Then
                    ce.Resize(i, n).Value = va
                    Set ce = ce.Offset(i)
                    i = 0
                End If
                If GetBookBool("path.3") Then
                    Dim p As String
                    p = fso.BuildPath(path, obj.name)
                    If Left(obj.name, 1) = "." Then
                    ElseIf Left(obj.name, 1) = "_" Then
                    Else
                        Dim sp2 As String
                        Select Case GetBookNum("path.5")
                        Case 1: sp2 = sp & obj.name & "/"
                        Case 2: sp2 = sp & obj.name & "/"
                        Case Else: sp2 = sp + "    "
                        End Select
                        Call FileListSubFolder(ce, p, va, i, j, m, n, d + 1, sp2)
                    End If
                End If
            End If
        End If
    Next obj
    '
    For Each obj In fso.GetFolder(path).Files
        i = i + 1
        j = j + 1
        va(i, 1) = j
        va(i, 2) = sp & obj.name
        va(i, 3) = d
        va(i, 4) = "�t�@�C��"
        va(i, 5) = obj.Size
        va(i, 6) = obj.DateLastModified
        If i = m Then
            ce.Resize(i, n).Value = va
            Set ce = ce.Offset(i)
            i = 0
        End If
    Next obj
End Sub

'----------------------------------
'���ʋ@�\
'----------------------------------

'�v���p�e�B
Private Sub SetInfoSheet(Optional ws As Worksheet, Optional v As String = "")
    If ws Is Nothing Then Set ws = ActiveSheet
    ws.CustomProperties.Add "info", v
End Sub

Private Function TestInfoSheet(ws As Worksheet) As Boolean
    If GetBookBool("info.info") Then Exit Function
    If SheetPropIndex(ws, "info") > 0 Then TestInfoSheet = True
End Function

