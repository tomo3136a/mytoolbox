Attribute VB_Name = "Report"
'==================================
'���|�[�g�ҏW�@�\
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'���|�[�g�ɃT�C��
'----------------------------------------

Sub ReportSign(ra As Range)
    Dim s As String
    Dim c As Long
    Dim r As Long
    s = Date & " " & Application.UserName
    c = ra.Column
    r = ra.Row
    '
    Do Until Cells(r, c).Value = ""
        r = r + 1
    Loop
    If r > 1 Then
        If Cells(r - 1, c).Value <> "" Then
            s = "�X�V " & s
        End If
    End If
    '
    With Cells(r, c)
        .HorizontalAlignment = xlRight
        .Value = s
    End With
End Sub

'----------------------------------------
'�y�[�W�t�H�[�}�b�g�ݒ�
'----------------------------------------

Sub PagePreview()
    Dim fsu As Boolean
    fsu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim ws As Worksheet
    For Each ws In ActiveWindow.SelectedSheets
        Call SheetPagePreview(ws)
    Next ws
    '
    Application.ScreenUpdating = fsu
End Sub

Private Sub SheetPagePreview(ws As Worksheet)
    Dim wnd As Window
    Set wnd = ws.Application.ActiveWindow
    wnd.Zoom = 100
    ws.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ws.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .CenterHorizontally = True
        .CenterVertically = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
    End With
    Application.PrintCommunication = True
    wnd.ActiveSheet.Cells(1, 1).Select
    wnd.ScrollColumn = 1
    wnd.ScrollRow = 1
    wnd.View = xlPageBreakPreview
    wnd.Zoom = 100
End Sub

'----------------------------------
'�e�L�X�g�ϊ�
'----------------------------------

Sub Cells_Conv(ra As Range, Optional mode As Integer = 0)
    Select Case mode
    Case 1
        '�g����(�璷�ȃX�y�[�X�폜)
        Call Cells_RemoveSpace(ra)
    Case 2
        '�V���O�����C��(�璷�ȃX�y�[�X�폜����1�s��)
        Call Cells_RemoveSpace(ra, SingleLine:=True)
    Case 3
        '�X�y�[�X�폜
        Call Cells_RemoveSpace(ra, sep:="")
    Case 4
        '������ύX(�啶���ɕϊ�)
        Call Cells_StrConv(ra, vbUpperCase)
    Case 5
        '������ύX(�������ɕϊ�)
        Call Cells_StrConv(ra, vbLowerCase)
    Case 6
        '������ύX(�e�P��̐擪�̕�����啶���ɕϊ�)
        Call Cells_StrConv(ra, vbProperCase)
    Case 7
        '������ύX(���p������S�p�����ɕϊ�)
        Call Cells_StrConv(ra, vbWide)
    Case 8
        '������ύX(�S�p�����𔼊p�����ɕϊ�)
        Call Cells_StrConv(ra, vbNarrow)
    Case 9
        '������ύX(ASCII�����̂ݔ��p��)
        Call Cells_StrConvNarrow(ra)
    Case Else
        '������ύX(ASCII�����̂ݔ��p���A�璷�ȃX�y�[�X�폜)
        Call Cells_StrConvNarrow(ra)
        Call Cells_RemoveSpace(ra)
    End Select
End Sub

'�X�y�[�X�폜
Private Sub Cells_RemoveSpace( _
        ra As Range, _
        Optional SingleLine As Boolean = False, _
        Optional sep As String = " ")
    Dim re1 As Object: Set re1 = regex("\s+")
    Dim re2 As Object: Set re2 = regex("[ �@\t]+")
    Dim re3 As Object: Set re3 = regex(" (\r?\n)")
    '
    Dim ce As Range
    For Each ce In ra.Cells
        If ce.Value <> "" Then
            If Not ce.HasFormula Then
                Dim s As String
                If SingleLine Then
                    s = re1.Replace(ce.Value, sep)
                Else
                    s = re2.Replace(ce.Value, sep)
                    s = re3.Replace(s, "$1")
                End If
                ce.Value = Trim(s)
            End If
        End If
    Next ce
End Sub

'������ύX
' vbUpperCase   1   �啶���ɕϊ�
' vbLowerCase   2   �������ɕϊ�
' vbProperCase  3   �e�P��̐擪�̕�����啶���ɕϊ�
' vbWide        4   ���p������S�p�����ɕϊ�
' vbNarrow      8   �S�p�����𔼊p�����ɕϊ�
' vbKatakana    16  �Ђ炪�Ȃ��J�^�J�i�ɕϊ�
' vbHiragana    32  �J�^�J�i���Ђ炪�Ȃɕϊ�
Private Sub Cells_StrConv(ra As Range, mode As Integer)
    Dim ce As Range
    For Each ce In ra.Cells
        If ce.Value <> "" Then
            If Not ce.HasFormula Then
                ce.Value = StrConv(ce.Value, mode)
            End If
        End If
    Next ce
End Sub

'������ύX(ASCII�����̂ݔ��p��)
Private Sub Cells_StrConvNarrow(ra As Range)
    Dim re As Object: Set re = regex("[�I-�`]+")
    '
    Dim ce As Range
    For Each ce In ra.Cells
        If ce.Value <> "" Then
            If Not ce.HasFormula Then
                Dim s As String
                s = ce.Value
                Dim m As Object
                For Each m In re.Execute(s)
                    s = Replace(s, m.Value, StrConv(m.Value, vbNarrow))
                Next m
                ce.Value = s
            End If
        End If
    Next ce
End Sub

'---------------------------------------------
'�\������
'---------------------------------------------

Sub ShowHide(mode As Integer)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    '
    Select Case mode
    Case 1
        '��\���s�폜�E��\����폜
        DeleteHideColumn
        DeleteHideRow
    Case 2
        '��\����폜
        DeleteHideColumn
    Case 3
        '��\���s�폜
        DeleteHideRow
    Case 4
        '��\���V�[�g�폜
        Call DeleteHideSheet
    Case 8
        '��\���V�[�g�\��
        Call VisibleHideSheet
    Case 9
        '��\�����O�\��
        Call VisibleHideName
    Case Else
    End Select
    '
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'��\���s�폜
Private Sub DeleteHideRow()
    Dim ra As Range
    Set ra = Selection
    Dim i As Long
    For i = Cells.SpecialCells(xlCellTypeLastCell).Row + 1 To 1 Step -1
        If Rows(i).Hidden Then Rows(i).Delete
    Next i
    On Error Resume Next
    ra.Select
    On Error GoTo 0
End Sub

'��\����폜
Private Sub DeleteHideColumn()
    Dim ra As Range
    Set ra = Selection
    Dim i As Long
    For i = Cells.SpecialCells(xlCellTypeLastCell).Column + 1 To 1 Step -1
        If Columns(i).Hidden Then Columns(i).Delete
    Next i
    On Error Resume Next
    ra.Select
    On Error GoTo 0
End Sub

'��\���V�[�g�폜
Private Sub DeleteHideSheet()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    '
    Dim cnt As Integer
    Dim msg As String
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = 0 Then
            Dim s As String
            s = ws.name
            ws.Delete
            cnt = cnt + 1
            msg = msg & vbCrLf & s
        End If
    Next ws
    '
    If cnt > 0 Then
        MsgBox cnt & "�V�[�g���폜���܂����B" & msg
    End If
End Sub

'��\���V�[�g�\��
Private Sub VisibleHideSheet()
    Dim ws As Worksheet
    Dim cnt As Integer
    Dim msg As String
    '
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = 0 Then
            ws.Visible = True
            cnt = cnt + 1
            msg = msg & vbCrLf & ws.name
        End If
    Next ws
    '
    If cnt > 0 Then
        MsgBox cnt & "�V�[�g��\���ɂ��܂����B" & msg
    End If
End Sub

'��\�����O�\��
Private Sub VisibleHideName()
    Dim nm As name
    Dim cnt As Integer
    Dim msg As String
    '
    For Each nm In ActiveWorkbook.Names
        If nm.Visible = 0 Then
            If Right(nm.name, 8) <> "Database" Then
                nm.Visible = True
                cnt = cnt + 1
                msg = msg & vbCrLf & nm.name
            End If
        End If
    Next nm
    '
    If cnt > 0 Then
        MsgBox "���O��" & cnt & "���\���ɂ��܂����B" & msg
    End If
End Sub

