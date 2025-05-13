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
    '�������ݐ�̋󔒒T��
    Do Until Cells(r, c).Value = ""
        r = r + 1
    Loop
    
    '�X�V�Ȃ�u�X�V�v�𖾋L
    If r > 1 Then
        If Cells(r - 1, c).Value <> "" Then
            s = "�X�V " & s
        End If
    End If
    
    '�E�l�߂ɂ��ăT�C����������
    With Cells(r, c)
        .HorizontalAlignment = xlRight
        .Value = s
    End With
End Sub

'----------------------------------------
'�y�[�W�t�H�[�}�b�g�ݒ�
'----------------------------------------

Sub PagePreview()
    ScreenUpdateOff
    '
    Dim ws As Worksheet
    For Each ws In ActiveWindow.SelectedSheets
        '�ŏI�s���󔒍s�łȂ��Ȃ�E���ɋ󔒒ǉ�
        Dim ra As Range
        Set ra = ws.UsedRange
        Set ra = ra(ra.Rows.Count, ra.Columns.Count)
        If ra <> " " Then ra.Offset(1, 0) = " "
        
        '����͈͕\���ɐݒ�        '
        Dim wnd As Window
        Set wnd = ws.Application.ActiveWindow
        wnd.Zoom = 100
        ws.PageSetup.PrintArea = ""
        Application.PrintCommunication = False
        With ws.PageSetup
            .Orientation = xlPortrait
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
    Next ws
    '
    ScreenUpdateOn
End Sub

'----------------------------------
'�e�L�X�g�ϊ�
' mode=1: �g����(�璷�ȃX�y�[�X�폜)
'      2: �V���O�����C��(�璷�ȃX�y�[�X�폜����1�s��)
'      3: �X�y�[�X�폜
'      4: ������ύX(�啶���ɕϊ�)
'      5: ������ύX(�������ɕϊ�)
'      6: ������ύX(�e�P��̐擪�̕�����啶���ɕϊ�)
'      7: ������ύX(���p������S�p�����ɕϊ�)
'      8: ������ύX(�S�p�����𔼊p�����ɕϊ�)
'      9: ������ύX(ASCII�����̂ݔ��p��)
'      *: ������ύX(ASCII�����̂ݔ��p���A�璷�ȃX�y�[�X�폜)
'----------------------------------

Sub TextConvProc(ra As Range, mode As Integer)
    ScreenUpdateOff
    '
    Dim rb As Range
    For Each rb In ra.Areas
        Set rb = Intersect(rb, ra.Parent.UsedRange)
        If rb Is Nothing Then Exit For
        '
        Dim va As Variant
        va = RangeToFormula2(rb)
        '
        Select Case mode
        Case 1: Call Cells_RemoveSpace(va)
        Case 2: Call Cells_RemoveSpace(va, SingleLine:=True)
        Case 3: Call Cells_RemoveSpace(va, sep:="")
        Case 4: Call Cells_StrConv(va, vbUpperCase)
        Case 5: Call Cells_StrConv(va, vbLowerCase)
        Case 6: Call Cells_StrConv(va, vbProperCase)
        Case 7: Call Cells_StrConv(va, vbWide)
        Case 8: Call Cells_StrConv(va, vbNarrow)
        Case 9: Call Cells_StrConvNarrow(va)
        Case Else
            Call Cells_StrConvNarrow(va)
            Call Cells_RemoveSpace(va)
        End Select
        '
        rb.Value = va
    Next rb
    '
    ScreenUpdateOn
End Sub

'�X�y�[�X�폜
Private Sub Cells_RemoveSpace( _
        va As Variant, _
        Optional SingleLine As Boolean = False, _
        Optional sep As String = " ")
    Dim re1 As Object, re2 As Object, re3 As Object
    Set re1 = regex("[\s\u00A0\u3000]+")
    Set re2 = regex("[ \t\v\f\u00A0\u3000]+")
    Set re3 = regex(sep & "(\r?\n)")
    '
    Dim r As Long, c As Long
    For r = LBound(va, 1) To UBound(va, 1)
        For c = LBound(va, 2) To UBound(va, 2)
            Dim s As String
            s = va(r, c)
            If s <> "" Then
                If SingleLine Then
                    s = re1.Replace(s, sep)
                Else
                    s = re2.Replace(s, sep)
                    s = re3.Replace(s, "$1")
                End If
                va(r, c) = Trim(s)
            End If
        Next c
    Next r
End Sub

'������ύX
' vbUpperCase   1   �啶���ɕϊ�
' vbLowerCase   2   �������ɕϊ�
' vbProperCase  3   �e�P��̐擪�̕�����啶���ɕϊ�
' vbWide        4   ���p������S�p�����ɕϊ�
' vbNarrow      8   �S�p�����𔼊p�����ɕϊ�
' vbKatakana    16  �Ђ炪�Ȃ��J�^�J�i�ɕϊ�
' vbHiragana    32  �J�^�J�i���Ђ炪�Ȃɕϊ�
Private Sub Cells_StrConv(va As Variant, mode As Integer)
    Dim s As String
    Dim r As Long, c As Long
    For r = LBound(va, 1) To UBound(va, 1)
        For c = LBound(va, 2) To UBound(va, 2)
            s = va(r, c)
            If s <> "" And Left(s, 1) <> "=" Then
                va(r, c) = StrConv(s, mode)
            End If
        Next c
    Next r
End Sub

'������ύX(ASCII�����̂ݔ��p��)
Private Sub Cells_StrConvNarrow(va As Variant)
    Dim re As Object: Set re = regex("[�I-�`]+")
    '
    Dim s As String
    Dim r As Long, c As Long
    For r = LBound(va, 1) To UBound(va, 1)
        For c = LBound(va, 2) To UBound(va, 2)
            s = va(r, c)
            If s <> "" And Left(s, 1) <> "=" Then
                Dim m As Object
                For Each m In re.Execute(s)
                    s = Replace(s, m.Value, StrConv(m.Value, vbNarrow))
                Next m
                va(r, c) = s
            End If
        Next c
    Next r
End Sub

Private Function RangeToFormula2(ra As Range) As Variant
    Dim va As Variant
    va = ra.Formula2
    If ra.Count = 1 Then
        ReDim va(1 To 1, 1 To 1)
        va(1, 1) = ra.Formula2
    End If
    RangeToFormula2 = va
End Function

'---------------------------------------------
'�\��/��\������
'mode=1: ��\���s�폜�E��\����폜
'     2: ��\����폜
'     3: ��\���s�폜
'     4: ��\���V�[�g�폜
'     8: ��\���V�[�g�\��
'     9: ��\�����O�\��
'---------------------------------------------

Sub ShowHide(mode As Integer)
    ScreenUpdateOff
    '
    Select Case mode
    Case 1: DeleteHideColumn: DeleteHideRow
    Case 2: DeleteHideColumn
    Case 3: DeleteHideRow
    Case 4: DeleteHideSheet
    Case 8: VisibleHideSheet
    Case 9: VisibleHideName
    End Select
    '
    ScreenUpdateOn
End Sub

'��\���s�폜
Private Sub DeleteHideRow()
    Dim ra As Range
    Set ra = Selection
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Dim e As Long
    e = ws.UsedRange.Row
    e = e + ws.UsedRange.Rows.Count
    Dim i As Long
    For i = e + 1 To 1 Step -1
        If ws.Rows(i).Hidden Then ws.Rows(i).Delete
    Next i
    On Error Resume Next
    ra.Select
    On Error GoTo 0
End Sub

'��\����폜
Private Sub DeleteHideColumn()
    Dim ra As Range
    Set ra = Selection
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Dim e As Long
    e = ws.UsedRange.Column
    e = e + ws.UsedRange.Columns.Count
    Dim i As Long
    For i = e + 1 To 1 Step -1
        If ws.Columns(i).Hidden Then ws.Columns(i).Delete
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

'---------------------------------------------
'��������
' mode=1: �����ɏ����t��������ǉ�
'      2: 0�ɏ����t��������ǉ�
'      3: �󔒂ɏ����t��������ǉ�
'      4: �Q�ƂɐF��t����
'      8: �Q�ƃX�^�C���폜
'---------------------------------------------

Sub UserFormatProc(ra As Range, mode As Long)
    Dim rb As Range
    Set rb = ra.Worksheet.UsedRange
    Set rb = Intersect(ra, rb)
    If rb Is Nothing Then Set rb = ra
    '
    Select Case mode
    Case 1: Call AddFormulaConditionColor(rb)
    Case 2: Call AddZeroConditionColor(rb)
    Case 3: Call AddBlankConditionColor(rb)
    Case 4: Call MarkRef(rb)
    Case 5: Call ListConditionFormat
    Case 8: Call ClearMarkRef
    End Select
End Sub

Private Sub AddFormulaConditionColor(ra As Range)
    '�������ݒ�
    Dim s As String
    s = ra.Cells(1, 1).Address(False, False)
    s = "=ISFORMULA(" & s & ")"
    With ra.FormatConditions
        .Add Type:=xlExpression, Formula1:=s
        .Item(.Count).SetFirstPriority
    End With
    
    '�������w�i�F�I��
    Dim i As Long
    Call Application.Dialogs(xlDialogEditColor).Show(1)
    i = ActiveWorkbook.Colors(1)

    '�������w�i�F�w��
    With ra.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = i
        .TintAndShade = 0
    End With
    ra.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub AddZeroConditionColor(ra As Range)
    '�������ݒ�
    Dim s As String
    s = ra.Cells(1, 1).Address(False, False)
    s = "=AND(" & s & "<>""""," & s & "=0)"
    With ra.FormatConditions
        .Add Type:=xlExpression, Formula1:=s
        .Item(.Count).SetFirstPriority
    End With
    
    '�������w�i�F�I��
    Dim i As Long
    Call Application.Dialogs(xlDialogEditColor).Show(1)
    i = ActiveWorkbook.Colors(1)

    '�������w�i�F�w��
    With ra.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = i
        .TintAndShade = 0
    End With
    ra.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub AddBlankConditionColor(ra As Range)
    '�������ݒ�
    Dim s As String
    s = ra.Cells(1, 1).Address(False, False)
    s = "=TRIM(" & s & ")="""""
    With ra.FormatConditions
        .Add Type:=xlExpression, Formula1:=s
        .Item(.Count).SetFirstPriority
    End With
    
    '�������w�i�F�I��
    Dim i As Long
    Call Application.Dialogs(xlDialogEditColor).Show(1)
    i = ActiveWorkbook.Colors(1)

    '�������w�i�F�w��
    With ra.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = i
        .TintAndShade = 0
    End With
    ra.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub MarkRef(ra As Range)
    Dim wb As Workbook
    Set wb = ra.Worksheet.Parent
    
    Dim rb As Range
    On Error Resume Next
    Set rb = ra.DirectPrecedents
    On Error GoTo 0
    If rb Is Nothing Then Exit Sub
    Set rb = rb.DirectDependents
    Set rb = Intersect(ra, rb)
    
    Dim s As String
    s = "�Q��"
    On Error Resume Next
    If wb.Styles(s) Is Nothing Then
        With wb.Styles.Add(s)
            .IncludeNumber = False
            .IncludeFont = False
            .IncludeAlignment = False
            .IncludeBorder = False
            .IncludePatterns = True
            .IncludeProtection = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = 0
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.8
                .PatternTintAndShade = 0
            End With
        End With
    End If
    On Error GoTo 0
    rb.Style = s
End Sub

Private Sub ClearMarkRef()
    Dim s As String
    s = "�Q��"
    On Error Resume Next
    ActiveWorkbook.Styles(s).Delete
    On Error GoTo 0
End Sub

Private Sub ListConditionFormat()
    Dim ra As Range
    Set ra = ActiveCell
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim fc As FormatCondition
    For Each fc In ws.Cells.FormatConditions
        ra.Value = "'" & fc.Formula1
        ra.Offset(, 1).Value = fc.PTCondition
        ra.Offset(, 2).Value = fc.AppliesTo.Address
        ra.Offset(, 3).Value = fc.Type
        Set ra = ra.Offset(1)
    Next fc
End Sub


Private Sub NewStyle(ra As Range, name As String)
    If ra Is Nothing Then Exit Sub
    If name = "" Then Exit Sub
    
    Dim wb As Workbook
    Set wb = ra.Parent.Parent
    On Error Resume Next
    If ra.Parent.Parent.Styles(name) Is Nothing Then
        With ra.Parent.Parent.Styles.Add(name)
            .IncludeNumber = False
            .IncludeFont = False
            .IncludeAlignment = False
            .IncludeBorder = False
            .IncludePatterns = True
            .IncludeProtection = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = 0
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.8
                .PatternTintAndShade = 0
            End With
        End With
    End If
    On Error GoTo 0
    ra.Style = name
End Sub

'---------------------------------------------
'��^���}��
'mode=1: �����񕪊�(�p���E���l)
'     2: �����񕪊�(���l�E�p���E���l)
'     3: �����}�[�J�[
'---------------------------------------------

Sub UserFormulaProc(ra As Range, mode As Integer)
    Dim v1 As Integer, v2 As Integer, v3 As Integer
    Dim r0 As Range, r1 As Range, r2 As Range, r3 As Range
    Select Case mode
    Case 1
        '�����񕪊�(�p���E���l)
        Set r0 = ra.Cells(1, 1)
        On Error Resume Next
        Set r1 = Application.InputBox("�L���ʒu", "�����񕪊�", r0.Offset(0, 1).Address, Type:=8)
        On Error GoTo 0
        If r1 Is Nothing Then Exit Sub
        v1 = ra.Column - r1.Column
        ra.Offset(, -v1).Formula2R1C1 = "=LET(v,RC[" & v1 & "],LEFT(v,MIN(FIND({1,2,3,4,5,6,7,8,9,0},v&""1234567890""))-1))"
        
        On Error Resume Next
        Set r2 = Application.InputBox("���l�ʒu", "�����񕪊�", r1.Offset(0, 1).Address, Type:=8)
        On Error GoTo 0
        If r2 Is Nothing Then Exit Sub
        If r2.Address = r1.Address Then Exit Sub
        v2 = ra.Column - r2.Column
        ra.Offset(, -v2).FormulaR1C1 = "=MID(RC[" & v2 & "],LEN(RC[" & (v2 - v1) & "])+1,LEN(RC[" & v2 & "]))"
    Case 2
        '�����񕪊�(���l�E�p���E���l)
        If True Then
            Set r0 = ra.Cells(1, 1)
            On Error Resume Next
            Set r1 = Application.InputBox("�擪���l�ʒu", "�����񕪊�", r0.Offset(0, 1).Address, Type:=8)
            On Error GoTo 0
            If r1 Is Nothing Then Exit Sub
            v1 = ra.Column - r1.Column
            ra.Offset(, -v1).FormulaR1C1 = "=IFERROR(VALUE(LEFT(RC[" & v1 & "],2)),IFERROR(VALUE(LEFT(RC[" & v1 & "],1)),""""))"
            
            On Error Resume Next
            Set r2 = Application.InputBox("�L���ʒu", "�����񕪊�", r1.Offset(0, 1).Address, Type:=8)
            On Error GoTo 0
            If r2 Is Nothing Then Exit Sub
            If r2.Address = r1.Address Then Exit Sub
            v2 = ra.Column - r2.Column
            ra.Offset(, -v2).Formula2R1C1 = "=LET(v,MID(RC[" & v2 & "],LEN(RC[" & (v2 - v1) & "])+1,LEN(RC[" & v2 & "])),LEFT(v,MIN(FIND({1,2,3,4,5,6,7,8,9,0},v&""1234567890""))-1))"
            
            On Error Resume Next
            Set r3 = Application.InputBox("���l�ʒu", "�����񕪊�", r2.Offset(0, 1).Address, Type:=8)
            On Error GoTo 0
            If r3 Is Nothing Then Exit Sub
            If r3.Address = r1.Address Then Exit Sub
            If r3.Address = r2.Address Then Exit Sub
            v3 = ra.Column - r3.Column
            ra.Offset(, -v3).FormulaR1C1 = "=MID(RC[" & v3 & "],LEN(RC[" & (v3 - v1) & "]&RC[" & (v3 - v2) & "])+1,LEN(RC[" & v3 & "]))"
        Else
            ra.Offset(, 1).FormulaR1C1 = "=IFERROR(VALUE(LEFT(RC[" & -1 & "],2)),IFERROR(VALUE(LEFT(RC[" & -1 & "],1)),""""))"
            ra.Offset(, 2).Formula2R1C1 = _
                "=LET(v,MID(RC[-2],LEN(RC[-1])+1,LEN(RC[-2])),LEFT(v,MIN(FIND({1,2,3,4,5,6,7,8,9,0},v&""1234567890""))-1))"
            ra.Offset(, 3).FormulaR1C1 = "=MID(RC[-3],LEN(RC[-2]&RC[-1])+1,LEN(RC[-3]))"
        End If
    Case 3
        '�����}�[�J�[
        Set r0 = ra.Cells(1, 1)
        On Error Resume Next
        Set r1 = Application.InputBox("��r���ʒu1", "�����}�[�J�\", r0.Offset(0, 1).Address, Type:=8)
        On Error GoTo 0
        If r1 Is Nothing Then Exit Sub
        On Error Resume Next
        Set r2 = Application.InputBox("��r���ʒu2", "�����}�[�J�\", r1.Offset(0, 1).Address, Type:=8)
        On Error GoTo 0
        If r2 Is Nothing Then Exit Sub
        If r2.Address = r1.Address Then Exit Sub
        v1 = r1.Column - ra.Column
        v2 = r2.Column - ra.Column
        '
        ra.Formula2R1C1 = _
            "=IF(OFFSET(RC,0," & v1 & ")=OFFSET(RC,0," & v2 & "),""�Z"","""")"
    End Select
End Sub

