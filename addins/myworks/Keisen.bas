Attribute VB_Name = "Keisen"
'==================================
'�r���g����
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'�@�\�Ăяo��
'----------------------------------------

'�e�[�u���I��
' mode=0: �e�[�u���I��
'      1: �擪�ֈړ�
'      2: �����ֈړ�
'      3: �s�I��
'      4: ��I��
'      5: �w�b�_�s�I��
Sub TableSelect(ra As Range, Optional mode As Integer)
    Dim rb As Range
    Set rb = ra.CurrentRegion
    Select Case mode
    Case 0: rb.Select
    Case 1: rb(1, 1).Select
    Case 2: rb(rb.Rows.Count + 1, 1).Select
    Case 3: Intersect(rb, ra.EntireRow).Select
    Case 4: Intersect(rb, ra.EntireColumn).Select
    Case 5: rb.Rows(1).Select
    End Select
End Sub

'�r���g
' mode=0: �r���g(�W���ݒ�)
'      1: �r���g(�W��)
'      2: �r���g(�K�w�\��)
'      3: �w�b�_�t�B���^
'      4: �w�b�_�����킹
'      5: �w�b�_�Œ�
'      6: �w�b�_�F
'      7: �g�N���A
'      8: �l�N���A
'      9: �e�[�u���N���A
Sub TableWaku(ra As Range, Optional mode As Integer)
    Select Case mode
    Case 0: Waku ra, fit:=True
    Case 1: Waku ra
    Case 2: WakuLayered ra
    Case 3: HeaderFilter ra
    Case 4: HeaderAutoFit ra
    Case 5: HeaderFixed ra
    Case 6: HeaderColor ra
    Case 7: WakuClear ra: ra.FormatConditions.Delete
    Case 8: TableRange(TableHeaderRange(TableLeftTop(ra)).Offset(1)).Clear
    Case 9: TableRange(TableHeaderRange(TableLeftTop(ra))).Clear
    End Select
End Sub

'��ǉ�
' mode=1: �ԍ���ǉ�
Sub AddColumn(ra As Range, mode As Integer)
    Dim rb As Range
    Set rb = Intersect(ra.CurrentRegion, ra.EntireColumn)
    If ra.Rows.Count > 1 Then Set rb = ra
    Set rb = rb.Columns(1)
    rb.EntireColumn.Insert shift:=xlShiftToRight
    
    Dim rc As Range
    Set rc = rb.Offset(0, -1)
    rb.Copy
    rc.PasteSpecial Paste:=xlPasteFormats, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = 0
    
    Dim bhdr As Boolean
    If ra.CurrentRegion.Row = rb.Row Then bhdr = True
    
    Select Case mode
    Case 1: AddNoColumn rc, bhdr
    Case 2: AddRankColumn rc, bhdr
    End Select
    
    ra.Select
End Sub


'----------------------------------------
'API
'----------------------------------------

'�͂�
Sub Waku(ByVal ra As Range, _
        Optional filter As Boolean, _
        Optional fit As Boolean, _
        Optional fixed As Boolean, _
        Optional icolor As Integer = 15 _
    )
    If ra.Count = 1 Then Set ra = ra.CurrentRegion
    ra.Borders.LineStyle = xlContinuous
    '
    Dim rh As Range
    Set rh = ra.Rows(1)
    If GetHeaderColor = 0 Then
        rh.Interior.ColorIndex = icolor
    Else
        rh.Interior.Color = GetHeaderColor
    End If
    If filter Then HeaderFilter rh
    '
    If ra.Rows.Count > 1 Then Set ra = ra.Resize(ra.Rows.Count - 1).Offset(1)
    If fit Then ra.Columns.AutoFit
End Sub

'�͂�(�K�w�\��)
Private Sub WakuLayered(ByVal ra As Range)
    If ra.Rows.Count = 1 And ra.Count > 1 Then
        Set ra = Intersect(ra.CurrentRegion, ra.EntireColumn)
    End If
    If ra.Count = 1 Then Set ra = ra.CurrentRegion
    
    Waku ra, fit:=True
    ra.FormatConditions.Delete
    Dim s As String
    s = ra(1, 1).Address(False, False)
    s = "=LET(a," & s & ",b,OFFSET(a,-1,0),OR(""""&a="""",AND(""""&a=""""&b,SUBTOTAL(3,b)>0)))"
    ra.FormatConditions.Add Type:=xlExpression, Formula1:=s
    ra.FormatConditions(ra.FormatConditions.Count).SetFirstPriority
    ra.FormatConditions(1).NumberFormat = ";;;"
    ra.FormatConditions(1).Borders(xlTop).LineStyle = xlNone
    ra.FormatConditions(1).StopIfTrue = False
End Sub

'�͂��N���A
Private Sub WakuClear(ByVal ra As Range)
    If ra.Rows.Count = 1 And ra.Count > 1 Then
        Set ra = Intersect(ra.CurrentRegion, ra.EntireColumn)
    End If
    If ra.Count = 1 Then Set ra = ra.CurrentRegion
    
    ra.FormatConditions.Delete
    ra.Interior.ColorIndex = xlColorIndexNone
    ra.Borders.LineStyle = xlNone
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
    ActiveWindow.FreezePanes = False
End Sub

'----------------------------------------
'�ԍ���ǉ�
'----------------------------------------

'�ԍ���ǉ�
Sub AddNoColumn(ra As Range, bhdr As Boolean, Optional shdr As String = "No.")
    Dim arr() As Variant
    arr = ra.Value
    'ReDim arr(0 To ra.Rows.Count, 1 To 1)
    
    arr(0, 1) = shdr
    Dim i As Long, j As Long
    If bhdr And shdr <> "" Then j = 1
    For i = 1 To ra.Rows.Count
        arr(j, 1) = i
        j = j + 1
    Next i
    ra.Value = arr
    ra.EntireColumn.Columns.AutoFit
End Sub

'�����N��ǉ�
Sub AddRankColumn(ra As Range, bhdr As Boolean, Optional shdr As String = "No.")
    Dim arr() As Variant
    ReDim arr(0 To ra.Rows.Count, 1 To 1)
    
    arr(0, 1) = shdr
    Dim i As Long, j As Long
    If shdr <> "" Then j = 1
    For i = 1 To ra.Rows.Count
        arr(j, 1) = i
        j = j + 1
    Next i
    ra.Value = arr
    ra.EntireColumn.Columns.AutoFit
End Sub




Private Function IsLineNo(ra As Range) As Boolean
    Dim s As String
    s = ra.Cells(1, 1).Value
    If s = "No." Or s = "#" Or s = "�ԍ�" Then IsLineNo = True
End Function



