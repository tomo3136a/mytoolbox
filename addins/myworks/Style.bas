Attribute VB_Name = "Style"
Option Explicit
Option Private Module

'---------------------------------------------
'�J���[�}�[�J
'---------------------------------------------

'�Z���ɃJ���[�}�[�J��ݒ�
Public Sub AddMarker(ra As Range, id As Integer, Optional ByVal kw As String)
    
    If kw = "" Then kw = Replace(Mid(Date, 5), "/", "")
    If InStr(1, kw, "_") = 0 Then kw = kw & "_" & (id Mod 10)
    
    Dim wb As Workbook
    Set wb = ra.Parent.Parent
    
    On Error Resume Next
    If wb.Styles(kw) Is Nothing Then
        With wb.Styles.Add(kw)
            .IncludeNumber = False
            .IncludeFont = False
            .IncludeAlignment = False
            .IncludeBorder = False
            .IncludePatterns = True
            .IncludeProtection = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                Select Case id Mod 10
                Case 0: .Color = RGB(255, 241, 0)   '��
                Case 1: .Color = RGB(240, 125, 136) '��
                Case 2: .Color = RGB(85, 171, 229)  '��
                Case 3: .Color = RGB(95, 190, 125)  '����
                Case 4: .Color = RGB(185, 192, 203) '�D�F
                Case 5: .Color = RGB(255, 140, 0)   '��
                Case 6: .Color = RGB(51, 186, 177)  '��
                Case 7: .Color = RGB(163, 179, 103) '��
                Case 8: .Color = RGB(168, 149, 226) '��
                Case 9: .Color = RGB(2, 104, 2)     '��
                End Select
                .TintAndShade = 0
            End With
        End With
    End If
    On Error GoTo 0
    ra.Style = kw

End Sub

'�J���[�}�[�J�폜
Sub DelMarker(kw As String, Optional ByVal wb As Workbook)
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    Dim v As Variant
    For Each v In wb.Styles
        If CStr(v) Like kw Then wb.Styles(CStr(v)).Delete
    Next v

End Sub

'�J���[�}�[�J���X�g�擾
Sub ListMarker(ByVal ra As Range)
    
    Set ra = ra.Cells(1, 1)
    Dim wb As Workbook
    Set wb = ra.Parent.Parent
    
    Dim arr As Variant
    arr = re_extract(wb.Styles, "^\d{4}_\d{1,2}$")
    arr = wsf.Transpose(arr)
    If Not TypeName(arr) = "Variant()" Then Exit Sub
    ra.Resize(UBound(arr, 1), 1).Value = arr
    
    ScreenUpdateOff
    Dim v As Variant
    For Each v In arr
        ra.Style = v
        Set ra = ra.Offset(1)
    Next v
    ScreenUpdateOn
    
End Sub
