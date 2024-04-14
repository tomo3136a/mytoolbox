Attribute VB_Name = "Config"
'==================================
'�ݒ�
'==================================

Option Explicit
Option Private Module

'�ݒ�V�[�g�擾
Function ConfigSheet(Optional name As String = "#config") As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets(name)
    If ws Is Nothing Then
        Set wb = ThisWorkbook
        Set ws = wb.Sheets(name)
    End If
    Set ConfigSheet = ws
End Function

'�Z�N�V�����ꗗ���擾
Function SectionRange(ra As Range) As Range
    Dim arr As Range
    Dim ce As Range
    Dim cnt As Long
    For Each ce In ra
        Dim s As String
        s = ce
        If Left(ce.Value, 1) = "[" And Right(ce.Value, 1) = "]" Then
            If cnt = 0 Then Set arr = ce Else Set arr = Union(arr, ce)
            cnt = cnt + 1
        End If
    Next ce
    Set SectionRange = arr
End Function

'�ݒ�擾
Function LoadConfig( _
        Optional sec As String, _
        Optional sht As String = "#config", _
        Optional wb As Workbook = Null) As Range
    '
    If wb Is Nothing Then Set wb = ThisWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sht)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = SelectSheet(wb)
    If ws Is Nothing Then Exit Function
    '
    Dim ra As Range
    If sec = "" Then
        Set ra = SectionRange(ws.UsedRange)
        Set ra = SelectCell(ra)
    Else
        For Each ra In SectionRange(ws.UsedRange)
            If ra.Value = sec Then Exit For
        Next ra
    End If
    If ra Is Nothing Then Exit Function
    Set ra = ra.Offset(1)
    Set LoadConfig = ra
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function SelectDictionary(dic As Object) As String
    If dic Is Nothing Then Exit Function
    SelectForm.reset
    Dim v As Variant
    For Each v In dic.Keys
        SelectForm.AddItem ("" & v)
    Next v
    SelectForm.Show
    SelectDictionary = dic(SelectForm.Result)
    Unload SelectForm
End Function

Private Sub test_dic(ra As Range)
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    For i = 0 To 100
        Dim k As String
        k = Trim(ra.Value)
        If k = "" Then Exit For
        dic.Add k, ra.Offset(0, 1).Value
        Set ra = ra.Offset(1)
    Next i
    Dim s As String
    s = SelectDictionary(dic)
End Sub

'----------------------------------
'���ʋ@�\
'----------------------------------

'�A�h�C�������擾
Public Function AddinName() As String
    AddinName = Replace(ThisWorkbook.name, ".xlam", "")
End Function

'�A�h�C���̃p�X�擾
Public Function AddinsPath() As String
    AddinsPath = ThisWorkbook.path
End Function

'�A�h�C���ݒ�V�[�g�̗̈�擾
Function AddinsListRange(Optional name As String = "#config") As Range
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets.Item(name)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Set AddinsListRange = ws.UsedRange
End Function

'----------------------------------
'�e�[�u���ǂݍ��݋@�\
'----------------------------------

Function GetTextArray(ra As Range) As Variant
    Dim dic() As Variant
    Dim ce As Range
    Dim cnt As Long
    For Each ce In ra
        If ce <> "" Then
            ReDim Preserve dic(cnt)
            dic(cnt) = ce.text
            cnt = cnt + 1
        End If
    Next ce
    cnt = 0
    GetTextArray = dic
End Function

Function GetValueArray(ra As Range) As Variant
    Dim dic() As Variant
    Dim ce As Range
    Dim cnt As Long
    For Each ce In ra
        If ce <> "" Then
            ReDim Preserve dic(cnt)
            dic(cnt) = ce.Value
            cnt = cnt + 1
        End If
    Next ce
    cnt = 0
    GetValueArray = dic
End Function


