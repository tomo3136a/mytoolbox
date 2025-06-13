Attribute VB_Name = "Config"
'==================================
'設定
'==================================

Option Explicit
Option Private Module

'設定シート取得
Function ConfigSheet(Optional name As String = "#config") As Worksheet
    Set ConfigSheet = ActiveWorkbook.Sheets(name)
    If Not ConfigSheet Is Nothing Then Exit Function
    Set ConfigSheet = ThisWorkbook.Sheets(name)
End Function

'セクション一覧を取得
Function SectionRange(ra As Range) As Range
    Dim arr As Range
    Dim ce As Range
    Dim cnt As Long
    For Each ce In ra
        Dim s As String
        Dim ce2 As Range
        For Each ce2 In ce.Cells
            s = ce2.Value
            If Left(ce2.Value, 1) = "[" And Right(ce2.Value, 1) = "]" Then
                If cnt = 0 Then Set arr = ce2 Else Set arr = Union(arr, ce2)
                cnt = cnt + 1
            End If
        Next ce2
    Next ce
    Set SectionRange = arr
End Function

'設定取得
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
'共通機能
'----------------------------------

'アドイン名を取得
Public Function AddinName() As String
    AddinName = Replace(ThisWorkbook.name, ".xlam", "")
End Function

'アドインのパス取得
Public Function AddinsPath() As String
    AddinsPath = ThisWorkbook.path
End Function

'アドイン設定シートの領域取得
Function AddinsListRange(Optional name As String = "#config") As Range
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets.Item(name)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Set AddinsListRange = ws.UsedRange
End Function

'----------------------------------
'テーブル読み込み機能
'----------------------------------

Function GetTextArray(ra As Range) As Variant
    Dim dic() As Variant
    Dim ce As Range
    Dim cnt As Long
    For Each ce In ra
        If ce <> "" Then
            ReDim Preserve dic(cnt)
            dic(cnt) = ce.Text
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


