Attribute VB_Name = "Style"
Option Explicit
Option Private Module

'---------------------------------------------
'カラーマーカ
'---------------------------------------------

'セルにカラーマーカを設定
Public Sub AddMarker(ra As Range, id As Integer, Optional name As String)
    If name = "" Then name = Replace(Mid(Date, 5), "/", "")
    If InStr(1, name, "_") = 0 Then name = name & "_" & id
    '
    Dim wb As Workbook
    Set wb = ra.Parent.Parent
    '
    On Error Resume Next
    If wb.Styles(name) Is Nothing Then
        With wb.Styles.Add(name)
            .IncludeNumber = False
            .IncludeFont = False
            .IncludeAlignment = False
            .IncludeBorder = False
            .IncludePatterns = True
            .IncludeProtection = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                Select Case id
                Case 1  '赤
                    .ColorIndex = 22
                Case 2  '青
                    .ColorIndex = 33
                Case 3  '黄緑
                    .ColorIndex = 43
                Case 4  '灰色
                    .ColorIndex = 15
                Case 5  '橙
                    .ColorIndex = 45
                Case 6  '青緑
                    .ColorIndex = 42
                Case 7  '茶
                    .ColorIndex = 40
                Case 8  '紫
                    .ColorIndex = 39
                Case 9  '緑
                    .ColorIndex = 10
                Case 10 '黄
                    .ColorIndex = 6
                Case Else
                End Select
                .TintAndShade = 0
            End With
        End With
    End If
    On Error GoTo 0
    '
    ra.Style = name
End Sub

'カラーマーカ削除
Sub DelMarker(s As String)
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    '
    On Error Resume Next
    Dim v As Variant
    For Each v In wb.Styles
        
        Dim name As String
        name = v
        
        If name = s Then
            wb.Styles(name).Delete
        End If
    Next v
    On Error GoTo 0
End Sub

'カラーマーカ全削除
Sub DelMarkerAll()
    Dim re As Object
    Set re = regex("^\d{4}_\d{1,2}$")
    '
    Dim res As Integer
    res = MsgBox("全て削除しますか？(" & ActiveWorkbook.Styles.Count & ")", vbYesNo, "マーカ削除")
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    On Error Resume Next
    Dim v As Variant
    For Each v In wb.Styles
        Dim name As String
        name = v
        If re.Test(name) Then
            wb.Styles(name).Delete
        End If
    Next v
    On Error GoTo 0
End Sub

'カラーマーカリスト取得
Sub ListMarker(ra As Range)
    Dim re As Object
    Set re = regex("^\d{4}_\d{1,2}$")
    '
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    '
    ScreenUpdateOff
    On Error Resume Next
    Dim v As Variant
    For Each v In ActiveWorkbook.Styles
        Dim name As String
        name = v
        If re.Test(name) Then
            ce.Value = name
            ce.Style = v
            Set ce = ce.Offset(1)
        End If
    Next v
    On Error GoTo 0
    ScreenUpdateOn
End Sub

Sub PickupFillColor()
    Dim ra As Range
    Set ra = Selection
    Dim ce As Range
    On Error Resume Next
    Set ce = Application.InputBox("対象のセル", Type:=8)
    On Error GoTo 0
    ra.Value = ce.Interior.color
End Sub

