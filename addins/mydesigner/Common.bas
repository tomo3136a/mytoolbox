Attribute VB_Name = "Common"
Option Explicit
Option Private Module

'----------------------------------------
'オブジェクト呼び出し
'----------------------------------------

'worksheet.function
Function wsf() As WorksheetFunction
    Set wsf = WorksheetFunction
End Function


'----------------------------------------
'正規表現
'----------------------------------------

'regex
Function regex( _
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

'文字列有無
Public Function re_test(s As String, ptn As String) As Boolean
    On Error Resume Next
    re_test = regex(ptn).test(s)
    On Error GoTo 0
End Function

'文字列抽出
Public Function re_match(s As String, ptn As String, _
        Optional idx As Integer = 0) As String
    Dim mc As Object
    Set mc = regex(ptn).Execute(s)
    If idx < 0 Or idx >= mc.Count Then Exit Function
    re_match = mc(idx).Value
End Function

'文字列置き換え
Public Function re_replace(s As String, ptn As String, rep As String) As String
    re_replace = regex(ptn).Replace(s, rep)
End Function


'----------------------------------------
'領域の値文字列取得
'----------------------------------------

Function StrRange(s As String) As String
    Dim ra As Range
    Set ra = Range(s)
    If ra.Count = 1 Then
        StrRange = s
        Exit Function
    End If
    Dim n As Integer
    n = ra.Column + ra.Columns.Count - 1
    Dim ce As Range
    Dim ss As String
    For Each ce In ra
        ss = ss & Chr(34) & ce.Value & Chr(34)
        If n = ce.Column Then
            ss = ss & vbLf
        Else
            ss = ss & ","
        End If
    Next ce
    StrRange = Left(ss, Len(ss) - 1)
End Function


'----------------------------------------
'コレクションを配列に変換
'----------------------------------------

Function ToArray(col As Collection) As Variant()
    Dim arr() As Variant
    ReDim arr(0 To col.Count - 1)
    Dim i As Integer
    For i = 1 To col.Count
        arr(i - 1) = col.Item(i)
    Next i
    ToArray = arr
End Function


'----------------------------------------
'領域の角取得
'----------------------------------------

Function LeftTop(ra As Range) As Range
    Set LeftTop = ra.Cells(1, 1)
End Function

Function RightTop(ra As Range) As Range
    Set RightTop = ra.Cells(1, ra.Columns.Count)
End Function

Function LeftBottom(ra As Range) As Range
    Set LeftBottom = ra.Cells(ra.Rows.Count, 1)
End Function

Function RightBottom(ra As Range) As Range
    Set RightBottom = ra.Cells(ra.Rows.Count, ra.Columns.Count)
End Function


'----------------------------------------
'セル探索
'----------------------------------------

'上端取得
Public Function FarTop(ra As Range, Optional margin As Integer) As Range
    If margin < 1 Then margin = 1
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Row
    Dim cnt As Integer
    Dim rs As Range
    Set rs = ce
    Do While ce.Row > p And cnt < margin
        If ce.Offset(-1).Value = "" Then
            Set ce = ce.Offset(-1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlUp)
            Set rs = ce
            cnt = 0
        End If
    Loop
    Set FarTop = rs
End Function

'下端取得
Public Function FarBottom(ra As Range, Optional margin As Integer) As Range
    If margin < 1 Then margin = 1
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Row
    p = p + ra.Worksheet.UsedRange.Rows.Count
    Dim cnt As Integer
    Dim re As Range
    Set re = ce
    Do While ce.Row < p And cnt < margin
        If ce.Offset(1).Value = "" Then
            Set ce = ce.Offset(1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlDown)
            Set re = ce
            cnt = 0
        End If
    Loop
    Set FarBottom = re
End Function

'左端取得
Public Function FarLeft(ra As Range, Optional margin As Integer) As Range
    If margin < 1 Then margin = 1
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Column
    Dim cnt As Integer
    Dim rs As Range
    Set rs = ce
    Do While ce.Column > p And cnt < margin
        If ce.Offset(0, -1).Value = "" Then
            Set ce = ce.Offset(0, -1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlToLeft)
            Set rs = ce
            cnt = 0
        End If
    Loop
    Set FarLeft = rs
End Function

'右端取得
Public Function FarRight(ra As Range, Optional margin As Integer) As Range
    If margin < 1 Then margin = 1
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim p As Long
    p = ra.Worksheet.UsedRange.Column
    p = p + ra.Worksheet.UsedRange.Columns.Count
    Dim cnt As Integer
    Dim re As Range
    Set re = ce
   Do While ce.Column < p And cnt < margin
        If ce.Offset(0, 1).Value = "" Then
            Set ce = ce.Offset(0, 1)
            cnt = cnt + 1
        Else
            Set ce = ce.End(xlToRight)
            Set re = ce
            cnt = 0
        End If
    Loop
    Set FarRight = re
End Function

'文字列が一致するcellを探す
Public Function FindCell(s As String, ra As Range) As Range
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Dim r As Long
    Dim c As Integer
    For r = ra.Row To ws.UsedRange.Rows.Count
        For c = ra.Column To ws.UsedRange.Columns.Count
            Dim ce As Range
            Set ce = ws.Cells(r, c)
            If ce.Value = s Then
                Set FindCell = ce
                Exit Function
            End If
        Next c
    Next r
End Function

'ブランクをスキップする
Public Function SkipBlankCell(ra As Range) As Range
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    Dim r As Long
    Dim c As Integer
    For r = ra.Row To ws.UsedRange.Rows.Count
        For c = ra.Column To ws.UsedRange.Columns.Count
            Dim ce As Range
            Set ce = ws.Cells(r, c)
            If ce.Value <> "" Then
                Set SkipBlankCell = ce
                Exit Function
            End If
        Next c
    Next r
End Function


'----------------------------------------
'パス操作
'----------------------------------------

'filesystemobject
Function fso() As Object
    Static obj As Object
    If obj Is Nothing Then
        Set obj = CreateObject("Scripting.FileSystemObject")
    End If
    Set fso = obj
End Function

'基本名取得
'  パス排除、拡張子排除
'  複製情報排除
Function BaseName(s As String) As String
    Dim re As Object
    Set re = regex("[\(（]\d+[\)）]|\s*-\s*コピー")
    BaseName = re.Replace(fso.GetBaseName(s), "")
End Function

Function CanonicalPath(path As String)
    Dim arr As Variant
    arr = Array("Box", "OneDrive", "LOCALAPPDATA", "APPDATA", "USERPROFILE")
    '
    Dim p As String
    p = path
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        Dim name As String
        name = arr(i)
        Dim Base As String
        Base = Environ(name)
        If Mid(p & "\", 1, Len(Base & "\")) = Base & "\" Then
            p = Replace(p, Base, "(" & name & ")", compare:=vbTextCompare)
            Exit For
        End If
    Next i
    '
    CanonicalPath = p
End Function

Function EnvironmentPath(path As String)
    EnvironmentPath = re_replace(path, "^\((\w+)\)", "%$1%")
End Function


'----------------------------------------
'シート名操作
'----------------------------------------

'シート名有無のチェック
Function HasSheetName(wb As Workbook, name As String) As Boolean
    Dim i As Integer
    For i = 1 To wb.Worksheets.Count
        If wb.Worksheets(i).name = name Then
            HasSheetName = True
            Exit Function
        End If
    Next i
End Function

'重複しないシート名取得
Function UniqueSheetName(wb As Workbook, name As String) As String
    Dim i As Integer: i = 1
    Dim s As String: s = name
    Do While HasSheetName(wb, s)
        s = name & " (" & i & ")"
        i = i + 1
    Loop
    UniqueSheetName = s
End Function

'シート名リネームダイアログ
Sub SheetRenameDialog()
    CommandBars.ExecuteMso "SheetRename"
End Sub


'----------------------------------------
'シートプロパティ操作
'----------------------------------------

'シートプロパティ名リストを取得
Function GetSheetPropertyNames(ws As Worksheet) As String()
    Dim lst() As String
    ReDim Preserve lst(ws.CustomProperties.Count)
    Dim i As Integer
    For i = 1 To ws.CustomProperties.Count
        lst(i) = ws.CustomProperties(i).name
    Next i
    GetSheetPropertyNames = lst
End Function

'シートプロパティ名から番号取得
Function SheetPropertyIndex(ws As Worksheet, name As String) As Integer
    Dim i As Integer
    For i = 1 To ws.CustomProperties.Count
        If ws.CustomProperties(i).name = name Then
            SheetPropertyIndex = i
            Exit Function
        End If
    Next i
End Function

'シートプロパティ名からプロパティ取得
Function GetSheetProperty(ws As Worksheet, name As String) As CustomProperty
    Dim idx As Integer
    idx = SheetPropertyIndex(ws, name)
    If idx > 0 Then
        Set GetSheetProperty = ws.CustomProperties(idx)
        Exit Function
    End If
    Set GetSheetProperty = ws.CustomProperties.Add(name, "")
End Function

'シートプロパティ名から値取得
Function StrSheetProperty(ws As Worksheet, name As String) As String
    Dim idx As Integer
    idx = SheetPropertyIndex(ws, name)
    If idx > 0 Then StrSheetProperty = ws.CustomProperties(idx).Value
End Function


'----------------------------------------
'画面チラつき防止
'----------------------------------------

Public Sub ScreenUpdateOff()
    '画面チラつき防止処置
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Interactive = False
    Application.Cursor = xlWait
End Sub

Public Sub ScreenUpdateOn()
    '画面チラつき防止処置解除
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


'----------------------------------------
'進行状況表示(status-bar)
'----------------------------------------

Sub ProgressStatusBar(Optional i As Long = 1, Optional cnt As Long = 1)
    Static tm_start As Double
    If i < 1 Then
        tm_start = Timer
        Application.StatusBar = "進捗状況(0%)"
        Exit Sub
    End If
    If i >= cnt Then
        Application.StatusBar = False
        Exit Sub
    End If
    Dim p As Double: p = i / cnt
    Dim s As String: s = "進捗状況(" & Int(p * 100) & "%)"
    s = s & " : " & ProgressBar(p)
    Dim tm As Double: tm = (Timer - tm_start) / p * (1 - p)
    Application.StatusBar = s & " : 残り" & Int(tm) & "秒"
End Sub

Private Function ProgressBar(p As Double) As String
    If p < 0.2 Then
        ProgressBar = "□□□□□"
    ElseIf p < 0.4 Then
        ProgressBar = "■□□□□"
    ElseIf p < 0.6 Then
        ProgressBar = "■■□□□"
    ElseIf p < 0.8 Then
        ProgressBar = "■■■□□"
    ElseIf p < 1 Then
        ProgressBar = "■■■■□"
    Else
        ProgressBar = "■■■■■"
    End If
End Function


