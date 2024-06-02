Attribute VB_Name = "Base"
'==================================
'共通
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'オブジェクト呼び出し
'----------------------------------------

'worksheet.function
Function wsf() As WorksheetFunction
    Set wsf = WorksheetFunction
End Function

'regex
Function regex(ptn As String, _
        Optional g As Boolean = True, _
        Optional ic As Boolean = True) As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = g
        .IgnoreCase = ic
        .Pattern = ptn
    End With
End Function

'filesystemobject
Function fso() As Object
    Static obj As Object
    If obj Is Nothing Then
        Set obj = CreateObject("Scripting.FileSystemObject")
    End If
    Set fso = obj
End Function

'----------------------------------------
'正規表現
'----------------------------------------

'文字列判定
Public Function re_test(s As String, ptn As String) As Boolean
    On Error Resume Next
    re_test = regex(ptn).test(s)
    On Error GoTo 0
End Function

'文字列抽出
Public Function re_match(s As String, ptn As String, _
        Optional idx As Integer = 0, _
        Optional idx2 As Integer = -1) As Variant
    On Error Resume Next
    Dim re As Object
    Set re = regex(ptn)
    Dim mc As Object
    Set mc = re.Execute(s)
    
    If idx >= mc.Count Then
        re_match = ""
    ElseIf idx < 0 Then
        re_match = mc.Count
    ElseIf idx2 < 0 Then
        re_match = mc(idx).Value
    ElseIf idx2 < mc(idx).SubMatches.Count Then
        re_match = mc(idx).SubMatches(idx2)
    Else
        re_match = ""
    End If
    On Error GoTo 0
End Function

'文字列置き換え
Public Function re_replace(s As String, ptn As String, rep As String) As String
    On Error Resume Next
    re_replace = regex(ptn).Replace(s, rep)
    On Error GoTo 0
End Function

'----------------------------------
'パラメータ機能
'----------------------------------

Sub SetParam(grp As String, k As String, ByVal v As String)
    Static dic As Dictionary
    Set dic = param_dict
    Dim kw As String
    kw = grp & "_" & k
    On Error Resume Next
    dic.Remove kw
    dic.Add kw, v
    On Error GoTo 0
End Sub

Function GetParam(grp As String, k As String) As String
    Static dic As Dictionary
    Set dic = param_dict
    Dim kw As String
    kw = grp & "_" & k
    On Error Resume Next
    GetParam = dic.Item(kw)
    On Error GoTo 0
End Function

Function GetParamBool(grp As String, k As String) As Boolean
    Dim s As String
    s = GetParam(grp, k)
    If s = "" Then s = "False"
    GetParamBool = s
End Function

Private Function param_dict() As Dictionary
    Static dic As Dictionary
    If dic Is Nothing Then Set dic = New Dictionary
    Set param_dict = dic
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

'シートプロパティ数を取得
Function SheetPropertyCount(ws As Worksheet) As Integer
    SheetPropertyCount = ws.CustomProperties.Count
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

'シートプロパティ名からプロパティ値取得
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
Function GetSheetPropertyValue(ws As Worksheet, name As String) As String
    Dim idx As Integer
    idx = SheetPropertyIndex(ws, name)
    If idx > 0 Then GetSheetPropertyValue = ws.CustomProperties(idx).Value
End Function

'----------------------------------------
'データ変換
'----------------------------------------

'コレクションを配列に変換
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
'領域の値文字列取得
'----------------------------------------

Function StrRange(s As String) As String
    If Range(s).Count = 1 Then
        StrRange = s
        Exit Function
    End If
    Dim n As Integer
    n = Range(s).Column + Range(s).Columns.Count - 1
    Dim ra As Range
    Dim ss As String
    For Each ra In Range(s)
        ss = ss & Chr(34) & ra.Value & Chr(34)
        If n = ra.Column Then
            ss = ss & vbLf
        Else
            ss = ss & ","
        End If
    Next ra
    StrRange = Left(ss, Len(ss) - 1)
End Function

'----------------------------------------
'セル操作
'----------------------------------------

'領域の角取得
Function LeftTop(ra As Range) As Range
    Set LeftTop = ra.Parent.Cells(ra.Row, ra.Column)
End Function

Function RightTop(ra As Range) As Range
    Set RightTop = ra.Parent.Cells(ra.Row, ra.Column + ra.Columns.Count - 1)
End Function

Function LeftBottom(ra As Range) As Range
    Set LeftBottom = ra.Parent.Cells(ra.Row + ra.Rows.Count - 1, ra.Column)
End Function

Function RightBottom(ra As Range) As Range
    Set RightBottom = ra.Parent.Cells(ra.Row + ra.Rows.Count - 1, ra.Column + ra.Columns.Count - 1)
End Function

'文字列が一致するcellを探す
Function FindCell(ra As Range, s As String) As Range
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
Function SkipBlankCell(ra As Range) As Range
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

'基本名取得
'  パス削除、拡張子削除、複製情報削除
Function BaseName(s As String) As String
    Dim re As Object
    Set re = regex("[\(（]\d+[\)）]|\s*-\s*コピー")
    BaseName = re.Replace(fso.GetBaseName(s), "")
End Function

'短縮パス取得
Function GetShortPath(path As String, Optional pc As Boolean) As String
    Dim col As Collection
    Set col = GetEnvPathName()
    '
    Dim p As String
    p = Replace(path, "/", "\")
    If Right(p, 1) <> "\" Then p = p & "\"
    '
    Dim s As String
    Dim name As Variant
    For Each name In col
        s = Environ(name)
        s = Replace(s, "/", "\")
        If Right(s, 1) <> "\" Then s = s & "\"
        '
        If UCase(Mid(p, 1, Len(s))) = UCase(s) Then
            If pc Then
                p = "%" & name & "%" & Mid(path, Len(s))
            Else
                p = "(" & name & ")" & Mid(path, Len(s))
            End If
            GetShortPath = p
            Exit Function
        End If
    Next name
    '
    GetShortPath = path
End Function

Private Function GetEnvPathName(Optional reset As Boolean) As Collection
    Static col As Collection
    If reset Then
        Set col = Nothing
        Exit Function
    ElseIf Not col Is Nothing Then
        Set GetEnvPathName = col
        Exit Function
    End If
    '
    Dim arr As Variant
    arr = Array("Box", "OneDrive", _
        "TMP", "TEMP", "LOCALAPPDATA", "APPDATA", "PUBLIC", _
        "USERPROFILE", "HOME", _
        "ProgramData", "SystemRoot", _
        "CommonProgramFiles", "CommonProgramFiles(x86)", _
        "ProgramFiles", "ProgramFiles(x86)")
    '
    Dim dic As Dictionary
    Set dic = New Dictionary
    '
    Dim ss As Variant
    Dim s As String
    Dim i As Integer
    Do
        i = i + 1
        s = Environ(i)
        If s = "" Then Exit Do
        ss = Split(s, "=", 2)
        If InStr(1, ss(1), "\") Then
            If Not dic.Exists(ss(0)) Then dic.Add ss(0), ss(1)
        End If
    Loop
    '
    Set col = New Collection
    Dim v As Variant
    For Each v In arr
        s = CStr(v)
        If dic.Exists(s) Then
            col.Add s
            dic.Remove s
        End If
    Next v
    '
    For Each v In dic.Keys
        col.Add CStr(v)
    Next v
    Set dic = Nothing
    '
    Set GetEnvPathName = col
End Function

'絶対パス取得
Function GetAbstructPath(path As String, base As String) As String
    Dim p As String
    Dim s As String, s2 As String
    p = path
    s = re_match(p, "^[\(%](\w+)[\)%]", 0, 0)
    If s <> "" Then
        s2 = Environ(s)
        If s2 <> "" Then p = s2 & Mid(p, Len(s) + 3)
    End If
    p = Replace(p, "/", "\")
    p = Replace(p, "\\", "\")
    If InStr(1, p, ":\") = 0 Then p = base & p
    Do
        s = p
        p = re_replace(p, "\\[^\\]+\\[.][.]\\", "\")
        If s = p Then Exit Do
    Loop
    Do
        s = p
        p = re_replace(p, "\\[.]\\", "\")
        If s = p Then Exit Do
    Loop
    GetAbstructPath = p
End Function

'相対パス取得
Function GetRelatedPath(path As String, base As String) As String
    Dim sep As String, s As String
    If Right(path, 1) = "\" Then sep = "\"
    Dim ss1 As Variant, ss2 As Variant
    ss1 = Split(GetAbstructPath(path, base), "\")
    ss2 = Split(base, "\")
    '
    Dim i As Integer, j As Integer
    Dim v As Variant
    For Each v In ss2
        If UBound(ss1) <= i Then Exit For
        If v <> ss1(i) Then Exit For
        i = i + 1
    Next v
    For j = i To UBound(ss2)
        If ss2(j) <> "" Then s = fso.BuildPath(s, "..")
    Next j
    For j = i To UBound(ss1)
        s = fso.BuildPath(s, ss1(j))
    Next j
    s = s & sep
    GetRelatedPath = s
End Function

