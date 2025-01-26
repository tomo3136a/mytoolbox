Attribute VB_Name = "Common"
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

'----------------------------------------
'正規表現
'----------------------------------------

'regex(VBScript.RegExp)
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
Function re_test(s As String, ptn As String) As Boolean
    On Error Resume Next
    re_test = regex(ptn).Test(s)
    On Error GoTo 0
End Function

'文字列抽出
Function re_match(s As String, ptn As String, _
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
Function re_replace(s As String, ptn As String, rep As String) As String
    On Error Resume Next
    re_replace = regex(ptn).Replace(s, rep)
    On Error GoTo 0
End Function

'文字列分割
Function re_split(s As String, ptn As String) As String()
    re_split = Split(regex(ptn).Replace(s, Chr(7)), Chr(7))
End Function

'配列からマッチした文字列を抽出
Function re_extract(col As Variant, ptn As String) As Variant
    Dim re As Object
    Set re = regex(ptn)
    
    Dim arr As Variant
    ReDim arr(50)
    
    Dim s As String
    Dim i As Integer
    Dim v As Variant
    For Each v In col
        s = v
        If re.Test(s) Then
            If i > UBound(arr) Then ReDim Preserve arr(UBound(arr) + 50)
            arr(i) = s
            i = i + 1
        End If
    Next v
    If i < 1 Then Exit Function
    ReDim Preserve arr(i - 1)
    re_extract = arr
End Function

'----------------------------------------
'検索
'----------------------------------------

'コレクションから名前を指定して検索(配列は除く)
Function SearchName(col As Object, name As String) As Object
    Dim v As Object
    For Each v In col
        If v.name = name Then
            Set SearchName = v
            Exit Function
        End If
    Next v
    Set SearchName = Nothing
End Function

'----------------------------------------
'パラメータ文字列
'  <text> = [ <line> \n ] <line>
'  <line> = \s* <key> \s* : \s* <val> \s* | .+
'  <key>  = \w+
'  <val>  = .+
'----------------------------------------

'パラメータ文字列からキーリスト取得
Function ParamStrKeys(s As String) As String()
    Dim lines() As String
    lines = Split(s, Chr(10), , vbTextCompare)
    Dim res() As String
    ReDim res(UBound(lines))
    Dim i As Integer, j As Integer
    Dim line As String
    For i = 0 To UBound(lines)
        line = lines(i)
        Dim kv() As String
        kv = Split(line, ":", 2, vbTextCompare)
        If UBound(kv) > 0 Then
            res(j) = Trim(kv(0))
            j = j + 1
        End If
    Next i
    ReDim Preserve res(j - 1)
    ParamStrKeys = res
End Function

'パラメータ文字列から値を取得
Function ParamStrVal(s As String, k As String) As String
    Dim line As Variant
    For Each line In Split(s, Chr(10), , vbTextCompare)
        Dim kv() As String
        kv = Split(line, ":", 2, vbTextCompare)
        If UBound(kv) > 0 Then
            If UCase(k) = UCase(Trim(kv(0))) Then
                ParamStrVal = Trim(kv(1))
                Exit Function
            End If
        End If
    Next line
End Function

'パラメータ文字列にキー・値を追加・更新
Function UpdateParamStr(s As String, k As String, v As String) As String
    Dim lines() As String
    lines = Split(s, Chr(10), , vbTextCompare)
    Dim line As String
    Dim i As Integer
    For i = 0 To UBound(lines)
        line = lines(i)
        Dim kv() As String
        kv = Split(line, ":", 2, vbTextCompare)
        If UBound(kv) > 0 Then
            If k = Trim(kv(0)) Then
                lines(i) = k & ":" & Trim(v)
                Exit For
            End If
        End If
    Next i
    line = Join(lines, Chr(10))
    If i > UBound(lines) Then
        line = Join(Array(line, k & ":" & Trim(v)), Chr(10))
    End If
    line = Replace(line, Chr(10) & Chr(10), Chr(10))
    UpdateParamStr = line
End Function

'パラメータ文字列から項目を削除
Function RemoveParamStr(s As String, k As String) As String
    Dim lines() As String
    lines = Split(s, Chr(10), , vbTextCompare)
    Dim res() As String
    ReDim res(0 To UBound(lines))
    Dim i As Integer, j As Integer
    Dim line As String
    For i = 0 To UBound(lines)
        line = lines(i)
        Dim kv() As String
        kv = Split(line, ":", 2, vbTextCompare)
        If UBound(kv) > 0 Then
            If k <> Trim(kv(0)) Then
                res(j) = line
                j = j + 1
            End If
        Else
            res(j) = line
            j = j + 1
        End If
    Next i
    ReDim Preserve res(j)
    RemoveParamStr = Join(res, Chr(10))
End Function

'パラメータ文字列からディクショナリ作成
Sub ParamStrDict(dict As Dictionary, s As String)
    If dict Is Nothing Then Set dict = New Dictionary
    Dim line As Variant
    For Each line In Split(s, Chr(10), , vbTextCompare)
        Dim kv() As String
        kv = Split(line, ":", 2, vbTextCompare)
        If UBound(kv) > 0 Then
            dict.Add Trim(kv(0)), Trim(kv(1))
        End If
    Next line
End Sub

'----------------------------------------
'配列文字列
'  <text> = [ <rows> ; ] <rows>
'  <rows> = [ \s* <item> \s+ , ] \s* <item> \s*
'  <item> = \w+
'----------------------------------------

'配列文字列から配列へ変換
Function StrToArr(s As String) As Variant
    Dim lines() As String
    lines = Split(s, ";", , vbTextCompare)
    If UBound(lines) < 2 Then
        StrToArr = Split(s, ",", , vbTextCompare)
        Exit Function
    End If
    
    Dim r As Long, c As Long
    Dim ss() As String
    Dim v As Variant
    For Each v In lines
        ss = Split(v, ",", , vbTextCompare)
        If UBound(ss) >= 0 Then
            If Trim(ss(0)) <> "" Then r = r + 1
            If c < UBound(ss) Then c = UBound(ss)
        End If
    Next v
    Dim res As Variant
    ReDim res(1 To r, 1 To c + 1)
    
    Dim i As Long, j As Long
    For Each v In lines
        ss = Split(v, ",", , vbTextCompare)
        If UBound(ss) >= 0 Then
            If Trim(ss(0)) <> "" Then i = i + 1
            For j = 0 To UBound(ss)
                res(i, j + 1) = Trim(ss(j))
            Next j
        End If
    Next v
    
    StrToArr = res
End Function

'配列から配列文字列へ変換
Function ArrToStr(arr As Variant) As String
    Dim s() As String
    ReDim s(0 To UBound(s, 0))
    Dim i As Long, j As Long
    For i = 0 To UBound(arr, 0)
        s(i) = arr(i, 0)
        For j = 1 To UBound(arr, 1)
            s(i) = s(i) & "," & arr(i, j)
        Next j
    Next i
    ArrToStr = Join(s, ";")
End Function

'----------------------------------------
'データ変換
'----------------------------------------

'コレクションを配列に変換
Function ColToArr(col As Collection) As Variant()
    Dim arr() As Variant
    ReDim arr(0 To col.Count - 1)
    Dim i As Integer
    For i = 1 To col.Count
        arr(i - 1) = col.Item(i)
    Next i
    ColToArr = arr
End Function

'二次配列文字列から配列辞書に変換
Sub ArrToDict(dic As Dictionary, arr As Variant, Optional n As Integer)
    
    If dic Is Nothing Then Set dic = New Dictionary
    If n > UBound(arr, 2) - LBound(arr, 2) Then Exit Sub
    
    Dim i As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Not dic.Exists(arr(i, n + LBound(arr, 2))) Then
            dic.Add arr(i, n + LBound(arr, 2)), arr = wsf.Index(arr, i, Array(2, 4))
        End If
    Next i
    
End Sub

'二次配列文字列から配列辞書に変換
Sub ArrStrToDict(dic As Dictionary, s As String, Optional n As Integer)
    
    If dic Is Nothing Then Set dic = New Dictionary
    Dim s1 As String
    s1 = Replace(s, " ", "")
    
    Dim va As Variant
    For Each va In Split(s1, ";")
        Dim ss() As String
        ss = Split(va, ",")
        If UBound(ss) > n Then
            Dim i As Integer
            For i = 0 To n
                Dim k As String
                k = UCase(ss(i))
                If k <> "" Then
                    If Not dic.Exists(k) Then
                        dic.Add k, ss
                    End If
                End If
            Next i
        End If
    Next va
    
End Sub

'配列の範囲抽出
Function TakeArray(arr() As String, Optional p As Integer, Optional n As Integer) As String()
    
    Dim sz As Integer
    sz = n
    If sz = 0 Then sz = UBound(arr) - LBound(arr) - p + 1
    Dim sp As Integer
    sp = p + LBound(arr)
    
    Dim sa() As String
    ReDim sa(0 To sz - 1)
    
    Dim i As Long
    For i = 0 To sz - 1
        sa(i) = arr(sp + i)
    Next i
    TakeArray = sa

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
Function GetAbstructPath(path As String, Base As String) As String
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
    If InStr(1, p, ":\") = 0 Then p = Base & p
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
Function GetRelatedPath(path As String, Base As String) As String
    Dim sep As String, s As String
    If Right(path, 1) = "\" Then sep = "\"
    Dim ss1 As Variant, ss2 As Variant
    ss1 = Split(GetAbstructPath(path, Base), "\")
    ss2 = Split(Base, "\")
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

'----------------------------------------
'実行時パラメータ機能
'----------------------------------------

'パラメータ設定
Sub SetRtParam(grp As String, k As String, Optional v As String)
    Dim dic As Dictionary
    Set dic = rt_param_dict
    Dim kw As String
    kw = grp & "_" & k
    If dic.Exists(kw) Then dic.Remove kw
    If v <> "" Then dic.Add kw, v
End Sub

'パラメータ取得
Function GetRtParam(grp As String, k As String, Optional v As String) As String
    Dim dic As Dictionary
    Set dic = rt_param_dict
    Dim kw As String
    kw = grp & "_" & k
    GetRtParam = v
    If dic.Exists(kw) Then GetRtParam = dic.Item(kw)
End Function

'パラメータ取得(boolean)
Function GetRtParamBool(grp As String, k As String) As Boolean
    Dim s As String
    s = GetRtParam(grp, k)
    If s = "" Then s = "False"
    GetRtParamBool = s
End Function

'パラメータディクショナリ
Private Function rt_param_dict() As Dictionary
    Static dic As Dictionary
    If dic Is Nothing Then Set dic = New Dictionary
    Set rt_param_dict = dic
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
Function StrSheetProperty(ws As Worksheet, name As String) As String
    Dim idx As Integer
    idx = SheetPropertyIndex(ws, name)
    If idx > 0 Then StrSheetProperty = ws.CustomProperties(idx).Value
End Function

'----------------------------------------
'画面チラつき防止
'----------------------------------------

'画面チラつき防止処置
Public Sub ScreenUpdateOff()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Interactive = False
    Application.Cursor = xlWait
End Sub

'画面チラつき防止処置解除
Public Sub ScreenUpdateOn()
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
    Dim TM As Double: TM = (Timer - tm_start) / p * (1 - p)
    Application.StatusBar = s & " : 残り" & Int(TM) & "秒"
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

'----------------------------------------
'アドインブック
'----------------------------------------

'アドインブック表示トグル
Sub ToggleAddinBook()
    If ThisWorkbook.IsAddin Then
        ThisWorkbook.IsAddin = False
        ThisWorkbook.Activate
    Else
        ThisWorkbook.IsAddin = True
        ThisWorkbook.Save
    End If
End Sub

'アドインブックからテンプレートシートを複製
Sub CopyAddinSheet()
    Dim ws As Worksheet
    Set ws = SelectSheet(ThisWorkbook, "^[^#]")
    If ws Is Nothing Then Exit Sub
    ws.Copy After:=ActiveSheet
End Sub

'アドインブックのテンプレートシート更新
Function UpdateAddinSheet(ws As Worksheet)
    Dim asu As Boolean
    asu = Application.ScreenUpdating
    Application.ScreenUpdating = False
    '
    Dim ws2 As Worksheet
    For Each ws2 In ThisWorkbook.Sheets
        If ws2.name = ws.name Then Exit For
    Next ws2
    If ws2 Is Nothing Then
        ThisWorkbook.IsAddin = False
        ws.Copy After:=ThisWorkbook.Sheets(1)
        ThisWorkbook.IsAddin = True
    Else
        Dim old As Range
        Set old = Selection
        ws.Cells.Select
        Selection.Copy
        Set ws2 = ThisWorkbook.Sheets(ws.name)
        ws2.Paste ws2.Cells(1, 1)
        Application.CutCopyMode = False
        old.Select
    End If
    '
    Application.ScreenUpdating = asu
End Function

