Attribute VB_Name = "Base"
'==================================
'共通
'==================================

'----------------------------------------
'APi:
'  オブジェクト取得
'  wsf()                        WorksheetFunction
'  fso()                        FileSystemObject
'  regex(ptn,[g,ic])            VBScript.RegExp
'
'  コレクション操作
'  TakeByName(col,s)            コレクションから名前を指定して検索
'  TakeByValue(col,s)           コレクションから値を指定して検索
'
'  名前変換
'  ColumnName(idx)              列名取得
'  KeywordName(s,[sep])         キーワード名取得
'
'  パラメータ文字列
'  ParamStrKeys(s)              パラメータ文字列からキーリスト取得
'  ParamStrVal(s,k)             パラメータ文字列から値を取得
'  UpdateParamStr(s,k,v)        パラメータ文字列にキー・値を追加・更新
'  RemoveParamStr(s,k)          パラメータ文字列から項目を削除
'  RemoveParamStrAll(s)         パラメータ文字列からパラメータ以外取得
'  ParamStrDict(dict,s)         パラメータ文字列からディクショナリ作成
'
'  配列文字列
'  StrToArr(s)                  配列文字列から配列へ変換
'  ArrToStr(arr)                配列から配列文字列へ変換
'
'  データ変換
'  ColToArr(col)                コレクションを配列に変換
'  ArrToDict(dic,arr,[n])       二次配列文字列から配列辞書に変換
'  ArrStrToDict(dic,s,[n])      二次配列文字列から配列辞書に変換
'  TakeArray(arr(),[p,n])       配列の範囲抽出
'
'  領域の値文字列取得
'  StrRange(s)                  領域の値文字列取得
'
'  パス操作
'  CoreName(s)                  基本名取得(パス削除、拡張子削除、複製情報削除)
'  UniqueFileName(s)            重複しないファイル名取得
'  GetShortPath(path,[pc])      短縮パス取得
'  GetAbstructPath(path,Base)   絶対パス取得
'  GetRelatedPath(path,Base)    相対パス取得
'
'  シート名操作
'  UniqueSheetName(wb,name)     重複しないシート名取得
'
'  変数アクセス：ランタイム変数
'  ExistRt(k)                   変数有無確認
'  GetRtStr(k,[v])              変数値取得(文字列)
'  GetRtBool(k)                 変数値取得(boolean)
'  GetRtNum(k)                  変数値取得(long)
'  SetRtStr(k,[v])              変数設定(文字列)
'  SetRtBool(k,v)               変数設定(boolean)
'  SetRtNum(k,v)                変数設定(long)
'
'  変数アクセス：ブックプロパティ
'  ExistBookProp(k,[wb])        Property Exists
'  GetBookStr(k,[wb])           Get Property value
'  GetBookBool(k,[wb])          Get Property value
'  GetBookNum(k,[wb])           Get Property value
'  SetBookStr(k,v,[week,wb])    Set Property
'  SetBookBool(k,v,[week,wb])   Set Property
'  SetBookNum(k,v[week,wb])     Set Property
'  RemoveBookProp([k,wb])       remove property
'  WriteBookKeys([wb])          get book properties
'
'  変数アクセス：シートプロパティ
'  SheetPropNames(ws)           シートプロパティ名リストを取得
'  SheetPropCount(ws)           シートプロパティ数を取得
'  SheetPropIndex(ws,k)         シートプロパティ名から番号取得
'  GetSheetStr(ws,k)            シートプロパティ値取得
'  GetSheetBool(ws,k)           シートプロパティ値取得
'  GetSheetNum(ws,k)            シートプロパティ値取得
'  SetSheetStr(ws,k,v)          シートプロパティ設定
'  SetSheetBool(ws,k,v)         シートプロパティ設定
'  SetSheetNum(ws,k,v)          シートプロパティ設定
'
'  画面表示操作
'  ScreenUpdateOff()            画面チラつき防止処置
'  ScreenUpdateOn()             画面チラつき防止処置解除
'  ProgressStatusBar([i,cnt])   進行状況表示ステータスバー
'
'----------------------------------------

Option Explicit
Option Private Module

'----------------------------------------
'オブジェクト取得
'----------------------------------------

'worksheet.function
Function wsf() As WorksheetFunction
    Set wsf = WorksheetFunction
End Function

'filesystemobject
Function fso() As FileSystemObject
    Static obj As FileSystemObject
    If obj Is Nothing Then
        Set obj = CreateObject("Scripting.FileSystemObject")
    End If
    Set fso = obj
End Function

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

'----------------------------------------
'コレクション操作
'----------------------------------------

'コレクションから名前を指定して検索(配列は除く)
Function TakeByName(col As Object, s As String) As Object
    Dim v As Object
    For Each v In col
        If v.name = s Then
            Set TakeByName = v
            Exit Function
        End If
    Next v
    Set TakeByName = Nothing
End Function

'コレクションから値を指定して検索(配列は除く)
Function TakeByValue(col As Object, s As String) As Object
    Dim v As Object
    For Each v In col
        If v.Value = s Then
            Set TakeByValue = v
            Exit Function
        End If
    Next v
    Set TakeByValue = Nothing
End Function

'----------------------------------------
'名前変換
'----------------------------------------

'列名取得
Function ColumnName(n As Long) As String
    Dim s As String
    Dim i As Long, j As Long
    i = n - 1
    Do While i >= 0
        j = i Mod 26
        s = Chr(65 + j) + s
        i = (i - j) / 26 - 1
    Loop
    ColumnName = s
End Function

'キーワード名取得
'スペースはまとめて置換,半角,大文字
Function KeywordName(s As String, Optional sep As String = "_") As String
    Dim kw As String
    kw = Trim(RE_REPLACE(s, "[\s\u00A0\u3000]+", " "))
    kw = RE_REPLACE(kw, "[ _]+", sep)
    kw = StrConv(kw, vbUpperCase + vbNarrow)
    KeywordName = kw
End Function

'----------------------------------------
'パラメータ文字列
'
'文字列形式：
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
                ParamStrVal = Trim(Replace(kv(1), Chr(13), ""))
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

'パラメータ文字列からパラメータ以外取得
Function RemoveParamStrAll(s As String) As String
    Dim sa As String
    Dim v As Variant
    sa = s
    For Each v In ParamStrKeys(sa)
        sa = RemoveParamStr(sa, CStr(v))
    Next v
    RemoveParamStrAll = sa
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
'
'文字列形式：
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

'領域の値文字列取得
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

'基本名取得(パス削除、拡張子削除、複製情報削除)
Function CoreName(s As String) As String
    Dim ptn As String
    ptn = "[\(（]\d+[\)）]|\s*-\s*コピー"
    CoreName = regex(ptn).Replace(fso.GetBaseName(s), "")
End Function

'重複しないファイル名取得
Function UniqueFileName(s As String) As String
    Dim p As String
    p = s
    If fso.FileExists(p) Then
        Dim r As String, e As String, b As String
        r = fso.GetParentFolderName(p)
        p = CoreName(fso.GetFileName(p))
        e = fso.GetExtensionName(p)
        b = fso.GetBaseName(p)
        If e <> "" Then e = "." & e
        If r <> "" Then b = fso.BuildPath(r, b)
        '
        Dim i As Long
        For i = 1 To 100
            p = b & "(" & i & ")" & e
            If Not fso.FileExists(p) Then Exit For
        Next i
    End If
    UniqueFileName = p
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
    s = RE_MATCH(p, "^[\(%](\w+)[\)%]", 0, 0)
    If s <> "" Then
        s2 = Environ(s)
        If s2 <> "" Then p = s2 & Mid(p, Len(s) + 3)
    End If
    p = Replace(p, "/", "\")
    p = Replace(p, "\\", "\")
    If InStr(1, p, ":\") = 0 Then p = Base & p
    Do
        s = p
        p = RE_REPLACE(p, "\\[^\\]+\\[.][.]\\", "\")
        If s = p Then Exit Do
    Loop
    Do
        s = p
        p = RE_REPLACE(p, "\\[.]\\", "\")
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
'シート名操作
'----------------------------------------

'重複しないシート名取得
Function UniqueSheetName(wb As Workbook, name As String) As String
    Dim i As Integer: i = 1
    Dim s As String: s = name
    Do Until TakeByName(wb.Sheets, s) Is Nothing
        s = name & " (" & i & ")"
        i = i + 1
    Loop
    UniqueSheetName = s
End Function

'----------------------------------------
'ランタイム変数
'----------------------------------------

'変数有無確認
Function ExistRt(k As String) As Boolean
    ExistRt = rt_dict.Exists(k)
End Function

'変数値取得
Function GetRtStr(k As String, Optional v As String) As String
    GetRtStr = v
    With rt_dict
        If .Exists(k) Then GetRtStr = .Item(k)
    End With
End Function

Function GetRtBool(k As String) As Boolean
    Dim s As String
    s = GetRtStr(k)
    If s <> "" And Not s Like "False" And s <> "0" Then GetRtBool = True
End Function

Function GetRtNum(k As String) As Long
    GetRtNum = CLng(GetRtStr(k))
End Function

'変数設定
Sub SetRtStr(k As String, Optional v As String)
    With rt_dict
        If .Exists(k) Then .Remove k
        If v <> "" Then .Add k, v
    End With
End Sub

Sub SetRtBool(k As String, v As Boolean)
    SetRtStr k, CStr(v)
End Sub

Sub SetRtNum(k As String, v As Long)
    SetRtStr k, CStr(v)
End Sub

'変数ディクショナリ
Private Function rt_dict() As Dictionary
    Static dic As Dictionary
    If dic Is Nothing Then Set dic = New Dictionary
    Set rt_dict = dic
End Function

'----------------------------------------
'book properties
'----------------------------------------

'Property Exists
Function ExistBookProp(k As String, Optional wb As Workbook) As Boolean
    Dim p As DocumentProperty
    Set p = TakeByName(GetWorkbook(wb).CustomDocumentProperties, k)
    ExistBookProp = Not p Is Nothing
End Function

'Get Property value
Function GetBookStr(k As String, Optional wb As Workbook) As String
    On Error Resume Next
    GetBookStr = CStr(GetWorkbook(wb).CustomDocumentProperties(k))
    On Error GoTo 0
End Function

Function GetBookBool(k As String, Optional wb As Workbook) As Boolean
    On Error Resume Next
    GetBookBool = CBool(GetWorkbook(wb).CustomDocumentProperties(k))
    On Error GoTo 0
End Function

Function GetBookNum(k As String, Optional wb As Workbook) As Long
    On Error Resume Next
    GetBookNum = CLng(GetWorkbook(wb).CustomDocumentProperties(k))
    On Error GoTo 0
End Function

Private Function GetWorkbook(wb As Workbook) As Workbook
    Set GetWorkbook = IIf(wb Is Nothing, ThisWorkbook, wb)
End Function

'Set Property
Sub SetBookStr(k As String, v As String, _
    Optional week As Boolean, Optional wb As Workbook)
    SetBookProp k, v, msoPropertyTypeString, week, wb
End Sub

Sub SetBookBool(k As String, v As Boolean, _
    Optional week As Boolean, Optional wb As Workbook)
    SetBookProp k, v, msoPropertyTypeBoolean, week, wb
End Sub

Sub SetBookNum(k As String, v As Long, _
    Optional week As Boolean, Optional wb As Workbook)
    SetBookProp k, v, msoPropertyTypeNumber, week, wb
End Sub

Private Sub SetBookProp(k As String, v As Variant, _
    t As Long, week As Boolean, wb As Workbook)
    With GetWorkbook(wb)
        Dim p As DocumentProperty
        Set p = TakeByName(.CustomDocumentProperties, k)
        If Not p Is Nothing Then
            If week Then Exit Sub
            If p.Value = v Then Exit Sub
            p.Delete
        End If
        .CustomDocumentProperties.Add k, False, t, v
    End With
End Sub

'remove property
Sub RemoveBookProp(Optional k As String = "*", Optional wb As Workbook)
    Dim ptn As String
    ptn = IIf(InStr(1, k, "*", vbTextCompare), k, k & "*")
    With GetWorkbook(wb)
        Dim p As DocumentProperty
        For Each p In .CustomDocumentProperties
            If p.name Like ptn Then p.Delete
        Next p
    End With
End Sub

'get book properties
Sub WriteBookKeys(Optional wb As Workbook)
    Dim ce As Range
    Set ce = ActiveCell
    Dim p As Object
    For Each p In GetWorkbook(wb).CustomDocumentProperties
        ce.Offset(0, 0).Value = p.name
        ce.Offset(0, 1).Value = p.Type
        ce.Offset(0, 2).Value = p.Value
        Set ce = ce.Offset(1)
    Next
End Sub

'----------------------------------------
'シートプロパティ機能
'----------------------------------------

'シートプロパティ名リストを取得
Function SheetPropNames(ws As Worksheet) As String()
    Dim lst() As String
    ReDim Preserve lst(ws.CustomProperties.Count)
    Dim i As Long
    For i = 1 To ws.CustomProperties.Count
        lst(i) = ws.CustomProperties(i).name
    Next i
    SheetPropNames = lst
End Function

'シートプロパティ数を取得
Function SheetPropCount(ws As Worksheet) As Long
    SheetPropCount = ws.CustomProperties.Count
End Function

'シートプロパティ名から番号取得
Function SheetPropIndex(ws As Worksheet, k As String) As Long
    Dim i As Long
    For i = 1 To ws.CustomProperties.Count
        If ws.CustomProperties(i).name Like k Then
            SheetPropIndex = i
            Exit Function
        End If
    Next i
End Function

'シートプロパティ名からプロパティ値取得
Private Function GetSheetProp(ws As Worksheet, k As String) As CustomProperty
    Dim i As Long
    i = SheetPropIndex(ws, k)
    If i > 0 Then
        Set GetSheetProp = ws.CustomProperties(i)
        Exit Function
    End If
End Function

'シートプロパティ値取得
Function GetSheetStr(ws As Worksheet, k As String) As String
    Dim i As Long
    For i = 1 To ws.CustomProperties.Count
        If ws.CustomProperties(i).name Like k Then
            GetSheetStr = ws.CustomProperties(i).Value
            Exit Function
        End If
    Next i
End Function

Function GetSheetBool(ws As Worksheet, k As String) As Boolean
    GetSheetBool = CBool(GetSheetStr(ws, k))
End Function

Function GetSheetNum(ws As Worksheet, k As String) As Long
    GetSheetNum = CLng(GetSheetStr(ws, k))
End Function

'シートプロパティ設定
Sub SetSheetStr(ws As Worksheet, k As String, v As String)
    With ws
        Dim i As Long
        For i = 1 To .CustomProperties.Count
            If .CustomProperties(i).name Like k Then
                If .CustomProperties(i) = v Then Exit Sub
                .CustomProperties(i).Delete
                Exit For
            End If
        Next i
        .CustomProperties.Add k, v
    End With
End Sub

Sub SetSheetBool(ws As Worksheet, k As String, v As Boolean)
    SetSheetStr ws, k, CStr(v)
End Sub

Sub SetSheetNum(ws As Worksheet, k As String, v As Long)
    SetSheetStr ws, k, CStr(v)
End Sub

'----------------------------------------
'画面表示操作
'----------------------------------------

'画面チラつき防止処置
Sub ScreenUpdateOff()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Interactive = False
    Application.Cursor = xlWait
End Sub

'画面チラつき防止処置解除
Sub ScreenUpdateOn()
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'進行状況表示ステータスバー
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
    Dim s As String: s = Mid("■■■■■□□□□□", 6 - CInt(5 * p), 5)
    s = "進捗状況(" & Int(p * 100) & "%) : " & s
    Dim TM As Double: TM = (Timer - tm_start) / p * (1 - p)
    Application.StatusBar = s & " : 残り" & Int(TM) & "秒"
End Sub

