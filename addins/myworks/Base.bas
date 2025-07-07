Attribute VB_Name = "Base"
'==================================
'����
'==================================

'----------------------------------------
'APi:
'  �I�u�W�F�N�g�擾
'  wsf()                        WorksheetFunction
'  fso()                        FileSystemObject
'  regex(ptn,[g,ic])            VBScript.RegExp
'
'  �R���N�V��������
'  TakeByName(col,s)            �R���N�V�������疼�O���w�肵�Č���
'  TakeByValue(col,s)           �R���N�V��������l���w�肵�Č���
'
'  ���O�ϊ�
'  ColumnName(idx)              �񖼎擾
'  KeywordName(s,[sep])         �L�[���[�h���擾
'
'  �p�����[�^������
'  ParamStrKeys(s)              �p�����[�^�����񂩂�L�[���X�g�擾
'  ParamStrVal(s,k)             �p�����[�^�����񂩂�l���擾
'  UpdateParamStr(s,k,v)        �p�����[�^������ɃL�[�E�l��ǉ��E�X�V
'  RemoveParamStr(s,k)          �p�����[�^�����񂩂獀�ڂ��폜
'  RemoveParamStrAll(s)         �p�����[�^�����񂩂�p�����[�^�ȊO�擾
'  ParamStrDict(dict,s)         �p�����[�^�����񂩂�f�B�N�V���i���쐬
'
'  �z�񕶎���
'  StrToArr(s)                  �z�񕶎��񂩂�z��֕ϊ�
'  ArrToStr(arr)                �z�񂩂�z�񕶎���֕ϊ�
'
'  �f�[�^�ϊ�
'  ColToArr(col)                �R���N�V������z��ɕϊ�
'  ArrToDict(dic,arr,[n])       �񎟔z�񕶎��񂩂�z�񎫏��ɕϊ�
'  ArrStrToDict(dic,s,[n])      �񎟔z�񕶎��񂩂�z�񎫏��ɕϊ�
'  TakeArray(arr(),[p,n])       �z��͈̔͒��o
'
'  �̈�̒l������擾
'  StrRange(s)                  �̈�̒l������擾
'
'  �p�X����
'  CoreName(s)                  ��{���擾(�p�X�폜�A�g���q�폜�A�������폜)
'  UniqueFileName(s)            �d�����Ȃ��t�@�C�����擾
'  GetShortPath(path,[pc])      �Z�k�p�X�擾
'  GetAbstructPath(path,Base)   ��΃p�X�擾
'  GetRelatedPath(path,Base)    ���΃p�X�擾
'
'  �V�[�g������
'  UniqueSheetName(wb,name)     �d�����Ȃ��V�[�g���擾
'
'  �ϐ��A�N�Z�X�F�����^�C���ϐ�
'  ExistRt(k)                   �ϐ��L���m�F
'  GetRtStr(k,[v])              �ϐ��l�擾(������)
'  GetRtBool(k)                 �ϐ��l�擾(boolean)
'  GetRtNum(k)                  �ϐ��l�擾(long)
'  SetRtStr(k,[v])              �ϐ��ݒ�(������)
'  SetRtBool(k,v)               �ϐ��ݒ�(boolean)
'  SetRtNum(k,v)                �ϐ��ݒ�(long)
'
'  �ϐ��A�N�Z�X�F�u�b�N�v���p�e�B
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
'  �ϐ��A�N�Z�X�F�V�[�g�v���p�e�B
'  SheetPropNames(ws)           �V�[�g�v���p�e�B�����X�g���擾
'  SheetPropCount(ws)           �V�[�g�v���p�e�B�����擾
'  SheetPropIndex(ws,k)         �V�[�g�v���p�e�B������ԍ��擾
'  GetSheetStr(ws,k)            �V�[�g�v���p�e�B�l�擾
'  GetSheetBool(ws,k)           �V�[�g�v���p�e�B�l�擾
'  GetSheetNum(ws,k)            �V�[�g�v���p�e�B�l�擾
'  SetSheetStr(ws,k,v)          �V�[�g�v���p�e�B�ݒ�
'  SetSheetBool(ws,k,v)         �V�[�g�v���p�e�B�ݒ�
'  SetSheetNum(ws,k,v)          �V�[�g�v���p�e�B�ݒ�
'
'  ��ʕ\������
'  ScreenUpdateOff()            ��ʃ`�����h�~���u
'  ScreenUpdateOn()             ��ʃ`�����h�~���u����
'  ProgressStatusBar([i,cnt])   �i�s�󋵕\���X�e�[�^�X�o�[
'
'----------------------------------------

Option Explicit
Option Private Module

'----------------------------------------
'�I�u�W�F�N�g�擾
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
'�R���N�V��������
'----------------------------------------

'�R���N�V�������疼�O���w�肵�Č���(�z��͏���)
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

'�R���N�V��������l���w�肵�Č���(�z��͏���)
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
'���O�ϊ�
'----------------------------------------

'�񖼎擾
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

'�L�[���[�h���擾
'�X�y�[�X�͂܂Ƃ߂Ēu��,���p,�啶��
Function KeywordName(s As String, Optional sep As String = "_") As String
    Dim kw As String
    kw = Trim(RE_REPLACE(s, "[\s\u00A0\u3000]+", " "))
    kw = RE_REPLACE(kw, "[ _]+", sep)
    kw = StrConv(kw, vbUpperCase + vbNarrow)
    KeywordName = kw
End Function

'----------------------------------------
'�p�����[�^������
'
'������`���F
'  <text> = [ <line> \n ] <line>
'  <line> = \s* <key> \s* : \s* <val> \s* | .+
'  <key>  = \w+
'  <val>  = .+
'----------------------------------------

'�p�����[�^�����񂩂�L�[���X�g�擾
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

'�p�����[�^�����񂩂�l���擾
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

'�p�����[�^������ɃL�[�E�l��ǉ��E�X�V
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

'�p�����[�^�����񂩂獀�ڂ��폜
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

'�p�����[�^�����񂩂�p�����[�^�ȊO�擾
Function RemoveParamStrAll(s As String) As String
    Dim sa As String
    Dim v As Variant
    sa = s
    For Each v In ParamStrKeys(sa)
        sa = RemoveParamStr(sa, CStr(v))
    Next v
    RemoveParamStrAll = sa
End Function

'�p�����[�^�����񂩂�f�B�N�V���i���쐬
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
'�z�񕶎���
'
'������`���F
'  <text> = [ <rows> ; ] <rows>
'  <rows> = [ \s* <item> \s+ , ] \s* <item> \s*
'  <item> = \w+
'----------------------------------------

'�z�񕶎��񂩂�z��֕ϊ�
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

'�z�񂩂�z�񕶎���֕ϊ�
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
'�f�[�^�ϊ�
'----------------------------------------

'�R���N�V������z��ɕϊ�
Function ColToArr(col As Collection) As Variant()
    Dim arr() As Variant
    ReDim arr(0 To col.Count - 1)
    Dim i As Integer
    For i = 1 To col.Count
        arr(i - 1) = col.Item(i)
    Next i
    ColToArr = arr
End Function

'�񎟔z�񕶎��񂩂�z�񎫏��ɕϊ�
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

'�񎟔z�񕶎��񂩂�z�񎫏��ɕϊ�
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

'�z��͈̔͒��o
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
'�̈�̒l������擾
'----------------------------------------

'�̈�̒l������擾
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
'�p�X����
'----------------------------------------

'��{���擾(�p�X�폜�A�g���q�폜�A�������폜)
Function CoreName(s As String) As String
    Dim ptn As String
    ptn = "[\(�i]\d+[\)�j]|\s*-\s*�R�s�["
    CoreName = regex(ptn).Replace(fso.GetBaseName(s), "")
End Function

'�d�����Ȃ��t�@�C�����擾
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

'�Z�k�p�X�擾
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

'��΃p�X�擾
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

'���΃p�X�擾
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
'�V�[�g������
'----------------------------------------

'�d�����Ȃ��V�[�g���擾
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
'�����^�C���ϐ�
'----------------------------------------

'�ϐ��L���m�F
Function ExistRt(k As String) As Boolean
    ExistRt = rt_dict.Exists(k)
End Function

'�ϐ��l�擾
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

'�ϐ��ݒ�
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

'�ϐ��f�B�N�V���i��
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
'�V�[�g�v���p�e�B�@�\
'----------------------------------------

'�V�[�g�v���p�e�B�����X�g���擾
Function SheetPropNames(ws As Worksheet) As String()
    Dim lst() As String
    ReDim Preserve lst(ws.CustomProperties.Count)
    Dim i As Long
    For i = 1 To ws.CustomProperties.Count
        lst(i) = ws.CustomProperties(i).name
    Next i
    SheetPropNames = lst
End Function

'�V�[�g�v���p�e�B�����擾
Function SheetPropCount(ws As Worksheet) As Long
    SheetPropCount = ws.CustomProperties.Count
End Function

'�V�[�g�v���p�e�B������ԍ��擾
Function SheetPropIndex(ws As Worksheet, k As String) As Long
    Dim i As Long
    For i = 1 To ws.CustomProperties.Count
        If ws.CustomProperties(i).name Like k Then
            SheetPropIndex = i
            Exit Function
        End If
    Next i
End Function

'�V�[�g�v���p�e�B������v���p�e�B�l�擾
Private Function GetSheetProp(ws As Worksheet, k As String) As CustomProperty
    Dim i As Long
    i = SheetPropIndex(ws, k)
    If i > 0 Then
        Set GetSheetProp = ws.CustomProperties(i)
        Exit Function
    End If
End Function

'�V�[�g�v���p�e�B�l�擾
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

'�V�[�g�v���p�e�B�ݒ�
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
'��ʕ\������
'----------------------------------------

'��ʃ`�����h�~���u
Sub ScreenUpdateOff()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Interactive = False
    Application.Cursor = xlWait
End Sub

'��ʃ`�����h�~���u����
Sub ScreenUpdateOn()
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'�i�s�󋵕\���X�e�[�^�X�o�[
Sub ProgressStatusBar(Optional i As Long = 1, Optional cnt As Long = 1)
    Static tm_start As Double
    If i < 1 Then
        tm_start = Timer
        Application.StatusBar = "�i����(0%)"
        Exit Sub
    End If
    If i >= cnt Then
        Application.StatusBar = False
        Exit Sub
    End If
    Dim p As Double: p = i / cnt
    Dim s As String: s = Mid("��������������������", 6 - CInt(5 * p), 5)
    s = "�i����(" & Int(p * 100) & "%) : " & s
    Dim TM As Double: TM = (Timer - tm_start) / p * (1 - p)
    Application.StatusBar = s & " : �c��" & Int(TM) & "�b"
End Sub

