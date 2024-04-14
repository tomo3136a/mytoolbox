Attribute VB_Name = "Util"
Option Explicit
'Option Private Module

'========================================
'汎用計算モジュール
'========================================

'----------------------------------------
'ハッシュ計算
'----------------------------------------

Public Function HASH_MD5(s As String) As String
    HASH_MD5 = hash("MD5CryptoServiceProvider", s)
End Function

Public Function HASH_SHA1(s As String) As String
    HASH_SHA1 = hash("SHA1CryptoServiceProvider", s)
End Function

Public Function HASH_SHA256(s As String) As String
    HASH_SHA256 = hash("SHA256Managed", s)
End Function

Private Function hash(alg As String, s As String) As String
    Dim utf8 As Object
    Dim bytes() As Byte
    Dim Conv As Object
    Dim code() As Byte
    Dim i As Integer
    Dim res As String
    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    bytes = utf8.GetBytes_4(s)
    Set Conv = CreateObject("System.Security.Cryptography." & alg)
    code = Conv.ComputeHash_2(bytes)
    For i = LBound(code) To UBound(code)
        res = res & Right("0" & Hex(code(i)), 2)
    Next i
    hash = UCase(res)
End Function

'----------------------------------------
'ファイルハッシュ計算
'----------------------------------------

Public Function FILEHASH_MD5(s As String, Optional s2 As String) As String
    If s2 <> "" Then s = fso.BuildPath(s, s2)
    s = GetAbstructPath(s, ActiveWorkbook.path)
    FILEHASH_MD5 = FileHash(s, "MD5CryptoServiceProvider")
End Function

Public Function FILEHASH_SHA1(s As String, Optional s2 As String) As String
    If s2 <> "" Then s = fso.BuildPath(s, s2)
    s = GetAbstructPath(s, ActiveWorkbook.path)
    FILEHASH_SHA1 = FileHash(s, "SHA1CryptoServiceProvider")
End Function

Public Function FILEHASH_SHA256(s As String, Optional s2 As String) As String
    If s2 <> "" Then s = fso.BuildPath(s, s2)
    s = GetAbstructPath(s, ActiveWorkbook.path)
    FILEHASH_SHA256 = FileHash(s, "SHA256Managed")
End Function

Private Function FileHash(path As String, alg As String) As String
    Dim bytes() As Byte
    Dim Conv As Object
    Dim code() As Byte
    Dim i As Integer
    Dim res As String
    
    bytes = GetFileBytes(path)
    Set Conv = CreateObject("System.Security.Cryptography." & alg)
    code = Conv.ComputeHash_2(bytes)
    For i = LBound(code) To UBound(code)
        res = res & Right("0" & Hex(code(i)), 2)
    Next i
    FileHash = UCase(res)
End Function

Private Function GetFileBytes(path As String) As Byte()
    Dim h As Long
    Dim ba() As Byte
    h = FreeFile

    If Dir(path) = "" Then
        err.Raise 53
    End If
    Open path For Binary Access Read As h
    ReDim ba(LOF(h) - 1&) As Byte
    Get h, , ba
    Close h
    '
    GetFileBytes = ba
    Erase ba
End Function

'----------------------------------------
'文字列抜き出し
'----------------------------------------

'英字文字抜き出し
Public Function ToAlpha(s As String) As String
    Dim ss As String
    Dim i As Integer
    For i = 1 To Len(s)
        Dim c As String
        c = Mid(s, i, 1)
        If c Like "[A-Z]" Then ss = ss + c
    Next i
    ToAlpha = ss
End Function

'数字抜き出し
Public Function ToNum(s As String) As Integer
    Dim ss As String
    Dim i As Integer
    For i = 1 To Len(s)
        Dim c As String
        c = Mid(s, i, 1)
        If c Like "[0-9]" Then ss = ss + c
    Next i
    ToNum = CInt(ss)
End Function

'----------------------------------------
'正規表現
'----------------------------------------

'文字列判定
Public Function RE_TEST(s As String, ptn As String) As Boolean
    Dim re As Object
    Set re = regex(ptn)
    RE_TEST = re.test(s)
End Function

'文字列抽出
Public Function RE_MATCH(s As String, ptn As String, _
        Optional idx As Integer = 0, _
        Optional idx2 As Integer = -1) As Variant
    Dim re As Object
    Set re = regex(ptn)
    Dim mc As Object
    Set mc = re.Execute(s)
    
    If idx >= mc.Count Then
        RE_MATCH = ""
    ElseIf idx < 0 Then
        RE_MATCH = mc.Count
    ElseIf idx2 < 0 Then
        RE_MATCH = mc(idx).Value
    ElseIf idx2 < mc(idx).SubMatches.Count Then
        RE_MATCH = mc(idx).SubMatches(idx2)
    Else
        RE_MATCH = ""
    End If
End Function

'文字列置き換え
Public Function RE_REPLACE(s As String, ptn As String, rep As String) As String
    Dim re As Object
    Set re = regex(ptn)
    RE_REPLACE = re.Replace(s, rep)
End Function
