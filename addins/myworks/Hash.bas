Attribute VB_Name = "Hash"
Option Explicit
'Option Private Module

'========================================
'汎用計算モジュール
'========================================

'----------------------------------------
'ハッシュ計算
'----------------------------------------

Public Function HASH_MD5(s As String) As String
    HASH_MD5 = Hash("MD5CryptoServiceProvider", s)
End Function

Public Function HASH_SHA1(s As String) As String
    HASH_SHA1 = Hash("SHA1CryptoServiceProvider", s)
End Function

Public Function HASH_SHA256(s As String) As String
    HASH_SHA256 = Hash("SHA256Managed", s)
End Function

Private Function Hash(alg As String, s As String) As String
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
    Hash = UCase(res)
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

