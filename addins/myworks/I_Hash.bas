Attribute VB_Name = "I_Hash"
'========================================
'汎用計算モジュール
'========================================

Option Explicit

'----------------------------------------
'ハッシュ値計算
'HASH_MD5(s)
'HASH_SHA1(s)
'HASH_SHA256(s)
'----------------------------------------

Function HASH_MD5(s As String) As String
    HASH_MD5 = Hash("MD5CryptoServiceProvider", s)
End Function

Function HASH_SHA1(s As String) As String
    HASH_SHA1 = Hash("SHA1CryptoServiceProvider", s)
End Function

Function HASH_SHA256(s As String) As String
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
'ファイルハッシュ値計算関数
'FILEHASH_MD5(s,[p])
'FILEHASH_SHA1(s,[p])
'FILEHASH_SHA256(s,[p])
'----------------------------------------

Function FILEHASH_MD5(s As String, Optional p As String) As String
    If p <> "" Then s = fso.BuildPath(s, p)
    s = GetAbstructPath(s, ActiveWorkbook.path)
    FILEHASH_MD5 = FileHash(s, "MD5CryptoServiceProvider")
End Function

Function FILEHASH_SHA1(s As String, Optional p As String) As String
    If p <> "" Then s = fso.BuildPath(s, p)
    s = GetAbstructPath(s, ActiveWorkbook.path)
    FILEHASH_SHA1 = FileHash(s, "SHA1CryptoServiceProvider")
End Function

Function FILEHASH_SHA256(s As String, Optional p As String) As String
    If p <> "" Then s = fso.BuildPath(s, p)
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

