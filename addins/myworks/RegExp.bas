Attribute VB_Name = "RegExp"
'----------------------------------------
'正規表現関数
' re_test(s,ptn)            文字列有無
' re_match(s,ptn,idx,idx2)  文字列抽出
' re_replace(s,ptn,rep)     文字列置き換え
' re_split(s,ptn)           文字列分割
' re_extract(col,ptn)       配列からマッチした文字列を抽出
'----------------------------------------

Option Explicit

'regex(VBScript.RegExp)
Private Function regex( _
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
Function RE_TEST(s As String, ptn As String) As Boolean
    On Error Resume Next
    RE_TEST = regex(ptn).Test(s)
    On Error GoTo 0
End Function

'文字列抽出
Function RE_MATCH(s As String, ptn As String, _
        Optional idx As Integer = 0, _
        Optional idx2 As Integer = -1) As Variant
    On Error Resume Next
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
    On Error GoTo 0
End Function

'文字列置き換え
Function RE_REPLACE(s As String, ptn As String, rep As String) As String
    On Error Resume Next
    RE_REPLACE = regex(ptn).Replace(s, rep)
    On Error GoTo 0
End Function

'文字列分割
Function RE_SPLIT(s As String, ptn As String) As String()
    RE_SPLIT = Split(regex(ptn).Replace(s, Chr(7)), Chr(7))
End Function

'配列からマッチした文字列を抽出
Function RE_EXTRACT(col As Variant, ptn As String) As Variant
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
            If i > UBound(arr) Then
                ReDim Preserve arr(UBound(arr) + 50)
            End If
            arr(i) = s
            i = i + 1
        End If
    Next v
    If i < 1 Then Exit Function
    ReDim Preserve arr(i - 1)
    RE_EXTRACT = arr
End Function

