Attribute VB_Name = "Module1"
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



