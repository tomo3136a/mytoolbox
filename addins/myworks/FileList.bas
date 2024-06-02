Attribute VB_Name = "FileList"
'==================================
'パス操作取得
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'機能呼び出し
'----------------------------------------

'パス取得
Sub PathMenu(mode As Integer, ra As Range)
    Application.ScreenUpdating = False
    '
    Select Case mode
    Case 1
        'ファイルリスト取得
        Call GetFileList(ra)
    Case 2
        'フォルダーパス取得
        Call GetFolderPath(ra, GetParamBool("path", 1))
    Case 3
        'ファイルパス取得
        Call GetFilePath(ra, GetParamBool("path", 1))
    Case 4
        'パス名変換
        Call ChangePath(ra)
    Case 5
        'パス区切り変更
        Call ChangePathSepalater(ra)
    Case 6
        '絶対パス化
        Call ToAbustoractPath(ra)
    Case 7
        '相対パス化
        Call ToRelatedPath(ra)
    Case 8
        'ベースパス化
        Call ToGetBasePath2(ra)
    Case 9
        'パスセグメント化
        Call ToPathSegment(ra)
    End Select
    '
    Application.ScreenUpdating = True
End Sub

'基準のパス取得(ツリー)
Private Sub ToPathSegment(ra As Range)
    Dim ce As Range
    For Each ce In ra
        Dim p As String, s As String
        p = ce.Value
        If Right(p, 1) = "\" Or Right(p, 1) = "/" Then
            s = re_match(p, "[^\\/]+[\\/]$")
            p = re_replace(p, "[^\\/]+[\\/]", "    ")
            p = p & s
            p = Mid(p, 5)
        Else
            p = re_replace(p, "[^\\/]+[\\/]", "    ")
        End If
        If p <> ce.Value Then ce.Value = p
    Next ce
End Sub

'基準のパス取得(ツリー)
Private Sub ToGetBasePath2(ra As Range)
    Dim ce As Range
    For Each ce In ra
        Dim p As String
        p = GetBasePath2(ce)
        If p <> ce.Value Then ce.Value = p
    Next ce
End Sub

Private Function GetBasePath2(ra As Range) As String
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    If Left(ce.Value, 1) <> " " Then
        GetBasePath2 = ce.Value
        Exit Function
    End If
    '
    Dim s As String, p As String
    Dim i As Long
    For i = 0 To ce.Row - 1
        s = Trim(ce.Offset(-i).Value)
        If s = "" Then Exit For
        If Right(s, 1) = "\" Or Right(s, 1) = "/" Then
            p = s & p
            If Left(ce.Offset(-i).Value, 1) <> " " Then Exit For
        ElseIf p = "" Then
            p = s
        End If
    Next i
    GetBasePath2 = p
End Function

'基準のパス取得
Private Function GetBasePath(ra As Range, Optional n As Integer = 1) As String
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    '
    Dim s As String, p As String
    Dim i As Long
    For i = n To ce.Row - 1
        s = ce.Offset(-i).Value
        If s = "" Then Exit For
        If InStr(1, s, "\") Then p = s
        If InStr(1, s, "/") Then p = s
    Next i
    GetBasePath = p
End Function

'短縮パス⇔フルパス変換
Private Sub ChangePath(ra As Range)
    Dim c As String
    c = Left(ra.Cells(1, 1).Value, 1)
    If c <> "%" And c <> "(" Then c = ""
    '
    Dim s1 As String, s2 As String
    Dim ce As Range
    For Each ce In ra
        s1 = ce.Value
        If c = "" Then
            s2 = GetShortPath(s1)
        Else
            s2 = GetAbstructPath(s1, ra.Parent.Parent.path & "\")
        End If
        If s1 <> s2 Then ce.Value = s2
    Next ce
End Sub

'パス区切り変換
Private Sub ChangePathSepalater(ra As Range)
    Dim c As String
    If InStr(1, ra.Cells(1, 1).Value, "\") Then c = "\"
    If InStr(1, ra.Cells(1, 1).Value, "/") Then c = "/"
    '
    Dim s1 As String, s2 As String
    Dim ce As Range
    For Each ce In ra
        s1 = ce.Value
        s2 = s1
        If c = "/" Then s2 = Replace(s2, "/", "\")
        If c = "\" Then s2 = Replace(s2, "\", "/")
        If s1 <> s2 Then ce.Value = s2
    Next ce
End Sub

'絶対パスに変更
Private Sub ToAbustoractPath(ra As Range)
    Dim base As String
    base = GetBasePath(ra, 1)
    base = GetAbstructPath(base, ra.Parent.Parent.path & "\")
    '
    Dim s1 As String, s2 As String
    Dim ce As Range
    For Each ce In ra
        s1 = ce.Value
        s2 = GetAbstructPath(s1, base)
        If s1 <> s2 Then ce.Value = s2
    Next ce
End Sub

'相対パスに変更
Private Sub ToRelatedPath(ra As Range)
    Dim base As String
    base = GetBasePath(ra, 1)
    base = GetAbstructPath(base, ra.Parent.Parent.path & "\")
    '
    Dim s1 As String, s2 As String
    Dim ce As Range
    For Each ce In ra
        s1 = ce.Value
        s2 = GetRelatedPath(s1, base)
        If s1 <> s2 Then ce.Value = s2
    Next ce
End Sub

'フォルダーパス取得
Private Sub GetFolderPath(ra As Range, link As Boolean)
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim path As String
    path = ce.Value
    If path = "" Then path = ActiveWorkbook.path
    path = SelectFolder(GetAbstructPath(path, ra.Parent.Parent.path & "\"))
    If path = "" Then Exit Sub
    If Right(path, 1) <> "\" Then path = path & "\"
    '
    ce.Clear
    ce.Value = GetShortPath(path)
    If Not link Then Exit Sub
    Application.CutCopyMode = False
    ce.Worksheet.Hyperlinks.Add Anchor:=ce, address:=path
End Sub

Private Function GetFolder(ra As Range) As String
    Dim path As String
    path = GetBasePath(ra)
    If path = "" Then
        path = SelectFolder(ActiveWorkbook.path)
        If path = "" Then path = ActiveWorkbook.path
    End If
    GetFolder = GetAbstructPath(path, ra.Parent.Parent.path & "\")
End Function


'ファイル名取得
Private Sub GetFilePath(ra As Range, link As Boolean)
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim base As String
    base = GetBasePath(ra)
    base = GetAbstructPath(base, ra.Parent.Parent.path & "\")
    '
    Dim col As Variant
    Set col = SelectFiles(base)
    If col.Count = 0 Then Exit Sub
    '
    Dim clrf As Boolean
    Dim fi As Variant
    For Each fi In col
        ce.Value = GetRelatedPath(CStr(fi), base)
        If link Then
            Application.CutCopyMode = False
            ce.Worksheet.Hyperlinks.Add _
                Anchor:=ce, address:=fi
        End If
        Set ce = ce.Offset(1)
        clrf = True
    Next fi
    If clrf Then ce.Clear
End Sub

'ファイルリスト取得
Private Sub GetFileList(ra As Range)
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Dim path As String
    Dim p As String
    path = ce.Value
    If path = "" Then path = GetFolder(ra)
    If path = "" Then Exit Sub
    If Right(path, 1) <> "\" Then path = path & "\"
    p = GetShortPath(path)
    If path <> p Then ce.Value = p
    path = GetAbstructPath(p, ra.Parent.Parent.path & "\")
    '
    Set ce = ce.Offset(2)
    Call GetFileListSubFolder(ce, path, 1, "")
    '
    Dim h
    h = Array("番号", "名前", "階層", "種別", "サイズ", "編集日")
    Set ce = ra.Cells(1, 1).Offset(1)
    Range(ce, ce.Offset(0, UBound(h))) = h
    Call Waku(ce, fit:=True)
End Sub

Private Sub GetFileListSubFolder(ByRef ra As Range, path As String, n As Integer, sp As String)
    On Error Resume Next
    Dim no As Integer
    no = ra.Offset(-1).Value
    Dim obj As Variant
    Dim re As Object
    Set re = regex("^[~._]|bak$|tmp$")
    '
    'フォルダリスト
    For Each obj In fso.GetFolder(path).SubFolders
        If GetParamBool("path", 4) Or Not re.test(obj.name) Then
            Dim p As String
            p = fso.BuildPath(path, obj.name)
            If GetParamBool("path", 2) Then
                no = no + 1
                ra.Clear
                ra.Value = no
                With obj
                    ra.Offset(0, 1).Value = sp & .name & "/"
                    ra.Offset(0, 2).Value = n
                    ra.Offset(0, 3).Value = "フォルダ"
                    ra.Offset(0, 4).Style = "Comma [0]"
                    ra.Offset(0, 4).Value = .Size
                    ra.Offset(0, 5).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss;@"
                    ra.Offset(0, 5).Value = .DateLastModified
                End With
                Set ra = ra.Offset(1)
            End If
            If GetParamBool("path", 3) Then
                If Left(obj.name, 1) = "." Then
                ElseIf Left(obj.name, 1) = "_" Then
                ElseIf GetParamBool("path", 2) Then
                    Call GetFileListSubFolder(ra, p, n + 1, sp + "    ")
                    no = ra.Offset(-1).Value
                Else
                    Call GetFileListSubFolder(ra, p, n + 1, fso.BuildPath(sp, obj.name))
                    no = ra.Offset(-1).Value
                End If
            End If
        End If
    Next obj
    '
    'ファイルリスト
    For Each obj In fso.GetFolder(path).Files
        If GetParamBool("path", 4) Or Not re.test(obj.name) Then
            no = no + 1
            ra.Value = no
            With obj
                p = fso.BuildPath(path, .name)
                If GetParamBool("path", 2) Then
                    ra.Offset(0, 1).Value = sp & .name
                Else
                    ra.Offset(0, 1).Value = fso.BuildPath(sp, .name)
                End If
                ra.Offset(0, 2).Value = n
                ra.Offset(0, 3).Value = "ファイル"
                ra.Offset(0, 4).Style = "Comma [0]"
                ra.Offset(0, 4).Value = .Size
                ra.Offset(0, 5).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss;@"
                ra.Offset(0, 5).Value = .DateLastModified
            End With
            Set ra = ra.Offset(1)
        End If
    Next obj
    '
    On Error GoTo 0
End Sub

