Attribute VB_Name = "MTimeChart"
'==================================
'タイムチャート
'==================================

Option Explicit
Option Private Module

'----------------------------------------
'図面シート
'----------------------------------------

'図面シートを追加
Public Sub AddDrawingSheet()
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=ActiveSheet)
    ws.Cells.ColumnWidth = 2.5
End Sub

'----------------------------------------
'タイミング図説明
'----------------------------------------

Public Sub HelpTimingChart()
    Dim msg As String
    msg = "説明:" & vbLf & vbLf
    msg = msg & "論理値0" & vbTab & vbTab & "0" & vbLf
    msg = msg & "論理値1" & vbTab & vbTab & "1" & vbLf
    msg = msg & "論地否定(NOT)" & vbTab & "Y = 1-A" & vbLf
    msg = msg & "論理積(AND)" & vbTab & "Y = A*B*..." & vbLf
    msg = msg & "論理和(OR)" & vbTab & "Y = 1-(1-A)*(1-B)*..." & vbLf
    msg = msg & "排他的論理和(XOR)" & vbTab & "Y = mod(A+B+...,2)" & vbLf
    msg = msg & "マルチプレクサ(MUX)" & vbTab & "Y = S * (B - A) + A" & vbLf
    msg = msg & "フリップフロップ(DFF)" & vbTab & "Q = C * (D - Q[-1]) + Q[-1]"
    MsgBox msg, Title:="タイミング図"
End Sub

'----------------------------------------
'タイミング図描画
'----------------------------------------

Public Sub DrawTimeChart(Optional mode As Long)
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    'データ範囲取得
    Dim ce As Range, ra As Range
    Set ra = Selection
    Dim c As Long, cn As Long, r As Long
    cn = ra.Columns.Count
    
    For r = 1 To ra.Rows.Count
        If ra(r, 1) <> "" And ra(r, 2) <> "" Then
            c = ra(r, 1).End(xlToRight).Column - ra.Column + 1
            If cn < c Then cn = c
        End If
    Next r
    
    If cn < 2 Then
        Dim s As String
        Dim v As Variant
        s = "横サイズを指定してください。"
        v = Application.InputBox(s, app_name, "16", , , , , Type:=1)
        If TypeName(v) <> "Boolean" Then cn = v
    End If
    If cn < 1 Then Exit Sub
    Set ra = ra.Resize(ra.Rows.Count, cn)
    
    '描画
    For r = 1 To ra.Rows.Count
        Set ce = ra(r, 1)
        If ce <> "" Then
            Set ce = ce.Resize(1, cn)
            Select Case mode
            Case 1: Call DrawTimeChart_1(ce)
            Case 2: Call DrawTimeChart_2(ce)
            End Select
        End If
    Next r
End Sub

Private Sub DrawTimeChart_1(ByVal ra As Range)
    Dim ce As Range
    Set ce = ra(1, 1)
    
    'データ範囲取得
    If ce.Offset(, 1) <> "" Then
        Set ra = ce.Worksheet.Range(ce, ce.End(xlToRight))
    End If
    
    Dim s1 As String, s2 As String
    s1 = ra(1, 1).Address(False, False)
    s2 = ce.Offset(0, -1).Address(False, False)

    With ra
        .FormatConditions.Delete
        
        .FormatConditions.Add Type:=xlExpression, Operator:=xlEqual, _
            Formula1:="=mod(" & s1 & ",2)=0"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Borders(xlBottom)
            .LineStyle = xlContinuous
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .FormatConditions(1).StopIfTrue = False
        
        .FormatConditions.Add Type:=xlExpression, Operator:=xlEqual, _
            Formula1:="=mod(" & s1 & ",2)=1"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Borders(xlTop)
            .LineStyle = xlContinuous
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .FormatConditions(1).StopIfTrue = False
    
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=" & s1 & "<>" & s2
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Borders(xlLeft)
            .LineStyle = xlContinuous
            .TintAndShade = 0
            .Weight = xlThin
        End With
        .FormatConditions(1).StopIfTrue = False
    
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=LEN(TRIM(" & s1 & "))=0"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Borders(xlTop).LineStyle = xlNone
        .FormatConditions(1).Borders(xlBottom).LineStyle = xlNone
        .FormatConditions(1).StopIfTrue = False
    
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
            "=ISERROR(" & s1 & ")"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.14996795556505
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub

Private Sub DrawTimeChart_2(ByVal ra As Range)
    Dim ws As Worksheet
    Set ws = ra.Worksheet
    
    Dim sh As Shape
    Dim fb As FreeformBuilder
    Dim ns As Collection
    Set ns = New Collection
    
    '描画位置補正
    Dim ce As Range
    Set ce = ra.Cells(1, 1)
    Do While Left(ce.Value, 1) = "!"
        Set ce = ce.Offset(-1)
    Loop
    If ce.Value = "" Then Set ce = ce.Offset(1)
    
    '状態位置取得
    Dim py(0 To 9) As Double
    Dim i As Long, j As Long
    j = 1
    For i = 0 To 9
        py(i) = ce.Offset(j).Top
        If ce.Row + j > 1 Then j = j - 1
    Next i
    
    Dim re As Object
    Set re = regex("[!#=.\[\]]")

    Dim x As Double, dx As Double, x0 As Double, x1 As Double, x2 As Double
    Dim y As Double, dy As Double, y0 As Double, y1 As Double, y2 As Double
    
    Dim xi As Double, yi As Double, yj As Double
    Dim yn As Long
    
    Dim ss As String, s As String, s0 As String
    Dim b_skip As Boolean
    Dim b_close As Boolean
    
    xi = ra.Left: yi = -1
    
    For Each ce In ra
        ss = CStr(ce.Value)
        i = Len(re.Replace(ss, ""))
        x = ce.Left: dx = (ce.Offset(, 1).Left - x) / IIf(i < 1, 1, i)
        x1 = x: x2 = x1 + dx
        '
        For i = 1 To Len(ss)
            s0 = s: s = Mid(ss, i, 1)
            x0 = x: If Not re.Test(s) Then x = x + dx
            '
            ' (x0,y0,s0)-(x1,y1,s)-[x,y,s]-(x2,y2,s)
            '0-9: レベル
            '-: 前を引き継ぐ
            '<: 位置保存
            '>: 位置保存
            '=: 位置保存
            Select Case s
            Case "0", "1", "2", "3", "4": yn = CLng(s): y = py(yn)
            Case "5", "6", "7", "8", "9": yn = CLng(s): y = py(yn)
            Case "\": yn = IIf(yn > 0, yn - 1, yn): y1 = y: y = py(yn): y0 = y
            Case "/": yn = IIf(yn < 9, yn + 1, yn): y1 = y: y = py(yn): y0 = y
            Case "=": xi = x0: y = y0: yn = yn Xor 1: yi = py(yn)
            Case "[": xi = x0: yi = y
            Case "<": xi = x0: yi = y
            Case ".": b_close = True
            End Select
            
            If s = "x" Then
                If yi < 0 Then y = py(yn): yn = yn Xor 1: y0 = py(yn): yi = y0
                If fb Is Nothing Then
                    Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, x0, y)
                End If
                fb.AddNodes msoSegmentLine, msoEditingAuto, x0, y
                fb.AddNodes msoSegmentLine, msoEditingAuto, x, yi
                Set sh = fb.ConvertToShape
                ns.Add sh.name
                Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, xi, yi)
                fb.AddNodes msoSegmentLine, msoEditingAuto, x0, yi
                fb.AddNodes msoSegmentLine, msoEditingAuto, x, y
                xi = x: x0 = x
            End If
            
            If x0 <> xi And y = yi And yi <> -1 And s <> "-" Then
                If Not fb Is Nothing Then
                    fb.AddNodes msoSegmentLine, msoEditingAuto, x0, y
                    fb.AddNodes msoSegmentLine, msoEditingAuto, xi, y
                    Set sh = fb.ConvertToShape
                    ns.Add sh.name
                    Set fb = Nothing
                End If
                xi = x0
                y = yi: yi = -1
            End If
            
            If fb Is Nothing Then
                If y <> 0 Then
                    Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, xi, y)
                End If
            ElseIf y <> y0 Then
                fb.AddNodes msoSegmentLine, msoEditingAuto, x0, y
            End If
            '
            If InStr("-<>[", s) = 0 Then
                If Not fb Is Nothing Then
                    fb.AddNodes msoSegmentLine, msoEditingAuto, x, y
                End If
                y0 = y
            End If
            
            '切断処理
            If b_close Then
                b_close = False
                If Not fb Is Nothing Then
                    Set sh = fb.ConvertToShape
                    ns.Add sh.name
                    Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, x, y)
                End If
            End If
            
        Next i
    Next ce
    
    If Not fb Is Nothing Then
        fb.AddNodes msoSegmentLine, msoEditingAuto, x, y
        Set sh = fb.ConvertToShape
        ns.Add sh.name
        Set fb = Nothing
    End If
    If x <> xi And yi <> -1 Then
        Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, xi, yi)
        fb.AddNodes msoSegmentLine, msoEditingAuto, x, yi
        Set sh = fb.ConvertToShape
        ns.Add sh.name
        Set fb = Nothing
    End If

    If ns.Count > 1 Then Set sh = ws.Shapes.Range(ColToArr(ns)).Group

End Sub

'----------------------------------------
'タイミングデータ操作
'----------------------------------------

Public Sub ApplyTimeChart(Optional mode As Long)
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Dim ra As Range
    Set ra = Selection
    If ra(1, 1) <> "" And ra(1, 1).Offset(, 1) <> "" Then
       Set ra = ra.Worksheet.Range(ra(1, 1), ra(1, 1).End(xlToRight))
    End If
    
    Select Case mode
    Case 1: Call InvertTimeChart(ra)
    Case 2: Call MaskTimeChart0(ra)
    Case 3: Call MaskTimeChart1(ra)
    End Select
End Sub

'INVERT
Private Sub InvertTimeChart(ByVal ra As Range)
    Dim exp As String
    Dim ce As Range
    Set ce = ra(1, 1)
    exp = ce.Formula
    If ce.HasFormula Then
        If Left(exp, 3) = "=1-" Then
            exp = "=" & Mid(exp, 4)
        Else
            exp = "=1-" & Mid(exp, 2)
        End If
    End If
    If exp <> "" Then ra.Formula = exp
End Sub

'MASK0
Private Sub MaskTimeChart0(ByVal ra As Range)
    Dim exp As String
    Dim ce As Range
    Set ce = ra(1, 1)
    exp = ce.Formula
    If ce.HasFormula Then
        Set ce = TimeChartRange(ra, "マスク信号(AND)を選択してください。")
        If ce Is Nothing Then Exit Sub
        Dim s As String
        s = ce.Address(False, False)
        If (MsgBox("負論理ですか。", vbYesNo, app_name) = vbYes) Then
            s = "(1-" & s & ")"
        End If
        ra.Formula = "=(" & Mid(exp, 2) & ")*" & s
    End If
End Sub

'MASK1
Private Sub MaskTimeChart1(ByVal ra As Range)
    Dim exp As String
    Dim ce As Range
    Set ce = ra(1, 1)
    exp = ce.Formula
    If ce.HasFormula Then
        Set ce = TimeChartRange(ra, "マスク信号(OR)選択してください。")
        If ce Is Nothing Then Exit Sub
        Dim s As String
        s = ce.Address(False, False)
        If (MsgBox("負論理ですか。", vbYesNo, app_name) <> vbYes) Then
            s = "(1-" & s & ")"
        End If
        ra.Formula = "=1-(1-" & Mid(exp, 2) & ")*" & s & ")"
    End If
End Sub

'----------------------------------------
'タイミングデータ生成
'----------------------------------------

Public Sub GenerateTimeChart(Optional mode As Long)
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Dim ra As Range
    Set ra = Selection
    If ra(1, 1) <> "" And ra(1, 1).Offset(, 1) <> "" Then
       Set ra = ra.Worksheet.Range(ra(1, 1), ra(1, 1).End(xlToRight))
    End If
    
    Select Case mode
    Case 1: Call GenerateTimeChart_CLK(ra)
    Case 2: Call GenerateTimeChart_CNT(ra)
    Case 3: Call GenerateTimeChart_LOGIC(ra)
    Case 5: Call GenerateTimeChart_SUBSET(ra)
    Case 6: Call GenerateTimeChart_JOIN(ra)
    
    Case 11: Call GenerateTimeChart_NOT(ra)
    Case 12: Call GenerateTimeChart_AND(ra)
    Case 13: Call GenerateTimeChart_OR(ra)
    Case 14: Call GenerateTimeChart_XOR(ra)
    Case 15: Call GenerateTimeChart_MUX(ra)
    Case 16: Call GenerateTimeChart_DFF(ra)
    Case 17: Call GenerateTimeChart_SRFF(ra)
    Case 18: Call GenerateTimeChart_SYNC(ra)
    Case 19: Call GenerateTimeChart_EDGE(ra)
    End Select
End Sub

'行取得
Private Function TimeChartRange(ra As Range, msg As String) As Range
    Dim c As Long, e As Long
    c = ra.Column
    e = c + ra.Columns.Count - 1
    On Error Resume Next
    Dim ce As Range
    Set ce = Application.InputBox(msg, app_name, Type:=8)
    On Error GoTo 0
    If ce Is Nothing Then Exit Function
    Set ce = ce.Worksheet.Cells(ce.Row, c)
    If ce(1, 1) <> "" And ce(1, 1).Offset(, 1) <> "" Then
       e = wsf.Max(e, ce(1, 1).End(xlToRight).Column)
    End If
    Set ra = ra.Worksheet.Range(ra(1, 1), ra(1, 1).Offset(, e - c))
    Set TimeChartRange = ce
End Function

'クロック
Private Sub GenerateTimeChart_CLK(ByVal ra As Range)
    If ra.Count < 2 Then
        Dim v As Variant
        v = Application.InputBox("横サイズを指定してください。", app_name, "16", , , , , Type:=1)
        If TypeName(v) <> "Boolean" Then Set ra = ra.Worksheet.Range(ra, ra.Offset(, v - 1))
    End If
    ra.FormulaR1C1 = "=1-RC[-1]"
End Sub

'カウンタ
Private Sub GenerateTimeChart_CNT(ByVal ra As Range)
    If ra.Count < 2 Then
        Dim v As Variant
        v = Application.InputBox("横サイズを指定してください。", app_name, "16", , , , , Type:=1)
        If TypeName(v) <> "Boolean" Then Set ra = ra.Worksheet.Range(ra, ra.Offset(, v - 1))
    End If
    
    Dim v1 As Variant, v2 As Variant, v3 As Variant
    v1 = Application.InputBox("カウント数を入力してください。", app_name, "16", Type:=0)
    If TypeName(v1) = "Boolean" Then Exit Sub
    If Left(v1, 1) = "=" Then v1 = Mid(v1, 2)
    v2 = Application.InputBox("ステップ数入力してください。", app_name, "1", Type:=0)
    If TypeName(v2) = "Boolean" Then Exit Sub
    If Left(v2, 1) = "=" Then v2 = Mid(v2, 2)
    v3 = Application.InputBox("開始値を入力してください。", app_name, "0", Type:=0)
    If TypeName(v3) = "Boolean" Then Exit Sub
    '
    ra.FormulaR1C1 = "=MOD(" & v2 & "+RC[-1]," & v1 & ")"
    ra.Cells(1, 1).FormulaR1C1 = v3
End Sub

'ロジック
Private Sub GenerateTimeChart_LOGIC(ByVal ra As Range)
    Dim ce As Range
    Set ce = TimeChartRange(ra, "信号を選択してください。")
    ra.Formula = "=1-" & ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
End Sub

'抽出
Private Sub GenerateTimeChart_SUBSET(ByVal ra As Range)
    Dim ce As Range
    Set ce = TimeChartRange(ra, "データ信号を選択してください。")
    If ce Is Nothing Then Exit Sub
    Dim s As String
    s = ra(ce.Row - ra.Row + 1, 1).Address(False, False)
    Dim v1 As Variant, v2 As Variant, v3 As Variant
    v1 = Application.InputBox("開始ビットを入力してください。", app_name, "0", Type:=0)
    If TypeName(v1) = "Boolean" Then Exit Sub
    If Left(v1, 1) = "=" Then v1 = Mid(v1, 2)
    v2 = Application.InputBox("ビット数を入力してください。", app_name, "1", Type:=0)
    If TypeName(v2) = "Boolean" Then Exit Sub
    If Left(v2, 1) = "=" Then v2 = Mid(v2, 2)
    s = "=mod(bitrshift(" & s & "," & v1 & "),2 ^ " & v2 & ")"
    ra.Formula = s
End Sub

'結合
Private Sub GenerateTimeChart_JOIN(ByVal ra As Range)
End Sub

'NOT
Private Sub GenerateTimeChart_NOT(ByVal ra As Range)
    Dim ce As Range
    Set ce = TimeChartRange(ra, "NOT:信号を選択してください。")
    If ce Is Nothing Then Exit Sub
    ra.Formula = "=1-" & ra(ce.Row - ra.Row + 1, 1).Address(False, False)
End Sub

'AND
Private Sub GenerateTimeChart_AND(ByVal ra As Range)
    Dim exp As String
    Dim ce As Range
    Dim i As Long
    For i = 1 To 8
        Set ce = TimeChartRange(ra, "AND:信号を選択してください。")
        If ce Is Nothing Then Exit For
        If exp <> "" Then exp = exp & "*"
        exp = exp & ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
    Next i
    ra.Formula = "=" & exp
End Sub

'OR
Private Sub GenerateTimeChart_OR(ByVal ra As Range)
    Dim exp As String
    Dim ce As Range
    Dim i As Long
    For i = 1 To 8
        Set ce = TimeChartRange(ra, "OR:信号を選択してください。")
        If ce Is Nothing Then Exit For
        If exp <> "" Then exp = exp & "*"
        exp = exp & "(1-" & ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False) & ")"
    Next i
    ra.Formula = "=1-" & exp
End Sub

'XOR
Private Sub GenerateTimeChart_XOR(ByVal ra As Range)
    Dim exp As String
    Dim ce As Range
    Dim i As Long
    For i = 1 To 8
        Set ce = TimeChartRange(ra, "XOR:信号を選択してください。")
        If ce Is Nothing Then Exit For
        If exp <> "" Then exp = exp & "+"
        exp = exp & ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
    Next i
    ra.Formula = "=mod(" & exp & ",2)"
End Sub

'MUX
Private Sub GenerateTimeChart_MUX(ByVal ra As Range)
    Dim exp As String, s As String
    Dim ce As Range
    Set ce = TimeChartRange(ra, "MUX:SEL信号を選択してください。")
    If ce Is Nothing Then Exit Sub
    exp = ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
    
    Dim i As Long
    For i = 0 To 10
        Set ce = TimeChartRange(ra, "MUX:信号A(SEL=" & i & ")を選択してください。")
        If ce Is Nothing Then Exit For
        s = ra(1, 1).Offset(ce.Row - ra.Row).Address(False, False)
        exp = exp & "," & s
    Next i
    ra.Formula = "=choose(1+" & exp & ")"
End Sub

'FF
Private Sub GenerateTimeChart_DFF(ByVal ra As Range)
    Dim exp As String
    Dim ce As Range
    Set ce = TimeChartRange(ra, "FF:クロック信号を選択してください。")
    If ce Is Nothing Then Exit Sub
    exp = ra(ce.Row - ra.Row + 1, 1).Address(False, False)
    
    Set ce = TimeChartRange(ra, "FF:イネーブル信号を選択してください。")
    If Not ce Is Nothing Then
        exp = exp & "*" & ra(ce.Row - ra.Row + 1, 0).Address(False, False)
    End If
    exp = exp & "," & ra(1, 1).Offset(0, -1).Address(False, False)
    
    Dim i As Long
    For i = 0 To 10
        Set ce = TimeChartRange(ra, "FF:データ信号を選択してください。")
        If ce Is Nothing Then Exit For
        exp = exp & "," & ra(ce.Row - ra.Row + 1, 0).Address(False, False)
    Next i
    ra.Formula = "=choose(1+" & exp & ")"
End Sub

'SRFF
Private Sub GenerateTimeChart_SRFF(ByVal ra As Range)
End Sub

'同期化
Private Sub GenerateTimeChart_SYNC(ByVal ra As Range)
    Dim exp As String
    Dim ce As Range
    Set ce = TimeChartRange(ra, "FF:クロック信号を選択してください。")
    If ce Is Nothing Then Exit Sub
    exp = ra(ce.Row - ra.Row + 1, 1).Address(False, False)
    exp = exp & "," & ra(1, 1).Offset(0, -1).Address(False, False)
    
    Set ce = TimeChartRange(ra, "FF:データ信号を選択してください。")
    If ce Is Nothing Then Exit Sub
    exp = exp & "," & ra(ce.Row - ra.Row + 1, 0).Address(False, False)
    ra.Formula = "=choose(1+" & exp & ")"
End Sub

'エッジ
Private Sub GenerateTimeChart_EDGE(ByVal ra As Range)
End Sub

