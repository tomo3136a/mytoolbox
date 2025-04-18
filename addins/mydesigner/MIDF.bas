Attribute VB_Name = "MIDF"
Option Explicit
Option Private Module

'----------------------------------------
'IDF作図機能
'----------------------------------------

Private Const FHDR As String = _
    "ファイル名,ファイルタイプ,仕様,作成ツール,作成日,版数," & _
    "名称,単位,オーナー,セクション," & _
    "形状,部品番号,高さ,長さ,配置,関連,状態," & _
    "ラベル,順番,X座標,Y座標,角度," & _
    "属性名,属性値"

'IDF field IDs
Private Enum FID
    'file information
    N_FILE_NAME = 1
    
    'file name
    N_FILE_TYPE     'file type:         BOARD_FILE, PANEL_FILE, LIBRARY_FILE
    N_IDF_VERSION   'IDF version:       1.0, 2.0, 3.0
    N_BUILDER       'source system      (string)
    N_BUILD_DATE    'build date:        format = yyyy/mm/dd.hh:mm:ss
    N_FILE_VERSION  'file version       (integer)
    
    N_NAME          'name
    N_UNITS         'units:             MM, THOU(=0.0254mm)
    N_OWNER         'owner:             ECAD, MCAD, UNOWNED
    N_SECTION       'section

    'section information
    N_GEOMETORY     'geometry name
                    'plating style:     PTH, NPTH
    N_NUMBER        'part number
                    'hole type:         PIN, VIA, MTG, TOOL, Other
    N_HEIGHT        'height/thickness/mounting offset
    N_LENGTH        'length/diameter
    N_LAYER          'board side/layer:  TOP, BOTTOM, BOTH, INNER, ALL
    N_REFERENCE     'reference
                    'association part:  BOARD, NOREFDES, PANEL
    N_STATUS        'status:            PLACED, UNPLACED, ECAD, MCAD

    'record data
    N_LABEL         'Loop Label:        0,1,番号,PROP
    N_INDEX         'index
    N_XPOS          'x point
    N_YPOS          'y point
    N_ANGLE         'angle:             0=line, 360=circle, othrs=arc

    'attribule
    N_ATTRIBUTE     'attribute name
                    'CAPACITANCE[uF], Resistance[ohm], Tolerance[%],
                    'POWER_OPR[mW], POWER_MAX[mW], THERM_COND[W/m・℃],
                    'THETA_JB[℃/W], THETA_JC[℃/W], Other
    N_VAL           'attribute value
    N_END
End Enum

'mode1 ids
Private Enum EM1
    N_HEADER = 1
    N_OUTLINE_KEEPOUT
    N_DRILLED_HOLES
    N_NOTES
    N_PLACEMENT
    N_MATERIAL
End Enum

'mode2 ids
Private Enum EM2
    N_BOARD_OUTLINE = 1 ' BOARD_OUTLINE, PANEL_OUTLINE
    N_OTHER_OUTLINE
    N_ROUTE_OUTLINE     ' ROUTE_OUTLINE, ROUTE_KEEPOUT
    N_PLACE_OUTLINE     ' PLACE_OUTLINE, PLACE_KEEPOUT
    N_VIA_KEEPOUT
    N_PLACE_REGION
End Enum

Private Type T_EnvIDF
    x0 As Double
    y0 As Double
    z0 As Double
    t0 As Double
    sc As Double
    scz As Double
    angle As Double
    flip As Boolean
    dir As Double
End Type

Private Enum EFLAG
    EF_SIDEA = 1
    EF_SIDEB
    EF_PLACE
    EF_ROUTE
    EF_PTH
    EF_NOTE
End Enum


'環境パラメータ
Private g_scale As Double           'スケール
Private g_flag(0 To 10) As Boolean  'モードフラグ
                                    ' 1.A面, 2.B面
                                    ' 3.配置制約, 4.配線制約
                                    ' 5.PTH, 6:Note

'----------------------------------------
'パラメータ制御
'----------------------------------------

'IDFパラメータ初期化
Sub IDF_ResetParam(Optional id As Integer)
    If id = 0 Or id = 1 Then g_scale = 1
    If id <> 0 And id <> 2 Then Exit Sub
    Dim i As Integer
    For i = 1 To UBound(g_flag)
        g_flag(i) = False
    Next i
    g_flag(1) = True
    g_flag(2) = True
End Sub

'IDFパラメータ設定
Sub IDF_SetParam(id As Integer, ByVal val As String)
    Select Case id
    Case 1
        If val <= 0 Then
            MsgBox "比率の設定が間違っています。(設定値>0)" & Chr(10) & "設定値： " & val
            Exit Sub
        End If
        g_scale = val
    End Select
End Sub

'IDFパラメータ取得
Function IDF_GetParam(id As Integer) As String
    Select Case id
    Case 1
        If g_scale <= 0 Then
            IDF_ResetParam id
            MsgBox "比率の設定を初期化しました。(設定値" & g_scale & ")"
        End If
        IDF_GetParam = g_scale
    End Select
End Function

'IDFフラグ設定
Sub IDF_SetFlag(id As Integer, Optional ByVal val As Boolean = True)
    g_flag(id) = val
End Sub

'IDFフラグチェック
Function IDF_IsFlag(id As Integer) As Boolean
    IDF_IsFlag = g_flag(id)
End Function


'----------------------------------------
'IDFマクロ
'----------------------------------------

'IDFマクロ
Public Sub MacroIDF()

End Sub

'----------------------------------------
'IDFシート操作
'  1. IDFシート作成 AddSheetIDF(Optional ByVal s_name)
'  2. IDFファイルを読み込み、シート作成 ImportIDF()
'
'----------------------------------------

'IDFシート作成
Public Sub AddSheetIDF(Optional ByVal s_name As String)

    'シート名入力(拡張子がない場合は.emnをつける)
    If s_name = "" Then
        s_name = InputBox("シート名を入力してください。", app_name)
    End If
    If s_name = "" Then Exit Sub
    If InStr(s_name, ".") = 0 Then s_name = s_name & ".emn"
    
    'ワークシート作成
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.name = UniqueSheetName(ws.Parent, s_name)
    ws.Range("C:C").NumberFormatLocal = "#0.0###"
    ws.Range("M:N").NumberFormatLocal = "#0.0###"
    ws.Range("W:Y").NumberFormatLocal = "#0.0###"
    
    ws.Range("G:L").NumberFormatLocal = "@"
    ws.Range("O:Q").NumberFormatLocal = "@"
    
    'ヘッダー行作成
    Dim vh As Variant
    vh = Split(FHDR, ",")
    ws.Cells(1, 1).Resize(1, 1 + UBound(vh)).Value = vh

End Sub

'IDFファイルを読み込み、シート作成
Public Sub ImportIDF()

    'IDFパラメータ取得
    Dim pth As String
    pth = GetRtParam("IDF", "path")
    If pth = "" Then pth = ActiveWorkbook.path
    
    'ファイル選択
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "IDF file", "*.emn,*.brd,*.bdf,*.idb,*.emp,*.lib,*.ldf,*.idl,*.pro"
        .Filters.Add "design file", "*.emn,*.brd,*.bdf,*.idb"
        .Filters.Add "library file", "*.emp,*.lib,*.ldf,*.idl,*.pro"
        .Filters.Add "全てのファイル", "*.*"
        .FilterIndex = 1
        .InitialFileName = pth & "\"
        .AllowMultiSelect = True
        If Not .Show Then Exit Sub
        
        '画面チラつき防止処置
        ScreenUpdateOff
        
        'ファイル読み込み
        Dim v As Variant
        For Each v In .SelectedItems
            i_ImportIDF CStr(v)
        Next v
        
        '画面チラつき防止処置解除
        ScreenUpdateOn
    End With

End Sub

Private Sub i_ImportIDF(ByVal pth As String)
    
    'ファイルを配列に読み込み
    Dim arr As Variant
    ReadArrayIDF arr, pth
    If UBound(arr, 2) < 2 Then Exit Sub
    
    'ワークシート作成
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.name = UniqueSheetName(ws.Parent, fso.GetFileName(pth))
    ws.Range("C:C").NumberFormatLocal = "#0.0###"
    ws.Range("M:N").NumberFormatLocal = "#0.0###"
    ws.Range("W:Y").NumberFormatLocal = "#0.0###"
    
    ws.Range("G:L").NumberFormatLocal = "@"
    ws.Range("O:Q").NumberFormatLocal = "@"
    
    'ワークシートにヘッダー行出力
    Dim vh As Variant
    vh = Split(FHDR, ",")
    ws.Cells(1, 1).Resize(1, 1 + UBound(vh)).Value = vh
    
    'ワークシートに出力
    ws.Cells(2, 1).Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr

End Sub

'----------------------------------------
'IDF読み込み
'----------------------------------------

'IDFファイルを読み込み、配列作成
Private Sub ReadArrayIDF(arr As Variant, pth As String)
    
    Dim col As Collection
    Set col = New Collection

    '読み込み
    ReadIDF col, pth
    Dim rcnt As Long, ccnt As Long
    rcnt = col.Count
    ccnt = UBound(col(1))
    If rcnt < 1 Then Exit Sub
    
    '配列にコピー
    ReDim arr(1 To rcnt, 1 To ccnt)
    Dim r As Long
    Dim c As Long
    For r = 1 To rcnt
        For c = 1 To ccnt - 1
            arr(r, c) = col(r)(c)
        Next c
    Next r
    Set col = Nothing

End Sub

'IDFファイルを読み込み、コレクション作成
Private Sub ReadIDF(col As Collection, pth As String)
    
    SetRtParam "IDF", "path", fso.GetParentFolderName(pth)
    If Not fso.FileExists(pth) Then Exit Sub
    Dim file_name As String
    file_name = fso.GetFileName(pth)
    
    'モードディクショナリ作成
    Dim mode As Dictionary
    Set mode = SectionMode

    'フィールドセパレート用正規表現
    Dim re As Object
    Set re = regex("(\""[^""]*\""|\S+)+")
    
    Const cmax As Integer = 25
    Dim wa As Variant
    ReDim wa(1 To cmax)
    
    Dim sect As String
    Dim owner As String
    Dim mode1 As Integer
    Dim mode2 As Integer
    Dim seq As Long
    Dim Index As Long
    
    'ファイルを読み込み、行ごとに処理
    Dim st As Object
    Set st = fso.GetFile(pth).OpenAsTextStream
    Do Until st.AtEndOfStream = True
        Dim mc As Object
        Set mc = re.Execute(st.Readline)
        If mc.Count > 0 Then
            Dim va As Variant
            Set va = mc
            Dim s As String
            s = UCase(va(0))
            
            '1文字目がピリオドならセクション処理
            If Left(s, 1) = "." Then
                sect = UCase(Mid(s, 2))
                If Left(sect, 4) <> "END_" Then
                    If va.Count > 1 Then owner = va(1)
                    mode1 = mode(sect) \ 10
                    mode2 = mode(sect) Mod 10
                    seq = 1
                Else
                    Dim c As Integer
                    For c = N_GEOMETORY To cmax
                        wa(c) = ""
                    Next c
                    sect = ""
                    owner = ""
                    mode1 = 0
                    mode2 = 0
                    seq = 0
                End If
            End If

            Select Case mode1
            Case EM1.N_HEADER
                Select Case seq
                Case 1
                    wa(FID.N_FILE_NAME) = file_name
                Case 2
                    wa(FID.N_FILE_TYPE) = UCase(s)
                    wa(FID.N_IDF_VERSION) = va(1)
                    s = Replace(va(2), """", "")
                    wa(FID.N_BUILDER) = s
                    wa(FID.N_BUILD_DATE) = va(3)
                    wa(FID.N_FILE_VERSION) = CInt(va(4))
                Case 3
                    wa(FID.N_NAME) = s
                    wa(FID.N_UNITS) = UCase(va(1))
                End Select
                seq = seq + 1
            
            Case EM1.N_OUTLINE_KEEPOUT
                Select Case seq
                Case 1
                    wa(FID.N_SECTION) = sect
                    wa(FID.N_OWNER) = owner
                    Index = 0
                    If mode2 = EM2.N_VIA_KEEPOUT Then seq = seq + 1
                Case 2
                    Select Case mode2
                    Case EM2.N_BOARD_OUTLINE
                        wa(FID.N_HEIGHT) = va(0)
                    Case EM2.N_OTHER_OUTLINE
                        wa(FID.N_NUMBER) = va(0)
                        wa(FID.N_HEIGHT) = va(1)
                        wa(FID.N_LAYER) = va(2)
                    Case EM2.N_ROUTE_OUTLINE
                        wa(FID.N_LAYER) = va(0)
                    Case EM2.N_PLACE_OUTLINE
                        wa(FID.N_LAYER) = va(0)
                        wa(FID.N_HEIGHT) = va(1)
                    Case EM2.N_PLACE_REGION
                        wa(FID.N_LAYER) = va(0)
                        wa(FID.N_REFERENCE) = va(1)
                    End Select
                Case Else
                    If wa(FID.N_LABEL) = s Then Index = Index + 1 Else Index = 0
                    wa(FID.N_LABEL) = s
                    wa(FID.N_INDEX) = Index
                    wa(FID.N_XPOS) = va(1)
                    wa(FID.N_YPOS) = va(2)
                    wa(FID.N_ANGLE) = va(3)
                    col.Add wa
                End Select
                seq = seq + 1
            
            Case EM1.N_DRILLED_HOLES
                Select Case seq
                Case 1
                    wa(FID.N_SECTION) = sect
                    wa(FID.N_OWNER) = ""
               Case Else
                    wa(FID.N_LENGTH) = s
                    wa(FID.N_XPOS) = va(1)
                    wa(FID.N_YPOS) = va(2)
                    wa(FID.N_GEOMETORY) = va(3)
                    wa(FID.N_REFERENCE) = va(4)
                    wa(FID.N_NUMBER) = va(5)
                    wa(FID.N_OWNER) = va(6)
                    col.Add wa
                End Select
                seq = seq + 1
            
            Case EM1.N_NOTES
                Select Case seq
                Case 1
                    wa(FID.N_SECTION) = sect
                    wa(FID.N_OWNER) = ""
                Case Else
                    wa(FID.N_XPOS) = va(0)
                    wa(FID.N_YPOS) = va(1)
                    wa(FID.N_HEIGHT) = va(2)
                    wa(FID.N_LENGTH) = va(3)
                    wa(FID.N_VAL) = va(4)
                    col.Add wa
                End Select
                seq = seq + 1
            
            Case EM1.N_PLACEMENT
                Select Case seq
                Case 1
                    wa(FID.N_SECTION) = sect
                    wa(FID.N_OWNER) = ""
                Case 2
                    wa(FID.N_GEOMETORY) = s
                    wa(FID.N_NUMBER) = va(1)
                    wa(FID.N_REFERENCE) = va(2)
                Case Else
                    wa(FID.N_XPOS) = va(0)
                    wa(FID.N_YPOS) = va(1)
                    If wa(FID.N_IDF_VERSION) = 2 Then
                        wa(FID.N_ANGLE) = va(2)
                        wa(FID.N_LAYER) = va(3)
                        wa(FID.N_STATUS) = va(4)
                    Else
                        wa(FID.N_HEIGHT) = va(2)
                        wa(FID.N_ANGLE) = va(3)
                        wa(FID.N_LAYER) = va(4)
                        wa(FID.N_STATUS) = va(5)
                    End If
                    col.Add wa
                    seq = seq - 2
                End Select
                seq = seq + 1
            
            Case EM1.N_MATERIAL
                Select Case seq
                Case 1
                    wa(FID.N_SECTION) = sect
                    wa(FID.N_OWNER) = ""
                    Index = 0
                Case 2
                    wa(FID.N_GEOMETORY) = s
                    wa(FID.N_NUMBER) = va(1)
                    wa(FID.N_UNITS) = va(2)
                    wa(FID.N_HEIGHT) = va(3)
                Case Else
                    If wa(FID.N_LABEL) = s Then Index = Index + 1 Else Index = 0
                    s = UCase(s)
                    wa(FID.N_LABEL) = s
                    If s = "PROP" Then
                        wa(FID.N_ATTRIBUTE) = va(1)
                        wa(FID.N_VAL) = va(2)
                        wa(FID.N_INDEX) = ""
                        wa(FID.N_XPOS) = ""
                        wa(FID.N_YPOS) = ""
                        wa(FID.N_ANGLE) = ""
                    Else
                        wa(FID.N_ATTRIBUTE) = ""
                        wa(FID.N_VAL) = ""
                        wa(FID.N_INDEX) = Index
                        wa(FID.N_XPOS) = va(1)
                        wa(FID.N_YPOS) = va(2)
                        wa(FID.N_ANGLE) = va(3)
                    End If
                    col.Add wa
                End Select
                seq = seq + 1
            End Select
        End If
    Loop
    st.Close
    Set st = Nothing

End Sub

'モードディクショナリ
Private Function SectionMode() As Dictionary
    Dim dic As Dictionary
    Set dic = New Dictionary
    dic.Add "HEADER", 11
    dic.Add "BOARD_OUTLINE", 21
    dic.Add "PANEL_OUTLINE", 21
    dic.Add "OTHER_OUTLINE", 22
    dic.Add "ROUTE_OUTLINE", 23
    dic.Add "ROUTE_KEEPOUT", 23
    dic.Add "PLACE_OUTLINE", 24
    dic.Add "PLACE_KEEPOUT", 24
    dic.Add "VIA_KEEPOUT", 25
    dic.Add "PLACE_REGION", 26
    dic.Add "DRILLED_HOLES", 31
    dic.Add "NOTES", 41
    dic.Add "PLACEMENT", 51
    dic.Add "ELECTRICAL", 61
    dic.Add "MECHANICAL", 62
    Set SectionMode = dic
End Function

'----------------------------------------
'IDF書き出し
'----------------------------------------

'シートからIDFファイル書き出し
Public Sub ExportIDF()

    Dim root As String
    root = GetRtParam("IDF", "path", ActiveWorkbook.path)

    Dim ws As Worksheet
    For Each ws In ActiveWindow.SelectedSheets
        
        '出力パスの選択
        Dim pth As String
        pth = fso.BuildPath(root, re_replace(ws.name, "\s*\(\d+\)$", ""))
        Dim flt As String
        flt = "IDF file,*.emn,IDF file,*.brd,IDF file,*.bdf,IDF file,*.idb"
        flt = flt & ",library file,*.emp,library file,*.lib"
        flt = flt & ",library file,*.ldf,library file,*.idl"
        flt = flt & ",library file,*pro"
        flt = flt & ",all file,*.*"
        Dim idx As Integer
        idx = 9
        If LCase(Right(pth, 4)) = ".emn" Then idx = 1
        If LCase(Right(pth, 4)) = ".brd" Then idx = 2
        If LCase(Right(pth, 4)) = ".bdf" Then idx = 3
        If LCase(Right(pth, 4)) = ".idb" Then idx = 4
        If LCase(Right(pth, 4)) = ".emp" Then idx = 5
        If LCase(Right(pth, 4)) = ".lib" Then idx = 6
        If LCase(Right(pth, 4)) = ".ldf" Then idx = 7
        If LCase(Right(pth, 4)) = ".idl" Then idx = 8
        If LCase(Right(pth, 4)) = ".pro" Then idx = 9
        pth = Application.GetSaveAsFilename(pth, flt, idx)
        If pth = "False" Then Exit Sub
        
        'ファイル書き出し
        Dim ra As Range
        Set ra = ws.UsedRange
        If ra.Count > 2 Then
            Dim arr As Variant
            arr = ra.Value
            WriteIDF pth, arr
        End If
    Next ws
    
    SetRtParam "IDF", "path", fso.GetParentFolderName(pth)

End Sub

'配列からIDFファイル書き出し
Private Sub WriteIDF(pth As String, arr As Variant)
    
    'モードディクショナリ作成
    Dim mode As Dictionary
    Set mode = SectionMode
    
    Dim r As Long
    Dim c As Long
    Dim n As Long
    Dim hdr As Integer
    Dim seq As Integer
    Dim mode1 As Integer
    Dim mode2 As Integer

    Dim sect As String
    Dim id As String
    Dim label As String
    Dim line As String
    Dim s0 As String
    Dim s1 As String
    Dim s As String
    
    Open pth For Output As #1
    For r = 1 To UBound(arr, 1)
        
        'ヘッダ情報書き出し
        If hdr = 0 Then
            line = Trim(arr(r, FID.N_FILE_TYPE))
            Select Case line
            Case "BOARD_FILE"
                hdr = 1
            Case "PANEL_FILE"
                hdr = 2
            Case "LIBRARY_FILE"
                hdr = 3
            End Select
            If hdr > 0 Then
                Print #1, RTrim(".HEADER")
                line = line & "  " & Format(arr(r, FID.N_IDF_VERSION), "0.0")
                line = line & "  """ & arr(r, FID.N_BUILDER) & """"
                line = line & "  " & arr(r, FID.N_BUILD_DATE)
                line = line & "  " & arr(r, FID.N_FILE_VERSION)
                Print #1, RTrim(line)
                If hdr <> 3 Then
                    line = LeftAlignedFormat(arr(r, FID.N_NAME), 15)
                    line = line & RightAlignedFormat(arr(r, FID.N_UNITS), 4)
                    Print #1, RTrim(line)
                End If
                Print #1, RTrim(".END_HEADER")
            End If
        End If
        
        If hdr > 0 Then
            'セクションキーワード取得
            s0 = Trim(arr(r, FID.N_SECTION))
            s1 = ""
            If sect <> s0 Then
                mode1 = mode(s0) \ 10
                mode2 = mode(s0) Mod 10
            End If
            Select Case mode1
            Case EM1.N_OUTLINE_KEEPOUT
                s1 = arr(r, FID.N_GEOMETORY) & "-"
                s1 = s1 & arr(r, FID.N_NUMBER) & "-"
                s1 = s1 & arr(r, FID.N_HEIGHT) & "-"
                s1 = s1 & arr(r, FID.N_LAYER) & "-"
                s1 = s1 & arr(r, FID.N_REFERENCE)
            Case EM1.N_MATERIAL
                s1 = arr(r, FID.N_GEOMETORY) & "-"
                s1 = s1 & arr(r, FID.N_NUMBER)
            End Select
            
            'セクションクローズ検出
            Dim flg As Boolean
            flg = False
            If mode1 = 0 Then
            ElseIf mode1 = EM1.N_OUTLINE_KEEPOUT And arr(r, FID.N_INDEX) = 0 And label = arr(r, FID.N_LABEL) Then
                flg = True
            ElseIf sect <> s0 Or id <> s1 Then
                flg = True
            End If
            
            'セクションクローズ処理
            If flg Then
                If sect <> "" Then Print #1, (".END_" & sect)
                sect = s0
                id = s1
                label = arr(r, FID.N_LABEL)
                line = "." & sect
                If mode1 = EM1.N_DRILLED_HOLES Then
                ElseIf mode1 = EM1.N_NOTES Then
                ElseIf mode1 = EM1.N_PLACEMENT Then
                ElseIf mode1 = EM1.N_MATERIAL Then
                Else
                    line = LeftAlignedFormat(line, 16) & "   "
                    line = line & RightAlignedFormat(arr(r, FID.N_OWNER), 8)
                End If
                Print #1, (line)
                seq = 1
            End If

            line = ""
            Select Case mode1
            Case EM1.N_OUTLINE_KEEPOUT
                If seq = 1 Then
                    Select Case mode2
                    Case EM2.N_BOARD_OUTLINE
                        line = Format(arr(r, FID.N_HEIGHT), "0.0")
                    Case EM2.N_OTHER_OUTLINE
                        line = arr(r, FID.N_GEOMETORY) & "  "
                        line = line & Format(arr(r, FID.N_HEIGHT), "0.0")
                        line = line & "  " & arr(r, FID.N_LAYER)
                    Case EM2.N_ROUTE_OUTLINE
                        line = arr(r, FID.N_LAYER)
                    Case EM2.N_PLACE_OUTLINE
                        line = LeftAlignedFormat(arr(r, FID.N_LAYER), 8) & "   "
                        line = line & RightAlignedFormat(arr(r, FID.N_HEIGHT), 8)
                    Case EM2.N_PLACE_REGION
                        line = arr(r, FID.N_LAYER)
                        line = line & "  " & arr(r, FID.N_REFERENCE)
                    End Select
                    If mode2 <> EM2.N_VIA_KEEPOUT Then Print #1, RTrim(line)
                    seq = 2
                End If
                line = arr(r, FID.N_LABEL) & "  "
                line = line & RightAlignedFormat(arr(r, FID.N_XPOS), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_YPOS), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_ANGLE), 8)
                Print #1, RTrim(line)
            
            Case EM1.N_DRILLED_HOLES
                line = RightAlignedFormat(arr(r, FID.N_LENGTH), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_XPOS), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_YPOS), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_GEOMETORY), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_REFERENCE), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_NUMBER), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_OWNER), 8)
                Print #1, RTrim(line)
            
            Case EM1.N_NOTES
                line = RightAlignedFormat(arr(r, FID.N_XPOS), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_YPOS), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_HEIGHT), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_LENGTH), 8)
                line = line & "  " & arr(r, FID.N_VAL)
                Print #1, RTrim(line)
            
            Case EM1.N_PLACEMENT
                line = RightAlignedFormat(arr(r, FID.N_GEOMETORY))
                line = line & RightAlignedFormat(arr(r, FID.N_NUMBER))
                line = line & RightAlignedFormat(arr(r, FID.N_REFERENCE))
                Print #1, RTrim(line)
                line = RightAlignedFormat(arr(r, FID.N_XPOS), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_YPOS), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_HEIGHT), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_ANGLE), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_LAYER), 8)
                line = line & RightAlignedFormat(arr(r, FID.N_STATUS), 8)
                Print #1, RTrim(line)
            
            Case EM1.N_MATERIAL
                If seq = 1 Then
                    line = RightAlignedFormat(arr(r, FID.N_GEOMETORY))
                    line = line & RightAlignedFormat(arr(r, FID.N_NUMBER))
                    line = line & RightAlignedFormat(arr(r, FID.N_UNITS), 8)
                    line = line & RightAlignedFormat(arr(r, FID.N_HEIGHT), 8)
                    Print #1, RTrim(line)
                    seq = 2
                End If
                line = arr(r, FID.N_LABEL)
                If line = "PROP" Then
                    line = line & RightAlignedFormat(arr(r, FID.N_ATTRIBUTE), 12)
                    line = line & RightAlignedFormat(arr(r, FID.N_VAL), 8)
                Else
                    line = line & "  " & RightAlignedFormat(arr(r, FID.N_XPOS), 8)
                    line = line & RightAlignedFormat(arr(r, FID.N_YPOS), 8)
                    line = line & RightAlignedFormat(arr(r, FID.N_ANGLE), 8)
                End If
                Print #1, RTrim(line)
            
            End Select
        End If
    Next r
    If sect <> "" Then Print #1, (".END_" & sect)
    Close #1

End Sub

'-------------------------------------
'IDFレコード追加
'-------------------------------------

Public Sub AddRecordIDF(Optional mode As Long)
    
    Select Case mode
    Case 0: IDF_PartForm.Show       'ライブラリ部品追加
    Case 1: IDF_PlaceForm.Show      '配置指定追加
    Case 2: IDF_PanelForm.Show      'アウトライン追加
    End Select
    Unload IDF_PartForm
    Unload IDF_PlaceForm
    Unload IDF_PanelForm

End Sub

'-------------------------------------
'IDF描画
'-------------------------------------

'IDF描画
Public Sub DrawIDF()
    
    IDF_ModeForm.Show
    If Not IDF_IsFlag(0) Then Exit Sub

    With ActiveCell
        i1_DrawIDF .Worksheet, .Left, .Top, .Row
    End With

End Sub

'ワークシートからIDF描画
Private Sub i1_DrawIDF(ws As Worksheet, x As Double, ByVal y As Double, r As Long)
    
    'ライブラリ辞書作成
    Dim lib As Dictionary
    Set lib = New Dictionary
    
    'デザイン取得
    Dim name As String
    'LoadDesign_X name, lib, 1
    Dim ws2 As Worksheet
    SelectDesign ws2, 2
    If Not ws2 Is Nothing Then LoadDesign lib, name, ws2
    SelectDesign ws2, 1
    If ws2 Is Nothing Then Exit Sub
    LoadDesign lib, name, ws2
    If Not lib.Exists(name) Then Exit Sub
    If Not lib.Exists("$" & name) Then Exit Sub
    
    'データ配列取得
    Dim arr As Variant
    arr = lib("$" & name)
    
    '描画領域計算
    Dim dra As Variant
    CalcDrawingArea dra, arr
    
    'スケール・原点計算
    Dim ce As Range
    Dim sc As Double, x0 As Double, y0 As Double
    sc = g_scale
    Set ce = ws.Cells(r, 1)
    x0 = x + sc * 20
    y0 = y + sc * (dra(5) + 20)
    Do While y0 > ce.Top
        Set ce = ce.Offset(1)
    Loop
    y0 = ce.Top - sc * 20
    
    '環境変数作成
    Dim env As T_EnvIDF
    SetEnvIDF env, x0, y0, 0, sc
    
    Dim sh As Shape, sh2 As Shape
    Dim ns As Collection
    Set ns = New Collection
    
    '原点マーク
    Set sh = DrawOrigin(ws, env, arr, lib)
    If Not sh Is Nothing Then ns.Add sh.name
        
    'Assy描画
    Set sh = DrawAssy(ws, env, lib, name, ns)

    If ns.Count > 0 Then Set sh = GroupShape(ws, ns, name)
    
    Set ns = Nothing
    Set lib = Nothing

End Sub

'-------------------------------------
'デザイン読み込み
'-------------------------------------

'ライブラリシート選択
Private Sub SelectDesign(ws As Worksheet, Optional mode As Long)
    
    Dim ptn As String
    Select Case mode
    Case 1: ptn = "\.(emn|brd|bdf|idb)($|\s)"
    Case 2: ptn = "\.(emp|lib|ldf|idl|pro)($|\s)"
    Case 3: ptn = "\.(emn|brd|bdf|idb|emp|lib|ldf|idl|pro)($|\s)"
    End Select
    
    Set ws = SelectSheet(ActiveWorkbook, ptn, "デザイン一覧：", app_name)
    If ws Is Nothing Then
    ElseIf ws Is ActiveSheet Then Set ws = Nothing
    ElseIf ws.UsedRange.Rows.Count < 1 Then Set ws = Nothing
    ElseIf ws.UsedRange.Columns.Count < FID.N_VAL Then Set ws = Nothing
    End If

End Sub

'ライブラリを辞書に登録
Private Sub LoadDesign(lib As Dictionary, name As String, ws As Worksheet)
    
    name = UCase(ws.name)
    
    'ライブラリデータ配列取得
    Dim ra As Range
    Set ra = ws.UsedRange.Offset(1)
    Set ra = ra.Resize(ra.Rows.Count - 1)
    Dim arr As Variant
    arr = ra.Value
    Set ra = Nothing
    
    ConvertToMM arr
    
    'ライブラリ名取得
    Dim rmax As Long, r As Long
    rmax = UBound(arr, 1)
    If rmax > 10 Then rmax = 10
    For r = 1 To rmax
        If "" <> arr(r, FID.N_NAME) Then
            name = UCase(arr(r, FID.N_NAME))
            Exit For
        End If
    Next r
    
    'データ配列を解析し辞書に追加
    If lib.Exists("$" & name) Then Exit Sub
    lib.Add "$" & name, arr
        
    Dim kw As String
    If r > rmax Then
        'library
        For r = 1 To UBound(arr, 1)
            If "" & arr(r, FID.N_INDEX) = "0" Then
                kw = UCase(arr(r, FID.N_GEOMETORY))
                If Not lib.Exists(kw) Then lib.Add kw, Array(name, r)
            End If
        Next r
    Else
        'board
        Dim dic As Dictionary
        Set dic = New Dictionary
        For r = 1 To UBound(arr, 1)
            kw = "" & arr(r, FID.N_INDEX)
            If kw = "" Or kw = "0" Then
                kw = arr(r, FID.N_SECTION)
                kw = Join(Array(kw, arr(r, FID.N_LAYER), arr(r, FID.N_REFERENCE)), "-")
                kw = UCase(kw)
                Dim col As Collection
                If dic.Exists(kw) Then
                    Set col = dic(kw)
                Else
                    Set col = New Collection
                    dic.Add kw, col
                End If
                col.Add r
            End If
        Next r
        If Not lib.Exists(name) Then lib.Add name, dic
    End If
    
End Sub

'-------------------------------------
'描画配列がinchならmmに変換

Private Sub ConvertToMM(arr As Variant)

    Dim r As Long, c As Long
    For r = 1 To UBound(arr, 1)
        If arr(r, FID.N_UNITS) = "THOU" Then
            arr(r, FID.N_UNITS) = "MM"
            c = FID.N_HEIGHT
            If TypeName(arr(r, c)) = "Double" Then
                arr(r, c) = arr(r, c) * 0.0254
            End If
            c = FID.N_LENGTH
            If TypeName(arr(r, c)) = "Double" Then
                arr(r, c) = arr(r, c) * 0.0254
            End If
            c = FID.N_XPOS
            If TypeName(arr(r, c)) = "Double" Then
                arr(r, c) = arr(r, c) * 0.0254
            End If
            c = FID.N_YPOS
            If TypeName(arr(r, c)) = "Double" Then
                arr(r, c) = arr(r, c) * 0.0254
            End If
        End If
    Next r

End Sub

'-------------------------------------

'描画配列から全体範囲を取得
Private Sub CalcDrawingArea(da As Variant, arr As Variant)
    
    Dim r As Long
    Dim x As Double, y As Double
    r = 2
    x = arr(r, FID.N_XPOS)
    y = arr(r, FID.N_YPOS)
    
    Dim xs As Double, ys As Double, xe As Double, ye As Double
    xs = xe = x
    ys = ye = y
    For r = r + 1 To UBound(arr, 1)
        x = arr(r, FID.N_XPOS)
        y = arr(r, FID.N_YPOS)
        If x < xs Then xs = x
        If x > xe Then xe = x
        If y < ys Then ys = y
        If y > ye Then ye = y
    Next r

    da = Array(xs, ys, xe, ye, xe - xs, ye - ys)

End Sub

'-------------------------------------

'描画パラメータ設定
Private Sub SetEnvIDF( _
    env As T_EnvIDF, _
    Optional x0 As Double, Optional y0 As Double, Optional z0 As Double, _
    Optional sc As Double = 1, Optional scz As Double = 1, _
    Optional angle As Double, _
    Optional flip As Boolean, Optional dir As Double = 1)
    env.x0 = x0
    env.y0 = y0
    env.z0 = z0
    env.sc = sc
    env.scz = scz
    env.angle = angle
    env.flip = flip
    env.dir = dir
End Sub
    
'-------------------------------------
'描画
'-------------------------------------

'Assy描画
Private Function DrawAssy( _
        ws As Worksheet, env As T_EnvIDF, _
        lib As Dictionary, name As String, _
        Optional col As Collection) As Shape
        
    Dim ns As Collection
    Set ns = col
    If ns Is Nothing Then Set ns = New Collection
    
    Dim arr As Variant
    arr = lib("$" & name)
    If Not lib.Exists(name) Then Exit Function
    
    Dim dic As Object
    Set dic = lib(name)
    GetTinkness env, dic, arr
    
    Dim sh As Shape
    Dim ns2 As Collection
    Dim side As Variant
    Dim k As String

    'OUTLINE, KEEPOUT, REGION(BOTTOM)
    If g_flag(2) Then
        env.scz = -env.scz
        env.z0 = env.z0 + env.scz * env.t0
        
        For Each side In Array("ALL", "BOTH", "BOTTOM")
            Set ns2 = New Collection
            If g_flag(4) Then
                k = Join(Array("ROUTE_OUTLINE", side, ""), "-")
                Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
                If Not sh Is Nothing Then ns2.Add sh.name
                
                k = Join(Array("ROUTE_KEEPOUT", side, ""), "-")
                Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
                If Not sh Is Nothing Then ns2.Add sh.name
            End If
                
            If g_flag(3) Then
                k = Join(Array("PLACE_OUTLINE", side, ""), "-")
                Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
                If Not sh Is Nothing Then ns2.Add sh.name
                
                k = Join(Array("PLACE_KEEPOUT", side, ""), "-")
                Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
                If Not sh Is Nothing Then ns2.Add sh.name
                
                k = Join(Array("PLACE_REGION", side, ""), "-")
                Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
                If Not sh Is Nothing Then ns2.Add sh.name
            End If
                
            Set sh = Nothing
            If ns2.Count > 0 Then Set sh = GroupShape(ws, ns2, name & "_BOTTOM")
            If Not sh Is Nothing Then ns.Add sh.name
            Set ns2 = Nothing
        Next side
        
        env.z0 = env.z0 - env.scz * env.t0
        env.scz = -env.scz
    End If

    'PLACEMENT(BOTTOM)
    If g_flag(2) Then
        env.scz = -env.scz
        env.z0 = env.z0 + env.scz * env.t0
        env.flip = Not env.flip
        
        k = Join(Array("PLACEMENT", "BOTTOM", ""), "-")
        Set sh = DrawGroupPlace(ws, env, k, arr, lib(name), lib)
        If Not sh Is Nothing Then ns.Add sh.name
        
        k = Join(Array("OTHER_OUTLINE", "BOTTOM", ""), "-")
        Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
        If Not sh Is Nothing Then ns.Add sh.name
        
        env.flip = Not env.flip
        env.z0 = env.z0 - env.scz * env.t0
        env.scz = -env.scz
    End If
    
    'Board
    Set sh = DrawBoard(ws, env, lib, name)
    If Not sh Is Nothing Then ns.Add sh.name
    
    'VIA
    If g_flag(5) Then
        k = Join(Array("VIA_KEEPOUT", "", ""), "-")
        Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
        If Not sh Is Nothing Then ns.Add sh.name
    End If

    'PLACEMENT(TOP)
    If g_flag(1) Then
        k = Join(Array("PLACEMENT", "TOP", ""), "-")
        Set sh = DrawGroupPlace(ws, env, k, arr, lib(name), lib)
        If Not sh Is Nothing Then ns.Add sh.name
        
        k = Join(Array("OTHER_OUTLINE", "TOP", ""), "-")
        Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
        If Not sh Is Nothing Then ns.Add sh.name
    End If
    
    'OUTLINE, KEEPOUT, REGION(TOP)
    If g_flag(1) Then
        For Each side In Array("ALL", "BOTH", "INNER", "TOP")
            Set ns2 = New Collection
            If g_flag(4) Then
                k = Join(Array("ROUTE_OUTLINE", side, ""), "-")
                Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
                If Not sh Is Nothing Then ns2.Add sh.name
                
                k = Join(Array("ROUTE_KEEPOUT", side, ""), "-")
                Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
                If Not sh Is Nothing Then ns2.Add sh.name
            End If
            
            If g_flag(3) Then
                k = Join(Array("PLACE_OUTLINE", side, ""), "-")
                Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
                If Not sh Is Nothing Then ns2.Add sh.name
                
                k = Join(Array("PLACE_KEEPOUT", side, ""), "-")
                Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
                If Not sh Is Nothing Then ns2.Add sh.name
                
                k = Join(Array("PLACE_REGION", side, ""), "-")
                Set sh = DrawGroupOutline(ws, env, k, arr, lib(name))
                If Not sh Is Nothing Then ns2.Add sh.name
            End If
            
            Set sh = Nothing
            If ns2.Count > 0 Then Set sh = GroupShape(ws, ns2, name & "_TOP")
            If Not sh Is Nothing Then ns.Add sh.name
            Set ns2 = Nothing
        Next side
    End If
    
    'NOTES
    If g_flag(6) Then
        k = Join(Array("NOTES", "", ""), "-")
        Set sh = DrawGroupNote(ws, env, k, arr, lib(name))
        If Not sh Is Nothing Then ns.Add sh.name
    End If
    
    If col Is Nothing Then
        If ns.Count > 0 Then Set sh = GroupShape(ws, ns, name)
        Set ns = Nothing
        Set DrawAssy = sh
    End If
    
End Function

Private Sub GetTinkness(env As T_EnvIDF, dic As Dictionary, arr As Variant)

    env.t0 = 0
    Dim k As Variant, v As Variant
    For Each k In Array("BOARD_OUTLINE", "PANEL_OUTLINE")
        k = Join(Array(k, "", ""), "-")
        If dic.Exists(k) Then
            For Each v In dic(k)
                env.t0 = arr(CLng(v), FID.N_HEIGHT)
                Exit Sub
            Next v
        End If
    Next k

End Sub

'-------------------------------------

'原点描画
Private Function DrawOrigin( _
        ws As Worksheet, env As T_EnvIDF, _
        arr As Variant, dic As Dictionary) As Shape
    
    Dim sc As Double, w As Double
    sc = env.sc
    w = 20
    
    Dim tw As Double, tx As Double, ty As Double
    tw = sc * w / 2
    tx = env.x0 - sc * w
    ty = env.y0 + sc * w - tw
    
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeRightTriangle, tx, ty, tw, tw)
    With sh
        .LockAspectRatio = msoTrue
        .Placement = xlMove
        .name = "ORIGIN " & .id
    End With
    
    sh.Title = "ORIGIN"
    Dim s As String
    s = Join(Array("d:" & w & "," & w & ",0"), Chr(10))
    s = Join(Array(s, "sc:" & sc), Chr(10))
    sh.AlternativeText = s
    
    Set DrawOrigin = sh
    Set sh = Nothing

End Function

'-------------------------------------
'グループ描画
'-------------------------------------

'ボード描画
Private Function DrawBoard( _
        ws As Worksheet, env As T_EnvIDF, _
        lib As Dictionary, name As String) As Shape
        
    Dim arr As Variant
    arr = lib("$" & name)
    If Not lib.Exists(name) Then Exit Function
    
    Dim dic As Object
    Set dic = lib(name)
    
    Dim ns As Collection, ns2 As Collection
    Set ns = New Collection
    Dim sh As Shape
    Dim k As Variant, v As Variant
    
    'Board/Panel outline
    env.z0 = env.z0 - env.scz * env.t0
    For Each k In Array("BOARD_OUTLINE", "PANEL_OUTLINE")
        Set ns2 = New Collection
        Set sh = Nothing
        k = Join(Array(k, "", ""), "-")
        If dic.Exists(k) Then
            For Each v In dic(k)
                Set sh = DrawOutline(ws, env, arr, CLng(v))
                If Not sh Is Nothing Then
                    SetStyleIDF sh, CStr(k)
                    ns2.Add sh.name
                End If
            Next v
        End If
        If ns2.Count > 1 Then Set sh = GroupShape(ws, ns2, CStr(k))
        If Not sh Is Nothing Then ns.Add sh.name
        Set ns2 = Nothing
    Next k
    env.z0 = env.z0 + env.scz * env.t0
    
    'Hole(NPTH)
    Set ns2 = New Collection
    Set sh = Nothing
    For Each k In dic.Keys
        If k Like ("DRILLED_HOLES*") Then
            For Each v In dic(k)
                If arr(CLng(v), FID.N_GEOMETORY) <> "PTH" Then
                    Set sh = DrawHole(ws, env, arr, CLng(v))
                    If Not sh Is Nothing Then
                        SetStyleIDF sh, CStr(k)
                        ns2.Add sh.name
                    End If
                End If
            Next v
        End If
    Next k
    If ns2.Count > 1 Then Set sh = GroupShape(ws, ns2, "DRILLED_HOLES_NPTH")
    If Not sh Is Nothing Then ns.Add sh.name
    Set ns2 = Nothing
    
    'Hole(PTH)
    If g_flag(5) Then
        Set ns2 = New Collection
        Set sh = Nothing
        For Each k In dic.Keys
            If k Like ("DRILLED_HOLES*") Then
                For Each v In dic(k)
                    If arr(CLng(v), FID.N_GEOMETORY) = "PTH" Then
                        Set sh = DrawHole(ws, env, arr, CLng(v))
                        If Not sh Is Nothing Then
                            SetStyleIDF sh, CStr(k)
                            ns2.Add sh.name
                        End If
                    End If
                Next v
            End If
        Next k
        If ns2.Count > 1 Then Set sh = GroupShape(ws, ns2, "DRILLED_HOLES_PTH")
        If Not sh Is Nothing Then ns.Add sh.name
        Set ns2 = Nothing
    End If
    
    If ns.Count > 0 Then Set sh = GroupShape(ws, ns, name)
    Set ns = Nothing
    Set DrawBoard = sh

End Function

'OUTLINE/CUTOUT描画グループ化
Private Function DrawGroupOutline( _
        ws As Worksheet, env As T_EnvIDF, grp As String, _
        arr As Variant, dic As Dictionary) As Shape
    
    If Not dic.Exists(grp) Then Exit Function
    
    Dim ns As Collection
    Set ns = New Collection
    
    Dim col As Collection
    Dim v As Variant
    Set col = dic(grp)
    Dim sh As Shape
    
    Dim i As Long
    For i = 1 To col.Count
        Set sh = DrawOutline(ws, env, arr, CLng(col(i)))
        ns.Add sh.name
    Next i
    If ns.Count > 1 Then Set sh = GroupShape(ws, ns, grp)
    If sh Is Nothing Then Exit Function
    
    SetStyleIDF sh, grp
    
    Set DrawGroupOutline = sh

End Function

'ノート描画グループ化
Private Function DrawGroupNote( _
        ws As Worksheet, env As T_EnvIDF, grp As String, _
        arr As Variant, dic As Dictionary) As Shape
        
    Dim ns As Collection
    Set ns = New Collection
    
    Dim col As Collection
    Dim sh As Shape
    
    Dim kw As Variant
    For Each kw In dic.Keys
        Dim k As String
        k = kw
        If k Like (grp & "*") Then
            Set col = dic(k)
            Dim ns2 As Collection
            Set ns2 = New Collection
            Dim i As Long
            For i = 1 To col.Count
                Dim r As Long
                r = col(i)
                Set sh = DrawNote(ws, env, arr, r)
                If Not sh Is Nothing Then ns2.Add sh.name
            Next i
            If ns2.Count > 1 Then Set sh = GroupShape(ws, ns2, k)
            Set ns2 = Nothing
            If Not sh Is Nothing Then ns.Add sh.name
        End If
    Next kw
    If ns.Count > 1 Then Set sh = GroupShape(ws, ns, grp)
    Set ns = Nothing
    Set DrawGroupNote = sh

End Function

'配置グループ
Private Function DrawGroupPlace( _
        ws As Worksheet, env As T_EnvIDF, grp As String, _
        arr As Variant, dic As Dictionary, lib As Dictionary) As Shape
        
    Dim ns As Collection
    Set ns = New Collection
    
    Dim col As Collection
    Dim sh As Shape
    
    Dim kw As Variant
     For Each kw In dic.Keys
        Dim k As String
        k = kw
        If k Like (grp & "*") Then
            Set col = dic(k)
            Dim ns2 As Collection
            Set ns2 = New Collection
            Dim i As Long
            For i = 1 To col.Count
                Dim r As Long
                r = col(i)
                Set sh = DrawPlace(ws, env, k, arr, r, lib)
                If Not sh Is Nothing Then ns2.Add sh.name
            Next i
            If ns2.Count > 1 Then Set sh = GroupShape(ws, ns2, k)
            Set ns2 = Nothing
            If Not sh Is Nothing Then ns.Add sh.name
        End If
    Next kw
    If ns.Count > 1 Then Set sh = GroupShape(ws, ns, grp)
    Set ns = Nothing
    Set DrawGroupPlace = sh

End Function

'図形名のコレクションから図形グループ化
Private Function GroupShape(ws As Worksheet, ns As Collection, sect As String) As Shape
    Dim sh As Shape
    If ns.Count > 1 Then
        Set sh = ws.Shapes.Range(ColToArr(ns)).Group
        sh.LockAspectRatio = msoTrue
        sh.Placement = xlMove
        Dim s As String
        s = sect & " " & sh.id
        sh.name = s
    Else
        Set sh = ws.Shapes(ns(1))
    End If
    Set GroupShape = sh
    Set sh = Nothing
    Set ns = Nothing
End Function

'----------------------------------------
'IDF図形描画
'----------------------------------------

'OUTLINE/CUTOUT描画
Private Function DrawOutline( _
        ws As Worksheet, env As T_EnvIDF, _
        arr As Variant, r As Long) As Shape
    
    Dim kw As String
    kw = arr(r, FID.N_SECTION)
    
    Dim sh As Shape
    Set sh = DrawShape(ws, env, arr, r, 0, 0)
    If sh Is Nothing Then Exit Function
    sh.name = kw & " " & sh.id
    
    SetStyleIDF sh, kw, CStr(arr(r, FID.N_LABEL))
    Set DrawOutline = sh
    
    Set sh = Nothing
    
End Function

'ホール描画
Private Function DrawHole( _
        ws As Worksheet, env As T_EnvIDF, _
        arr As Variant, r As Long) As Shape
    
    Dim sc As Double, x0 As Double, y0 As Double, z0 As Double
    sc = env.sc
    x0 = env.x0
    y0 = env.y0
    z0 = env.z0

    Dim tw As Double, tx As Double, ty As Double
    tw = arr(r, FID.N_LENGTH)
    tx = arr(r, FID.N_LENGTH) / 2 * env.dir
    ty = arr(r, FID.N_LENGTH) / 2
    
    Dim pw As Double, px As Double, py As Double
    Dim kw As String
    pw = sc * tw
    px = x0 + sc * (arr(r, FID.N_XPOS) - tx) * env.dir
    py = y0 - sc * (arr(r, FID.N_YPOS) + ty)
    kw = arr(r, FID.N_GEOMETORY) & "-" & arr(r, FID.N_NUMBER)
    
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeOval, px, py, pw, pw)
    With sh
        .LockAspectRatio = msoTrue
        .Placement = xlMove
        .Title = kw
        .name = kw & " " & .id
    End With
    
    sh.Title = kw
    Dim s As String
    s = Join(Array("d:" & tw & "," & tw & ",0"), Chr(10))
    s = Join(Array(s, "sc:" & sc), Chr(10))
    sh.AlternativeText = s
    
    With sh.ThreeD
        .BevelTopType = msoBevelAngle
        .BevelTopInset = 0
        .BevelTopDepth = 0
        .BevelBottomType = msoBevelAngle
        .BevelBottomInset = 0
        .BevelBottomDepth = 0
        .Depth = 0
        .Z = sc * z0
    End With
    
    kw = arr(r, FID.N_SECTION)
    SetStyleIDF sh, kw
    
    Set DrawHole = sh
    Set sh = Nothing

End Function

'ノート描画
Public Function DrawNote( _
        ws As Worksheet, env As T_EnvIDF, _
        arr As Variant, r As Long) As Shape
    
    Dim sc As Double, x0 As Double, y0 As Double
    sc = env.sc
    x0 = env.x0
    y0 = env.y0
    
    Dim tw As Double, th As Double, tx As Double, ty As Double
    Dim kw As String
    tw = arr(r, FID.N_LENGTH)
    th = arr(r, FID.N_HEIGHT) * 2
    tx = 0
    ty = th
    kw = "NOTE"
    
    Dim pw As Double, ph As Double, px As Double, py As Double
    pw = sc * tw
    ph = sc * th
    px = x0 + sc * (arr(r, FID.N_XPOS) - tx) * env.dir
    py = y0 - sc * (arr(r, FID.N_YPOS) + ty)
    
    Dim sh As Shape
    Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, px, py, pw, ph)
    With sh
        .LockAspectRatio = msoTrue
        .Placement = xlMove
        .TextFrame.Characters.Font.Size = ph
        .TextFrame.Characters.text = arr(r, FID.N_VAL)
        With .TextFrame2
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorBottom
            .HorizontalAnchor = msoAnchorNone
            .WordWrap = msoFalse
        End With
        .TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
        .Title = kw
        .name = kw & " " & .id
    End With
    
    kw = arr(r, FID.N_SECTION)
    SetStyleIDF sh, kw
   
    With sh.ThreeD
        .BevelTopType = msoBevelAngle
        .BevelTopInset = 0
        .BevelTopDepth = 0
        .BevelBottomType = msoBevelAngle
        .BevelBottomInset = 0
        .BevelBottomDepth = 0
        .ExtrusionColor.RGB = RGB(255, 0, 0)
        .Depth = 0
        .Z = 20
    End With
    
    Set DrawNote = sh
    Set sh = Nothing
End Function

'部品配置描画
Private Function DrawPlace( _
        ws As Worksheet, env As T_EnvIDF, grp As String, _
        arr As Variant, r As Long, _
        lib As Dictionary) As Shape
    
    Dim sc As Double, a0 As Double, x0 As Double, y0 As Double
    a0 = env.angle
    sc = env.sc
    x0 = env.x0
    y0 = env.y0
    
    Dim sect As String
    Dim a1 As Double, x1 As Double, y1 As Double
    sect = arr(r, FID.N_SECTION)
    a1 = wsf.Pi * arr(r, FID.N_ANGLE) / 180
    x1 = arr(r, FID.N_XPOS)
    y1 = arr(r, FID.N_YPOS)
    
    Dim ta As Double, tx As Double, ty As Double
    ta = a0 + a1
    tx = x0 + sc * x1
    ty = y0 - sc * y1
    
    Dim tenv As T_EnvIDF
    SetEnvIDF tenv, tx, ty, 0, sc, env.scz, ta, env.flip, env.dir
    tenv.t0 = env.t0
    
    tenv.z0 = env.z0
    tenv.scz = env.scz
    tenv.flip = env.flip
    
    Dim kw As String
    kw = arr(r, FID.N_GEOMETORY)
    kw = UCase(kw)
    
    Dim name As String
    If Not lib.Exists(kw) Then
        MsgBox kw, vbOKOnly, "DrawPlace"
        'デザイン取得
        Dim ws2 As Worksheet
        SelectDesign ws2, 3
        If ws2 Is Nothing Then Exit Function
        LoadDesign lib, name, ws2
    End If
    
    Dim sh As Shape
    If TypeName(lib(kw)) = "Dictionary" Then
        If Not lib.Exists("$" & kw) Then Exit Function
        Dim arr2 As Variant
        arr2 = lib("$" & kw)
        tenv.dir = tenv.scz
        Set sh = DrawAssy(ws, tenv, lib, kw)
    Else
        If tenv.dir < 0 Then
            tx = x0 - sc * x1
            tenv.x0 = tx
        End If
        Set sh = DrawPart(ws, tenv, x1, -y1, arr, r, lib)
    End If
    If sh Is Nothing Then Exit Function
    sh.name = arr(r, FID.N_REFERENCE) & " " & sh.id
    Set DrawPlace = sh
    Set sh = Nothing

End Function

'----------------------------------------

Private Function DrawPart( _
        ws As Worksheet, env As T_EnvIDF, _
        x As Double, y As Double, _
        arr As Variant, r As Long, _
        lib As Dictionary) As Shape
    
    Dim kw As String
    kw = arr(r, FID.N_GEOMETORY)
    kw = UCase(kw)
    If Not lib.Exists(kw) Then
        MsgBox kw, vbOKOnly, "DrawPart"
        'デザイン取得
        Dim s As String
        Dim ws2 As Worksheet
        SelectDesign ws2, 3
        If ws2 Is Nothing Then
            r = UBound(arr, 1)
            Exit Function
        End If
        LoadDesign lib, s, ws2
    End If
    If Not lib.Exists(kw) Then Exit Function

    Dim v As Variant
    v = lib(kw)
    If TypeName(v) = "Empty" Then Exit Function
    
    Dim arr2 As Variant
    arr2 = lib("$" & lib(kw)(0))
    If TypeName(arr2) = "Empty" Then Exit Function
    Dim r2 As Long
    r2 = lib(kw)(1)
    
    Dim sh As Shape
    Set sh = DrawShape(ws, env, arr2, r2, 0, 0)
    If sh Is Nothing Then Exit Function
    sh.Title = kw
    sh.name = kw & " " & sh.id
    
    kw = arr(r, FID.N_SECTION)
    SetStyleIDF sh, kw
    
    Set DrawPart = sh
    Set sh = Nothing
    
End Function

'----------------------------------------
'IDF図形描画
'----------------------------------------

Private Function DrawShape( _
        ws As Worksheet, env As T_EnvIDF, _
        arr As Variant, r As Long, _
        x As Double, y As Double) As Shape
    
    Dim sc As Double, scx As Double, scy As Double
    Dim a0 As Double, f0 As Boolean
    Dim x0 As Double, y0 As Double, z0 As Double
    a0 = env.angle
    sc = env.sc
    If sc = 0 Then Exit Function
    scx = sc
    If env.flip Then scx = -scx
    scy = -sc
    x0 = env.x0
    y0 = env.y0

    Dim sh As Shape
    Dim fb As FreeformBuilder
    
    '開始点
    Dim sect As String
    Dim label As Integer
    Dim Index As Integer
    Dim a1 As Double, x1 As Double, y1 As Double, h1 As Double
    sect = arr(r, FID.N_SECTION)
    label = arr(r, FID.N_LABEL)
    Index = arr(r, FID.N_INDEX)
    a1 = arr(r, FID.N_ANGLE)
    x1 = arr(r, FID.N_XPOS)
    y1 = arr(r, FID.N_YPOS)
    h1 = arr(r, FID.N_HEIGHT)
    r = r + 1
    
    Dim tx As Double, tx1 As Double, tx2 As Double
    Dim ty As Double, ty1 As Double, ty2 As Double
    tx = x + (Cos(a0) * x1 - Sin(a0) * y1): tx1 = tx: tx2 = tx
    ty = y + (Sin(a0) * x1 + Cos(a0) * y1): ty1 = ty: ty2 = ty
    
    Dim px As Double, py As Double
    px = x0 + scx * tx
    py = y0 + scy * ty
    
    Dim a2 As Double, x2 As Double, y2 As Double
    Dim dx As Double, dy As Double
    Dim d As Double
    
    '弧の分割数
    Dim n As Integer
    n = 7
    
    '継続点
    Do While r <= UBound(arr)
        If sect <> arr(r, FID.N_SECTION) Then Exit Do
        If label <> arr(r, FID.N_LABEL) Then Exit Do
        If Index > arr(r, FID.N_INDEX) Then Exit Do
        label = arr(r, FID.N_LABEL)
        Index = arr(r, FID.N_INDEX)
        a2 = arr(r, FID.N_ANGLE)
        x2 = arr(r, FID.N_XPOS)
        y2 = arr(r, FID.N_YPOS)
        
        '円描画
        If CInt(a2) = 360 Then
            dx = x2 - x1
            dy = y2 - y1
            d = sc * Sqr(dx * dx + dy * dy)
            Set sh = ws.Shapes.AddShape(msoShapeOval, px - d, py - d, 2 * d, 2 * d)
            tx1 = tx1 - Sqr(dx * dx + dy * dy)
            tx2 = tx2 + Sqr(dx * dx + dy * dy)
            ty1 = ty1 - Sqr(dx * dx + dy * dy)
            ty2 = ty2 + Sqr(dx * dx + dy * dy)
            Set fb = Nothing
            r = r + 1
            Exit Do
        End If
        
        'フリーフォーム
        If fb Is Nothing Then
            Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, px, py)
        End If
        If CInt(a2) = 0 Then
            tx = x + (Cos(a0) * x2 - Sin(a0) * y2)
            ty = y + (Sin(a0) * x2 + Cos(a0) * y2)
            px = x0 + scx * tx
            py = y0 + scy * ty
            fb.AddNodes msoSegmentLine, msoEditingAuto, px, py
            If tx < tx1 Then tx1 = tx
            If tx > tx2 Then tx2 = tx
            If ty < ty1 Then ty1 = ty
            If ty > ty2 Then ty2 = ty
        Else
            tx = x + (Cos(a0) * x1 - Sin(a0) * y1)
            ty = y + (Sin(a0) * x1 + Cos(a0) * y1)
            px = x0 + scx * tx
            py = y0 + scy * ty
            fb.AddNodes msoSegmentLine, msoEditingAuto, px, py
            If tx < tx1 Then tx1 = tx
            If tx > tx2 Then tx2 = tx
            If ty < ty1 Then ty1 = ty
            If ty > ty2 Then ty2 = ty
            dx = (x2 - x1) / 2
            dy = (y2 - y1) / 2
            Dim a3 As Double, x3 As Double, y3 As Double
            a3 = Round(Tan(wsf.Pi * (180 - a2) / 360), 5)
            x3 = (x1 + x2) / 2 - dy * a3
            y3 = (y1 + y2) / 2 + dx * a3
            Dim aa As Double, ax As Double, ay As Double
            aa = a2 / n * wsf.Pi / 180
            ax = (x1 - x3)
            ay = (y1 - y3)
            Dim i As Integer
            For i = 0 To n - 1
                Dim x4 As Double, y4 As Double
                x4 = x3 + Cos(i * aa) * ax - Sin(i * aa) * ay
                y4 = y3 + Sin(i * aa) * ax + Cos(i * aa) * ay
                tx = x + (Cos(a0) * x4 - Sin(a0) * y4)
                ty = y + (Sin(a0) * x4 + Cos(a0) * y4)
                px = x0 + scx * tx
                py = y0 + scy * ty
                fb.AddNodes msoSegmentCurve, msoEditingAuto, px, py
                If tx < tx1 Then tx1 = tx
                If tx > tx2 Then tx2 = tx
                If ty < ty1 Then ty1 = ty
                If ty > ty2 Then ty2 = ty
            Next i
            tx = x + (Cos(a0) * x2 - Sin(a0) * y2)
            ty = y + (Sin(a0) * x2 + Cos(a0) * y2)
            px = x0 + scx * tx
            py = y0 + scy * ty
            fb.AddNodes msoSegmentCurve, msoEditingAuto, px, py
            fb.AddNodes msoSegmentLine, msoEditingAuto, px, py
            If tx < tx1 Then tx1 = tx
            If tx > tx2 Then tx2 = tx
            If ty < ty1 Then ty1 = ty
            If ty > ty2 Then ty2 = ty
        End If
    
        r = r + 1
        x1 = x2
        y1 = y2
    Loop
    r = r - 1
    
    If Not fb Is Nothing Then
        Set sh = fb.ConvertToShape
        Set fb = Nothing
    End If
    If sh Is Nothing Then Exit Function
    
    sh.LockAspectRatio = msoTrue
    sh.Placement = xlMove
    '
    With sh.ThreeD
        .BevelTopType = msoBevelAngle
        .BevelTopInset = 0
        .BevelTopDepth = 0
        .BevelBottomType = msoBevelAngle
        .BevelBottomInset = 0
        .BevelBottomDepth = 0
        .Depth = sc * h1
        If env.scz > 0 Then
            z0 = (env.z0 + h1)
        Else
            z0 = env.z0
        End If
        .Z = sc * z0
    End With
    
    px = (sh.Left - x0) / scx
    py = (sh.Top - y0) / scy
    dx = tx2 - tx1
    dy = ty2 - ty1
    
    sh.Title = sect
    Dim s As String
    's = Join(Array("p:" & px & "," & py & "," & z0), Chr(10))
    s = Join(Array("d:" & dx & "," & dy & "," & h1), Chr(10))
    s = Join(Array(s, "sc:" & sc), Chr(10))
    sh.AlternativeText = s
    
    Set DrawShape = sh
    Set sh = Nothing

End Function

'----------------------------------------
'IDF表示属性
'----------------------------------------

'種別に合わせたスタイル設定
Private Sub SetStyleIDF(obj As Object, k1 As String, Optional k2 As String)
    Select Case k1
    Case "ELECTRICAL"
    Case "MECHANICAL"
    Case "HEADER"
    Case "BOARD_OUTLINE"
        If k2 = "0" Then
            obj.Fill.ForeColor.RGB = RGB(0, 127, 0)
            obj.Fill.Transparency = 0
        Else
            obj.Fill.ForeColor.RGB = RGB(0, 0, 0)
            obj.Fill.Transparency = 0.4
        End If
    Case "PANEL_OUTLINE"
        If k2 = "0" Then
            obj.Fill.ForeColor.RGB = RGB(127, 0, 0)
            obj.Fill.Transparency = 0
        Else
            obj.Fill.ForeColor.RGB = RGB(0, 0, 0)
            obj.Fill.Transparency = 0.4
        End If
    Case "OTHER_OUTLINE"
        If k2 = "0" Then
            obj.Fill.ForeColor.RGB = RGB(0, 0, 127)
            obj.Fill.Transparency = 0
        Else
            obj.Fill.ForeColor.RGB = RGB(0, 0, 0)
            obj.Fill.Transparency = 0.4
        End If
    Case "ROUTE_OUTLINE"
        obj.Fill.Visible = msoFalse
        obj.ThreeD.PresetMaterial = msoMaterialWireFrame
    Case "PLACE_OUTLINE"
        obj.Fill.Visible = msoFalse
        obj.ThreeD.PresetMaterial = msoMaterialWireFrame
    Case "ROUTE_KEEPOUT"
        obj.Fill.Visible = msoFalse
        obj.ThreeD.PresetMaterial = msoMaterialWireFrame
    Case "VIA_KEEPOUT"
        obj.Fill.Visible = msoFalse
        obj.ThreeD.PresetMaterial = msoMaterialWireFrame
    Case "PLACE_KEEPOUT"
        obj.Fill.Visible = msoFalse
        obj.ThreeD.PresetMaterial = msoMaterialWireFrame
    Case "PLACE_REGION"
        obj.Fill.Visible = msoFalse
        obj.ThreeD.PresetMaterial = msoMaterialWireFrame
    Case "DRILLED_HOLES"
        obj.line.ForeColor.RGB = RGB(0, 0, 0)
        obj.line.Weight = 0
        obj.line.Visible = True
        obj.Fill.ForeColor.RGB = RGB(0, 0, 0)
        obj.Fill.Visible = False
    Case "NOTES"
    Case "PLACEMENT"
        obj.Fill.Visible = msoTrue
        obj.Fill.ForeColor.RGB = RGB(127, 127, 127)
        obj.Fill.Transparency = 0
    Case Else
    End Select
End Sub

'----------------------------------------
'IDF表示属性
'----------------------------------------

Sub ResetShapeSize()
    If TypeName(Selection) = "Range" Then Exit Sub
    Dim sr As Variant
    Set sr = Selection.ShapeRange
    
    'コレクション作成
    Dim col As Collection
    Set col = New Collection
    Dim sh As Shape, sh2 As Shape
    Dim s As String
    For Each sh In sr
        If sh.Type = msoGroup Then
            For Each sh2 In sh.GroupItems
                col.Add sh2.name
            Next sh2
        Else
            col.Add sh.name
        End If
    Next sh
    
    Dim v As Variant
    For Each v In col
        Set sh = ActiveSheet.Shapes(v)
        Dim sc As Double, d1, d2, d3, d4
        
        sc = ParamStrVal(sh.AlternativeText, "sc")
        Dim arr As Variant
        arr = StrToArr(ParamStrVal(sh.AlternativeText, "d"))
        sh.Width = sc * arr(0)
        sh.Height = sc * arr(1)
        sh.ThreeD.Depth = sc * arr(2)
    Next v
    
    Set col = Nothing
End Sub
    
Sub ResizeShapeScale()
    If TypeName(Selection) = "Range" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'コレクション作成
    Dim grp_name As String
    Dim sr As ShapeRange
    Set sr = Selection.ShapeRange
    
    If sr.Count <> 1 Then
    ElseIf sr.Type = msoGroup Then
        grp_name = sr.name
        Selection.ShapeRange.Ungroup.Select
        Set sr = Selection.ShapeRange
    End If
    
    'コレクション作成
    Dim col As Collection
    Set col = New Collection
    Dim sh As Shape
    For Each sh In sr
        col.Add sh.name
    Next sh
    
    '3D 奥行設定
    Dim sr2 As ShapeRange
    Dim v As Variant
    For Each v In col
        Set sh = ActiveSheet.Shapes(v)
        If sh.Type = msoGroup Then
            Dim sh2 As Shape
            For Each sh2 In sh.GroupItems
                If sh2.ThreeD.Z < -100000 Then
                ElseIf sh2.ThreeD.Depth < -10000 Then
                Else
                    sh2.ThreeD.Z = sh2.ThreeD.Depth - sh2.ThreeD.Z - 1.6 * 2
                End If
            Next sh2
        Else
            If sh.ThreeD.Z < -100000 Then
            ElseIf sh.ThreeD.Depth < -10000 Then
            Else
                sh.ThreeD.Z = sh.ThreeD.Depth - sh.ThreeD.Z - 1.6 * 2
            End If
        End If
    Next v
    
    '再グループ化
    If grp_name <> "" Then
        Selection.Group.Select
        Set sr = Selection.ShapeRange
        sr.name = grp_name
        grp_name = ""
    End If
    sr.Select

    If sr.ThreeD.Visible Then
        With sr.ThreeD
            .Visible = True
            .SetPresetCamera (msoCameraIsometricTopUp)
            .RotationX = 45.2809
            .RotationY = -35.3962666667
            .RotationZ = -60.1624166667
        End With
    End If

End Sub

