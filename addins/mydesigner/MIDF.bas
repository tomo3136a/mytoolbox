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
    angle As Double
    flip As Boolean
    dir As Boolean
End Type

'----------------------------------------
'IDF読み込み
'----------------------------------------

'IDFファイルを読み込み、シート作成
Public Sub ImportIDF()

    Dim path As String
    path = GetRtParam("IDF", "path")
    If path = "" Then path = ActiveWorkbook.path
    
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "IDF file", "*.emn,*.brd,*.bdf,*.idb"
        .Filters.Add "library file", "*.emp,*.lib,*.ldf,*.idl"
        .Filters.Add "全てのファイル", "*.*"
        .FilterIndex = 1
        .InitialFileName = path & "\"
        .AllowMultiSelect = True
        If Not .Show Then Exit Sub
        
        '画面チラつき防止処置
        ScreenUpdateOff
        
        'ファイル読み込み
        Dim v As Variant
        For Each v In .SelectedItems
            path = v
            
            'ファイルを配列に読み込み
            Dim arr As Variant
            ReadArrayIDF arr, path, True
            If UBound(arr, 2) < 2 Then Exit For
            
            'ワークシート作成
            Dim ws As Worksheet
            Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
            ws.name = UniqueSheetName(ws.Parent, fso.GetFileName(path))
            ws.Range("C:C").NumberFormatLocal = "#0.0###"
            ws.Range("M:N").NumberFormatLocal = "#0.0###"
            ws.Range("W:Y").NumberFormatLocal = "#0.0###"
            
            'ワークシートに出力
            Dim rcnt As Long
            Dim ccnt As Long
            rcnt = UBound(arr, 1)
            ccnt = UBound(arr, 2)
            ws.Range("A1").Resize(rcnt, ccnt).Value = arr
            Set ws = Nothing
        Next v
        
        '画面チラつき防止処置解除
        ScreenUpdateOn
    End With

End Sub

'IDFファイルを読み込み、配列作成
Private Sub ReadArrayIDF(arr As Variant, path As String, Optional hdr As Boolean)
    
    Dim col As Collection
    Set col = New Collection

    'ヘッダー行作成
    Dim ccnt As Long
    If hdr Then
        Dim vh As Variant
        vh = Split("," & FHDR, ",")
        col.Add vh
        ccnt = UBound(vh) + 1
    End If
    
    '読み込み
    ReadIDF col, path
    Dim rcnt As Long
    rcnt = col.Count
    If rcnt < 2 Then Exit Sub
    
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
Private Sub ReadIDF(col As Collection, path As String)
    
    SetRtParam "IDF", "path", fso.GetParentFolderName(path)
    If Not fso.FileExists(path) Then Exit Sub
    Dim file_name As String
    file_name = fso.GetFileName(path)
    
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
    Dim index As Long
    
    'ファイルを読み込み、行ごとに処理
    Dim st As Object
    Set st = fso.GetFile(path).OpenAsTextStream
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
                    index = 0
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
                    If wa(FID.N_LABEL) = s Then index = index + 1 Else index = 0
                    wa(FID.N_LABEL) = s
                    wa(FID.N_INDEX) = index
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
                    wa(FID.N_HEIGHT) = va(2)
                    wa(FID.N_ANGLE) = va(3)
                    wa(FID.N_LAYER) = va(4)
                    wa(FID.N_STATUS) = va(5)
                    col.Add wa
                    seq = seq - 2
                End Select
                seq = seq + 1
            
            Case EM1.N_MATERIAL
                Select Case seq
                Case 1
                    wa(FID.N_SECTION) = sect
                    wa(FID.N_OWNER) = ""
                    index = 0
                Case 2
                    wa(FID.N_GEOMETORY) = s
                    wa(FID.N_NUMBER) = va(1)
                    wa(FID.N_UNITS) = va(2)
                    wa(FID.N_HEIGHT) = va(3)
                Case Else
                    If wa(FID.N_LABEL) = s Then index = index + 1 Else index = 0
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
                        wa(FID.N_INDEX) = index
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
    root = GetRtParam("IDF", "path")
    If root = "" Then root = ActiveWorkbook.path

    Dim ws As Worksheet
    For Each ws In ActiveWindow.SelectedSheets
        
        '出力パスの選択
        Dim name As String
        Dim path As String
        name = re_replace(ws.name, "\s*\(\d+\)$", "")
        path = fso.BuildPath(root, name)
        Dim flt As String
        flt = "IDF file,*.emn,IDF file,*.brd,IDF file,*.bdf,IDF file,*.idb"
        flt = flt & ",library file,*.emp,library file,*.lib"
        flt = flt & ",library file,*.ldf,library file,*.idl"
        flt = flt & ",all file,*.*"
        Dim idx As Integer
        idx = 9
        If LCase(Right(path, 4)) = ".emn" Then idx = 1
        If LCase(Right(path, 4)) = ".brd" Then idx = 2
        If LCase(Right(path, 4)) = ".bdf" Then idx = 3
        If LCase(Right(path, 4)) = ".idb" Then idx = 4
        If LCase(Right(path, 4)) = ".emp" Then idx = 5
        If LCase(Right(path, 4)) = ".lib" Then idx = 6
        If LCase(Right(path, 4)) = ".ldf" Then idx = 7
        If LCase(Right(path, 4)) = ".idl" Then idx = 8
        path = Application.GetSaveAsFilename(path, flt, idx)
        If path = "False" Then Exit Sub
        
        'ファイル書き出し
        Dim ra As Range
        Set ra = ws.UsedRange
        If ra.Count > 2 Then
            Dim arr As Variant
            arr = ra.Value
            WriteIDF path, arr
        End If
    Next ws
    
    SetRtParam "IDF", "path", fso.GetParentFolderName(path)

End Sub

'配列からIDFファイル書き出し
Private Sub WriteIDF(path As String, arr As Variant)
    
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
    
    Open path For Output As #1
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
                    line = SL(arr(r, FID.N_NAME), 15)
                    line = line & sr(arr(r, FID.N_UNITS), 4)
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
                    line = SL(line, 16) & "   "
                    line = line & sr(arr(r, FID.N_OWNER), 8)
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
                        line = SL(arr(r, FID.N_LAYER), 8) & "   "
                        line = line & sr(arr(r, FID.N_HEIGHT), 8)
                    Case EM2.N_PLACE_REGION
                        line = arr(r, FID.N_LAYER)
                        line = line & "  " & arr(r, FID.N_REFERENCE)
                    End Select
                    If mode2 <> EM2.N_VIA_KEEPOUT Then Print #1, RTrim(line)
                    seq = 2
                End If
                line = arr(r, FID.N_LABEL) & "  "
                line = line & sr(arr(r, FID.N_XPOS), 8)
                line = line & sr(arr(r, FID.N_YPOS), 8)
                line = line & sr(arr(r, FID.N_ANGLE), 8)
                Print #1, RTrim(line)
            
            Case EM1.N_DRILLED_HOLES
                line = sr(arr(r, FID.N_LENGTH), 8)
                line = line & sr(arr(r, FID.N_XPOS), 8)
                line = line & sr(arr(r, FID.N_YPOS), 8)
                line = line & sr(arr(r, FID.N_GEOMETORY), 8)
                line = line & sr(arr(r, FID.N_REFERENCE), 8)
                line = line & sr(arr(r, FID.N_NUMBER), 8)
                line = line & sr(arr(r, FID.N_OWNER), 8)
                Print #1, RTrim(line)
            
            Case EM1.N_NOTES
                line = sr(arr(r, FID.N_XPOS), 8)
                line = line & sr(arr(r, FID.N_YPOS), 8)
                line = line & sr(arr(r, FID.N_HEIGHT), 8)
                line = line & sr(arr(r, FID.N_LENGTH), 8)
                line = line & "  " & arr(r, FID.N_VAL)
                Print #1, RTrim(line)
            
            Case EM1.N_PLACEMENT
                line = sr(arr(r, FID.N_GEOMETORY))
                line = line & sr(arr(r, FID.N_NUMBER))
                line = line & sr(arr(r, FID.N_REFERENCE))
                Print #1, RTrim(line)
                line = sr(arr(r, FID.N_XPOS), 8)
                line = line & sr(arr(r, FID.N_YPOS), 8)
                line = line & sr(arr(r, FID.N_HEIGHT), 8)
                line = line & sr(arr(r, FID.N_ANGLE), 8)
                line = line & sr(arr(r, FID.N_LAYER), 8)
                line = line & sr(arr(r, FID.N_STATUS), 8)
                Print #1, RTrim(line)
            
            Case EM1.N_MATERIAL
                If seq = 1 Then
                    line = sr(arr(r, FID.N_GEOMETORY))
                    line = line & sr(arr(r, FID.N_NUMBER))
                    line = line & sr(arr(r, FID.N_UNITS), 8)
                    line = line & sr(arr(r, FID.N_HEIGHT), 8)
                    Print #1, RTrim(line)
                    seq = 2
                End If
                line = arr(r, FID.N_LABEL)
                If line = "PROP" Then
                    line = line & sr(arr(r, FID.N_ATTRIBUTE), 12)
                    line = line & sr(arr(r, FID.N_VAL), 8)
                Else
                    line = line & "  " & sr(arr(r, FID.N_XPOS), 8)
                    line = line & sr(arr(r, FID.N_YPOS), 8)
                    line = line & sr(arr(r, FID.N_ANGLE), 8)
                End If
                Print #1, RTrim(line)
            
            End Select
        End If
    Next r
    If sect <> "" Then Print #1, (".END_" & sect)
    Close #1

End Sub

'右寄せ
Private Function sr(s As Variant, Optional n As Integer = 16)
    sr = Right("                " & Format(s, "0.0"), n)
End Function

'左寄せ
Private Function SL(s As Variant, Optional n As Integer = 16)
    SL = Left(Format(s, "0.0") & "                ", n)
End Function

'-------------------------------------
'IDF描画
'-------------------------------------

'ワークシートからIDF描画
Public Sub DrawIDF(ws As Worksheet, x As Double, y As Double)
    '
    'データライブラリ取得
    Dim name As String
    Dim lib As Dictionary
    name = GetDataIDF(lib)
    If name = "*" Then Exit Sub
    If name = "" Then Exit Sub
    
   'データ配列取得
    Dim arr As Variant
    If Not lib("$arr").Exists(name) Then Exit Sub
    arr = lib("$arr")(name)
    
    '描画スケール取得
    Dim sc As Double
    sc = GetDrawParam(2)
    
    '描画領域計算
    Dim dra As Variant
    dra = DrawingAreaIDF(arr)
    
    '原点計算
    Dim x0 As Double
    Dim y0 As Double
    x0 = x - sc * dra(0)
    y0 = y + sc * dra(1) + sc * dra(5)
    
    Dim a As Double
    Dim f As Boolean
    
    '描画環境設定
    Dim env As T_EnvIDF
    SetEnvIDF env, x0, y0, 0, sc
    
    Dim pos As Dictionary
    Set pos = lib("$pos")

    Dim ns2 As Collection
    Set ns2 = New Collection
    
    Dim Sh As Shape
    Dim k As String
    Dim i As Long
    Dim r As Long
    Dim s As String
    
    Dim ns As Collection
    Set ns = New Collection
    
    'OUTLINE, KEEPOUT
    'k = Join(Array(name, "BOARD_OUTLINE", "", ""), "-")
    'Set sh = DrawGroupOutline(ws, env, k, arr, pos)
    'If Not sh Is Nothing Then ns.Add sh.name
    'k = Join(Array(name, "PANEL_OUTLINE", "", ""), "-")
    'Set sh = DrawGroupOutline(ws, env, k, arr, pos)
    'If Not sh Is Nothing Then ns.Add sh.name
    k = Join(Array(name, "DRILLED_HOLES", "", ""), "-")
    Set Sh = DrawGroupHole(ws, env, k, arr, pos)
    'If Not sh Is Nothing Then ns.Add sh.name
    'If ns.Count > 0 Then Set sh = GroupShape(ws, ns, "OUTLINE")
    If Not Sh Is Nothing Then ns2.Add Sh.name
    
    'OUTLINE, KEEPOUT
    'OUTLINE, KEEPOUT
    k = Join(Array(name, "VIA_KEEPOUT", "", ""), "-")
    Set Sh = DrawGroupOutline(ws, env, k, arr, pos)
    If Not Sh Is Nothing Then ns2.Add Sh.name

    Dim side As Variant
    side = Array("ALL", "BOTH", "BOTTOM", "INNER", "TOP")
    For i = 0 To UBound(side)
        Set ns = New Collection
        k = Join(Array(name, "OTHER_OUTLINE", side(i), ""), "-")
        Set Sh = DrawGroupOutline(ws, env, k, arr, pos)
        If Not Sh Is Nothing Then ns.Add Sh.name
        k = Join(Array(name, "ROUTE_OUTLINE", side(i), ""), "-")
        Set Sh = DrawGroupOutline(ws, env, k, arr, pos)
        If Not Sh Is Nothing Then ns.Add Sh.name
        k = Join(Array(name, "ROUTE_KEEPOUT", side(i), ""), "-")
        Set Sh = DrawGroupOutline(ws, env, k, arr, pos)
        If Not Sh Is Nothing Then ns.Add Sh.name
        k = Join(Array(name, "PLACE_OUTLINE", side(i), ""), "-")
        Set Sh = DrawGroupOutline(ws, env, k, arr, pos)
        If Not Sh Is Nothing Then ns.Add Sh.name
        k = Join(Array(name, "PLACE_KEEPOUT", side(i), ""), "-")
        Set Sh = DrawGroupOutline(ws, env, k, arr, pos)
        If Not Sh Is Nothing Then ns.Add Sh.name
        k = Join(Array(name, "PLACE_REGION", side(i), ""), "-")
        Set Sh = DrawGroupOutline(ws, env, k, arr, pos)
        If Not Sh Is Nothing Then ns.Add Sh.name
        k = Join(Array(name, "PLACEMENT", side(i), ""), "-")
        Set Sh = DrawGroupPlace(ws, env, k, arr, pos, lib)
        If Not Sh Is Nothing Then ns.Add Sh.name
        Set Sh = Nothing
        If ns.Count > 0 Then Set Sh = GroupShape(ws, ns, k)
        If Not Sh Is Nothing Then ns2.Add Sh.name
    Next i
    
    k = Join(Array(name, "NOTES", "", ""), "-")
    Set Sh = DrawGroupNote(ws, env, k, arr, pos)
    If Not Sh Is Nothing Then ns2.Add Sh.name
    
    'ORIGIN
    Set Sh = DrawOrigin(ws, env, arr, pos)
    If Not Sh Is Nothing Then ns2.Add Sh.name
    
    If ns2.Count > 0 Then Set Sh = GroupShape(ws, ns2, name)

End Sub

'描画パラメータ設定
Private Function DrawEnv( _
    x0 As Double, y0 As Double, _
    g As Double, a As Double, f As Boolean) As Variant
    DrawEnv = Array(a, g, f, x0, y0)
End Function

Private Sub SetEnvIDF(env As T_EnvIDF, _
    Optional x0 As Double, Optional y0 As Double, Optional z0 As Double, _
    Optional sc As Double, Optional angle As Double, _
    Optional flip As Boolean, Optional dir As Boolean)
    env.x0 = x0
    env.y0 = y0
    env.z0 = z0
    env.sc = sc
    env.angle = angle
    env.flip = flip
    env.dir = dir
End Sub
    

'描画配列から全体範囲を取得
Private Function DrawingAreaIDF(arr As Variant) As Variant
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Select Case UCase(Trim(arr(r, FID.N_FILE_TYPE)))
        Case "BOARD_FILE"
            Exit For
        Case "PANEL_FILE"
            Exit For
        Case "LIBRARY_FILE"
            Exit For
        End Select
    Next r
    If r > UBound(arr, 1) Then Exit Function
    
    Dim x As Double, y As Double
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

    DrawingAreaIDF = Array(xs, ys, xe, ye, xe - xs, ye - ys)
End Function

'ライブラリを辞書に登録
Private Function GetDataIDF(dic As Dictionary) As String
    
    'ライブラリシート選択
    Dim ws As Worksheet
    Set ws = SelectSheetCB(ActiveWorkbook)
    If ws Is ActiveSheet Then
        GetDataIDF = "*"
        Exit Function
    End If
    
    'ライブラリデータ配列取得
    Dim arr As Variant
    If ws.UsedRange.Rows.Count < 1 Then Exit Function
    If ws.UsedRange.Columns.Count < FID.N_VAL Then Exit Function
    arr = ws.UsedRange.Value
    Set ws = Nothing
    
    'データ配列の名前取得
    Dim dataname As String
    dataname = GetDataName(arr)
    
    'データ配列を解析し辞書に追加
    ParseArrayIDF arr, dic
    
    GetDataIDF = dataname

End Function

'データ配列からデータ名取得
Private Function GetDataName(arr As Variant) As String
    
    Dim name As String
    Dim r As Long
    For r = 1 To 10
        If r > UBound(arr, 1) Then Exit For
        If FID.N_IDF_VERSION > UBound(arr, 2) Then Exit For
        If Not TypeName(arr(r, FID.N_IDF_VERSION)) = "String" Then
            If "" <> arr(r, FID.N_NAME) Then
                name = arr(r, FID.N_NAME)
                Exit For
            End If
        End If
    Next r
    GetDataName = UCase(name)

End Function

Private Sub ParseArrayIDF(arr As Variant, dic As Dictionary)
    
    Dim dic_arr As Dictionary
    Dim dic_pos As Dictionary
    If dic Is Nothing Then
        Set dic = New Dictionary
        Set dic_arr = New Dictionary
        Set dic_pos = New Dictionary
        dic.Add "$arr", dic_arr
        dic.Add "$pos", dic_pos
    Else
        Set dic_arr = dic("$arr")
        Set dic_pos = dic("$pos")
    End If
    
    Dim name As String
    name = GetDataName(arr)
    If dic_arr.Exists(name) Then Exit Sub
    dic_arr.Add name, arr
    
    Dim s As String
    Dim s1 As String
    Dim s2 As String
    Dim kw As String
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        kw = "" & arr(r, FID.N_INDEX)
        kw = UCase(kw)
        If kw = "" Or kw = "0" Then
            kw = UCase(arr(r, FID.N_SECTION))
            If kw = "ELECTRICAL" Or kw = "MECHANICAL" Then
                kw = arr(r, FID.N_GEOMETORY) & "-" & arr(r, FID.N_NUMBER)
                kw = UCase(kw)
                Dim v As Variant
                v = Array(name, r)
                If Not dic.Exists(kw) Then dic.Add kw, v
            Else
                kw = Join(Array(arr(r, FID.N_NAME), kw, arr(r, FID.N_LAYER), arr(r, FID.N_REFERENCE)), "-")
                kw = UCase(kw)
                Dim col As Collection
                If dic_pos.Exists(kw) Then
                    Set col = dic_pos(kw)
                Else
                    Set col = New Collection
                    dic_pos.Add kw, col
                End If
                col.Add r
            End If
        End If
    Next r
    
    Set dic_arr = Nothing
    Set dic_pos = Nothing
    
End Sub

'-------------------------------------

'原点表示
Private Function DrawOrigin( _
        ws As Worksheet, env As T_EnvIDF, _
        arr As Variant, dic As Dictionary) As Shape
    
    Dim g As Double
    g = env.sc
    
    Dim x0 As Double
    Dim y0 As Double
    x0 = env.x0
    y0 = env.y0
    
    Dim dra As Variant
    dra = DrawingAreaIDF(arr)
    
    Dim dx As Double
    Dim dy As Double
    Dim w As Double
    dx = dra(4)
    dy = dra(5)
    
    w = 50
    
    Dim tw As Double
    Dim tx As Double
    Dim ty As Double
    tw = g * w * 2
    tx = x0 - g * w
    ty = y0 - g * w
    
    Dim Sh As Shape
    Set Sh = ws.Shapes.AddShape(msoShapeFlowchartOr, tx, ty, tw, tw)
    With Sh
        .LockAspectRatio = msoTrue
        .Placement = xlMove
        .name = "ORIGIN " & .id
    End With
    
    Set DrawOrigin = Sh
    Set Sh = Nothing
    
    'Set sh = AddShapeRect(ws, env, x0, y0, dx, dy)
    'Set sh = Nothing

End Function

'-------------------------------------
'グループ描画
'-------------------------------------

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
    Dim Sh As Shape
    
    Dim i As Long
    Dim r As Long
    For i = 1 To col.Count
        r = col(i)
        Set Sh = DrawOutline(ws, env, arr, r)
        ns.Add Sh.name
    Next i
    If ns.Count > 1 Then Set Sh = GroupShape(ws, ns, grp)
    If Sh Is Nothing Then Exit Function
    
    SetStyleIDF Sh, grp
    
    Set DrawGroupOutline = Sh

End Function

'ホール描画グループ化
Private Function DrawGroupHole( _
        ws As Worksheet, env As T_EnvIDF, grp As String, _
        arr As Variant, dic As Dictionary) As Shape
        
    Dim ns As Collection
    Set ns = New Collection
    
    Dim col As Collection
    Dim Sh As Shape
    Dim i As Long
    Dim r As Long
    
    'Board
    Dim k As String
    Dim v As Variant
    k = Join(Array(Split(grp, "-")(0), "BOARD_OUTLINE", "", ""), "-")
    If dic.Exists(k) Then
        Set col = dic(k)
        For i = 1 To col.Count
            Set Sh = DrawOutline(ws, env, arr, CLng(col(i)))
            If Not Sh Is Nothing Then
                SetStyleIDF Sh, k
                ns.Add Sh.name
            End If
        Next i
    End If
    
    'Panel
    k = Join(Array(Split(grp, "-")(0), "PANEL_OUTLINE", "", ""), "-")
    If dic.Exists(k) Then
        Set col = dic(k)
        For i = 1 To col.Count
            Set Sh = DrawOutline(ws, env, arr, CLng(col(i)))
            If Not Sh Is Nothing Then
                SetStyleIDF Sh, k
                ns.Add Sh.name
            End If
        Next i
    End If
    
    'Hole
    Dim kw As Variant
    For Each kw In dic.Keys
        k = kw
        If k Like (grp & "*") Then
            Set col = dic(k)
            For i = 1 To col.Count
                env.z0 = env.z0 + 1
                Set Sh = DrawHole(ws, env, arr, CLng(col(i)))
                If Not Sh Is Nothing Then
                    SetStyleIDF Sh, k
                    ns.Add Sh.name
                End If
            Next i
        End If
    Next kw
    
    If ns.Count > 1 Then Set Sh = GroupShape(ws, ns, grp)
    Set ns = Nothing
    Set DrawGroupHole = Sh

End Function

'ノート描画グループ化
Private Function DrawGroupNote( _
        ws As Worksheet, env As T_EnvIDF, grp As String, _
        arr As Variant, dic As Dictionary) As Shape
        
    Dim ns As Collection
    Set ns = New Collection
    
    Dim col As Collection
    Dim Sh As Shape
    
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
                Set Sh = DrawNote(ws, env, arr, r)
                If Not Sh Is Nothing Then ns2.Add Sh.name
            Next i
            If ns2.Count > 1 Then Set Sh = GroupShape(ws, ns2, k)
            Set ns2 = Nothing
            If Not Sh Is Nothing Then ns.Add Sh.name
        End If
    Next kw
    If ns.Count > 1 Then Set Sh = GroupShape(ws, ns, grp)
    Set ns = Nothing
    Set DrawGroupNote = Sh

End Function

'配置グループ
Private Function DrawGroupPlace( _
        ws As Worksheet, env As T_EnvIDF, grp As String, _
        arr As Variant, dic As Dictionary, lib As Dictionary) As Shape
        
    Dim ns As Collection
    Set ns = New Collection
    
    Dim col As Collection
    Dim Sh As Shape
    
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
                Set Sh = DrawPlace(ws, env, k, arr, r, lib)
                If Not Sh Is Nothing Then ns2.Add Sh.name
            Next i
            If ns2.Count > 1 Then Set Sh = GroupShape(ws, ns2, k)
            Set ns2 = Nothing
            If Not Sh Is Nothing Then ns.Add Sh.name
        End If
    Next kw
    If ns.Count > 1 Then Set Sh = GroupShape(ws, ns, grp)
    Set ns = Nothing
    Set DrawGroupPlace = Sh

End Function

'図形名のコレクションから図形グループ化
Private Function GroupShape(ws As Worksheet, ns As Collection, sect As String) As Shape
    Dim Sh As Shape
    If ns.Count > 1 Then
        Set Sh = ws.Shapes.Range(ColToArr(ns)).Group
        Sh.LockAspectRatio = msoTrue
        Sh.Placement = xlMove
        Dim s As String
        s = sect & " " & Sh.id
        Sh.name = s
    Else
        Set Sh = ws.Shapes(ns(1))
    End If
    Set GroupShape = Sh
    Set Sh = Nothing
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
    
    env.z0 = 0
    env.dir = False
    Select Case kw
    Case "BOARD_OUTLINE"
        env.t0 = arr(r, FID.N_HEIGHT)
        env.z0 = -env.t0
    Case "PANEL_OUTLINE"
        env.t0 = arr(r, FID.N_HEIGHT)
        env.z0 = -env.t0
    Case "OTHER_OUTLINE"
    End Select
    
    Select Case CStr(arr(r, FID.N_LAYER))
    Case "TOP"
    Case "BOTTOM"
        env.z0 = -arr(r, FID.N_HEIGHT) - env.t0
        env.dir = True
    End Select
    
    Dim Sh As Shape
    Set Sh = DrawShape(ws, env, arr, r, 0, 0)
    If Sh Is Nothing Then Exit Function
    Sh.name = kw & " " & Sh.id
    
    SetStyleIDF Sh, kw, CStr(arr(r, FID.N_LABEL))
    Set DrawOutline = Sh
    Set Sh = Nothing
    
End Function

'ホール描画
Private Function DrawHole( _
        ws As Worksheet, env As T_EnvIDF, _
        arr As Variant, r As Long) As Shape
    
    Dim sc As Double, x0 As Double, y0 As Double, z0 As Double
    sc = env.sc
    x0 = env.x0
    y0 = env.y0
    z0 = -env.t0

    Dim tw As Double, tx As Double, ty As Double, th As Double
    tw = arr(r, FID.N_LENGTH)
    tx = arr(r, FID.N_LENGTH) / 2
    ty = arr(r, FID.N_LENGTH) / 2
    th = env.t0
    
    Dim pw As Double, px As Double, py As Double
    Dim kw As String
    pw = sc * tw
    px = x0 + sc * (arr(r, FID.N_XPOS) - tx)
    py = y0 - sc * (arr(r, FID.N_YPOS) + ty)
    kw = arr(r, FID.N_GEOMETORY) & "-" & arr(r, FID.N_NUMBER)
    
    Dim Sh As Shape
    Set Sh = ws.Shapes.AddShape(msoShapeOval, px, py, pw, pw)
    With Sh
        .LockAspectRatio = msoTrue
        .Placement = xlMove
        .Title = kw
        .name = kw & " " & .id
    End With
    
    Dim s As String
    s = "IDF " & kw
    s = Join(Array(s, "p:" & tx & "," & ty), Chr(10))
    s = Join(Array(s, "d:" & tw & "," & tw), Chr(10))
    s = Join(Array(s, "g:" & sc), Chr(10))
    Sh.AlternativeText = s
    
    With Sh.ThreeD
        .BevelTopType = msoBevelAngle
        .BevelTopInset = 0
        .BevelTopDepth = 0
        .BevelBottomType = msoBevelAngle
        .BevelBottomInset = 0
        .BevelBottomDepth = 0
        .Depth = sc * th
        .Z = sc * (z0 + th)
    End With
    
    kw = arr(r, FID.N_SECTION)
    SetStyleIDF Sh, kw
    
    Set DrawHole = Sh
    Set Sh = Nothing

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
    px = x0 + sc * (arr(r, FID.N_XPOS) - tx)
    py = y0 - sc * (arr(r, FID.N_YPOS) + ty)
    
    Dim Sh As Shape
    Set Sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, px, py, pw, ph)
    With Sh
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
    SetStyleIDF Sh, kw
    
    Dim s As String
    s = "IDF " & kw
    s = Join(Array(s, "p:" & tx & "," & ty), Chr(10))
    s = Join(Array(s, "d:" & tw & "," & th), Chr(10))
    s = Join(Array(s, "g:" & sc), Chr(10))
    Sh.AlternativeText = s
    
    With Sh.ThreeD
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
    
    Set DrawNote = Sh
    Set Sh = Nothing
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
    SetEnvIDF tenv, tx, ty, 0, env.sc, ta
    
    tenv.z0 = arr(r, FID.N_HEIGHT)
    tenv.t0 = env.t0
    tenv.dir = False
    Select Case CStr(arr(r, FID.N_LAYER))
    Case "TOP"
    Case "BOTTOM"
        tenv.z0 = -arr(r, FID.N_HEIGHT) - tenv.t0
        tenv.dir = True
        tenv.flip = True
    End Select
    
    Dim kw As String
    kw = arr(r, FID.N_GEOMETORY) & " " & arr(r, FID.N_NUMBER)
    kw = UCase(kw)
    
    Dim name As String
    If lib.Exists(kw) Then
        name = GetDataIDF(lib)
    End If
    
    Dim grp1 As String
    Dim grp2 As String
    grp1 = arr(r, FID.N_LAYER)
    grp2 = grp1
    
    Dim r2 As Long
    
    Dim ns As Collection
    Set ns = New Collection
    Dim Sh As Shape
    Set Sh = DrawPart(ws, tenv, tx, ty, arr, r, lib)
    If Sh Is Nothing Then Exit Function
    Sh.name = arr(r, FID.N_REFERENCE) & " " & Sh.id
    ns.Add Sh.name
    
    Set DrawPlace = Sh
    Set Sh = Nothing

End Function

'----------------------------------------

Private Function DrawPart( _
        ws As Worksheet, env As T_EnvIDF, _
        x As Double, y As Double, _
        arr As Variant, r As Long, _
        lib As Dictionary) As Shape
    
    Dim kw As String
    kw = arr(r, FID.N_GEOMETORY) & "-" & arr(r, FID.N_NUMBER)
    kw = UCase(kw)
    If Not lib.Exists(kw) Then
        Dim s As String
        s = GetDataIDF(lib)
        If s = "*" Then
            r = UBound(arr, 1)
            Exit Function
        End If
    End If
    
    Dim part As Variant
    Dim arr2 As Variant
    Dim r2 As Long
    Dim dic2 As Dictionary
    Set dic2 = lib("$arr")
    If Not dic2.Exists("") Then Exit Function
    arr2 = lib("$arr")("")
    part = lib(kw)
    If TypeName(part) = "Empty" Then Exit Function
    r2 = part(1)
    If TypeName(arr2) = "Empty" Then Exit Function
    
    If env.dir Then
        env.z0 = env.z0 - arr2(r2, FID.N_HEIGHT)
    End If
    
    Dim Sh As Shape
    Set Sh = DrawShape(ws, env, arr2, r2, 0, 0)
    If Sh Is Nothing Then Exit Function
    Sh.Title = kw
    Sh.name = kw & " " & Sh.id
    
    kw = arr(r, FID.N_SECTION)
    SetStyleIDF Sh, kw
    
    Set DrawPart = Sh
    Set Sh = Nothing
    
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
    scy = sc
    x0 = env.x0 + scx * x
    y0 = env.y0 - scy * y
    z0 = env.z0

    Dim Sh As Shape
    Dim fb As FreeformBuilder
    
    '開始点
    Dim sect As String
    Dim label As Integer
    Dim index As Integer
    Dim a1 As Double, x1 As Double, y1 As Double, h1 As Double
    sect = arr(r, FID.N_SECTION)
    label = arr(r, FID.N_LABEL)
    index = arr(r, FID.N_INDEX)
    a1 = arr(r, FID.N_ANGLE)
    x1 = arr(r, FID.N_XPOS)
    y1 = arr(r, FID.N_YPOS)
    h1 = arr(r, FID.N_HEIGHT)
    r = r + 1
    
    Dim px As Double, py As Double
    px = x0 + scx * (Cos(a0) * x1 - Sin(a0) * y1)
    py = y0 - scy * (Sin(a0) * x1 + Cos(a0) * y1)
    
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
        If index > arr(r, FID.N_INDEX) Then Exit Do
        label = arr(r, FID.N_LABEL)
        index = arr(r, FID.N_INDEX)
        a2 = arr(r, FID.N_ANGLE)
        x2 = arr(r, FID.N_XPOS)
        y2 = arr(r, FID.N_YPOS)
        
        '円描画
        If CInt(a2) = 360 Then
            dx = x2 - x1
            dy = y2 - y1
            d = sc * Sqr(dx * dx + dy * dy)
            Set Sh = ws.Shapes.AddShape(msoShapeOval, px - d, py - d, 2 * d, 2 * d)
            Set fb = Nothing
            r = r + 1
            Exit Do
        End If
        
        'フリーフォーム
        If fb Is Nothing Then
            Set fb = ws.Shapes.BuildFreeform(msoEditingAuto, px, py)
        End If
        If CInt(a2) = 0 Then
            px = x0 + scx * (Cos(a0) * x2 - Sin(a0) * y2)
            py = y0 - scy * (Sin(a0) * x2 + Cos(a0) * y2)
            fb.AddNodes msoSegmentLine, msoEditingAuto, px, py
        Else
            px = x0 + scx * (Cos(a0) * x1 - Sin(a0) * y1)
            py = y0 - scy * (Sin(a0) * x1 + Cos(a0) * y1)
            fb.AddNodes msoSegmentLine, msoEditingAuto, px, py
            dx = (x2 - x1) / 2
            dy = (y2 - y1) / 2
            Dim a3 As Double, x3 As Double, y3 As Double
            a3 = Round(Tan(wsf.Pi * (180 - a2) / 360), 5)
            x3 = (x1 + x2) / 2 - dy * a3
            y3 = (y1 + y2) / 2 - dx * a3
            Dim aa As Double, ax As Double, ay As Double
            aa = a2 / n * wsf.Pi / 180
            ax = (x1 - x3)
            ay = (y1 - y3)
            Dim i As Integer
            For i = 0 To n - 1
                Dim x4 As Double, y4 As Double
                x4 = x3 + Cos(i * aa) * ax - Sin(i * aa) * ay
                y4 = y3 + Sin(i * aa) * ax + Cos(i * aa) * ay
                px = x0 + scx * (Cos(a0) * x4 - Sin(a0) * y4)
                py = y0 - scy * (Sin(a0) * x4 + Cos(a0) * y4)
                fb.AddNodes msoSegmentCurve, msoEditingAuto, px, py
            Next i
            px = x0 + scx * (Cos(a0) * x2 - Sin(a0) * y2)
            py = y0 - scy * (Sin(a0) * x2 + Cos(a0) * y2)
            fb.AddNodes msoSegmentCurve, msoEditingAuto, px, py
            fb.AddNodes msoSegmentLine, msoEditingAuto, px, py
        End If
    
        r = r + 1
        x1 = x2
        y1 = y2
    Loop
    r = r - 1
    
    If Not fb Is Nothing Then
        Set Sh = fb.ConvertToShape
        Set fb = Nothing
    End If
    If Sh Is Nothing Then Exit Function
    
    Sh.LockAspectRatio = msoTrue
    Sh.Placement = xlMove
    '
    With Sh.ThreeD
        .BevelTopType = msoBevelAngle
        .BevelTopInset = 0
        .BevelTopDepth = 0
        .BevelBottomType = msoBevelAngle
        .BevelBottomInset = 0
        .BevelBottomDepth = 0
        .Depth = sc * h1
        .Z = sc * (z0 + h1)
    End With
    
    px = (Sh.Left - x0) / sc
    py = (Sh.Top - y0) / sc
    dx = Sh.Width / sc
    dy = Sh.Height / sc
    
    Dim s As String
    s = "IDF " & sect
    s = Join(Array(s, "p:" & px & "," & py), Chr(10))
    s = Join(Array(s, "d:" & dx & "," & dy), Chr(10))
    s = Join(Array(s, "g:" & sc), Chr(10))
    Sh.AlternativeText = s
    
    Set DrawShape = Sh
    Set Sh = Nothing

End Function

'----------------------------------------
'
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
        'obj.Fill.Transparency = 0.4
        obj.Fill.Visible = False
        'obj.ThreeD.PresetMaterial = msoMaterialPowder
        'obj.ThreeD.PresetMaterial = msoMaterialTranslucentPowder
        'obj.ThreeD.PresetMaterial = msoMaterialClear
    Case "NOTES"
    Case "PLACEMENT"
        obj.Fill.Visible = msoTrue
        obj.Fill.ForeColor.RGB = RGB(127, 127, 127)
        obj.Fill.Transparency = 0
    Case Else
    End Select
End Sub

Public Sub SetDefaultShapeStyle(Sh As Shape)
    With Sh
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
    End With
End Sub

'----------------------------------------
'
'----------------------------------------

'-------------------------------------

Private Function ArrayToCollection( _
        arr As Variant, s As String, _
        Optional id1 As Integer, _
        Optional id2 As Integer) As Collection
    
    Dim col As Collection
    Set col = New Collection
    
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim s1 As String
        s1 = arr(r, id1)
        Dim s2 As String
        s2 = arr(r, id2)
        If s2 <> "" Then s1 = s1 & "_" & s2
        If s1 = s Then col.Add r
    Next r
    
    Set ArrayToCollection = col

End Function



Public Sub eof()
            ScreenUpdateOn
End Sub

