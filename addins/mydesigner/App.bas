Attribute VB_Name = "App"
'==================================
'アプリケーション
'==================================

'[参照設定]
'「Microsoft Scripting Runtime」

Option Explicit
Option Private Module


'パラメータ
'  保存先：
'   メモリ、
'   アドインファイル、アドインブック、アドインシート、図形
'   エクセルファイル、エクセルブック、エクセルシート、図形


'   グローバル変数      EXCEL起動中かつ、アドイン内かつ、VBAモジュール内でのみで有効
'   実行時パラメータ    EXCEL起動中かつ、アドイン内でのみで有効
'   ブックパラメータ    ブックに付随(ファイルのプロパティで変更可能)
'   名前                ブックに付随(参照は自動計算)
'   シートパラメータ    シートに付随(参照は自動計算)
'   図形パラメータ      図形に付随
'

'----------------------------------------
'
'----------------------------------------

Public Sub eof()
    ScreenUpdateOn
End Sub




'-------------------------------------

Private Sub SetDefaultShapeStyle(sh As Shape)
    With sh
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

