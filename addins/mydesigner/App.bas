Attribute VB_Name = "App"
'==================================
'�A�v���P�[�V����
'==================================

'[�Q�Ɛݒ�]
'�uMicrosoft Scripting Runtime�v

Option Explicit
Option Private Module


'�p�����[�^
'  �ۑ���F
'   �������A
'   �A�h�C���t�@�C���A�A�h�C���u�b�N�A�A�h�C���V�[�g�A�}�`
'   �G�N�Z���t�@�C���A�G�N�Z���u�b�N�A�G�N�Z���V�[�g�A�}�`


'   �O���[�o���ϐ�      EXCEL�N�������A�A�h�C�������AVBA���W���[�����ł݂̂ŗL��
'   ���s���p�����[�^    EXCEL�N�������A�A�h�C�����ł݂̂ŗL��
'   �u�b�N�p�����[�^    �u�b�N�ɕt��(�t�@�C���̃v���p�e�B�ŕύX�\)
'   ���O                �u�b�N�ɕt��(�Q�Ƃ͎����v�Z)
'   �V�[�g�p�����[�^    �V�[�g�ɕt��(�Q�Ƃ͎����v�Z)
'   �}�`�p�����[�^      �}�`�ɕt��
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

