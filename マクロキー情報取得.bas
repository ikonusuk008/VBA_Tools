Attribute VB_Name = "�}�N���L�[���擾"
Sub �}�N���L�[���擾M()
    ' �V�[�g�u�L�[�v��I��
    On Error Resume Next
    sheets("�L�[�ꗗ").Select
    If Err.Number <> 0 Then
        MsgBox "�V�[�g�u�L�[�v��������܂���B�V�[�g�����m�F���Ă��������B", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' �ϐ��̏�����
    Dim �p�X As String
    Dim �t�@�C���ԍ� As Integer
    Dim �s�f�[�^ As String
    Dim �s�ԍ� As Integer
    �s�ԍ� = 1
    Dim �ꎞ�z��() As String
    Dim ���W���[���� As String
    Dim �L�[�� As String
    Dim �R���|�[�l���g As Object
    
    ' �萔�̒�`
    Const ����1 As String = "Attribute "
    Const ����2 As String = "VB_Invoke_Func ="
    Const �ꎞ�t�@�C�� As String = "Temp1.bas"
    
    ' �p�X�̎擾
    �p�X = ThisWorkbook.path & "\"
    
    ' VBProject����e�R���|�[�l���g���擾���A����
    With ThisWorkbook.VBProject
        For Each �R���|�[�l���g In .VBComponents
            ' �R���|�[�l���g���ꎞ�t�@�C���Ƃ��ăG�N�X�|�[�g
            .VBComponents(�R���|�[�l���g.name).Export fileName:=�p�X & �ꎞ�t�@�C��
            �t�@�C���ԍ� = FreeFile()
            Open �p�X & �ꎞ�t�@�C�� For Input As #�t�@�C���ԍ�
            
            ' �t�@�C�����e��ǂݎ��
            While Not EOF(�t�@�C���ԍ�)
                Line Input #�t�@�C���ԍ�, �s�f�[�^
                
                ' Sub�����擾
                If InStr(1, �s�f�[�^, "Sub", vbTextCompare) = 1 Then
                    ���W���[���� = Mid$(�s�f�[�^, InStr(�s�f�[�^, "Sub") + 4)
                End If
                
                ' Attribute�����擾
                If InStr(�s�f�[�^, ����1) = 1 And InStr(�s�f�[�^, ����2) > 0 Then
                    ReDim Preserve �ꎞ�z��(�s�ԍ�)
                    �L�[�� = ":" & Mid$(�s�f�[�^, InStrRev(�s�f�[�^, "=") + 3, 1)
                    �ꎞ�z��(�s�ԍ�) = ���W���[���� & �L�[��
                    
                    ' ���ʂ��V�[�g�ɏo��
                    Cells(�s�ԍ� + 1, 1) = ���W���[����
                    Cells(�s�ԍ� + 1, 2) = Replace(�L�[��, ":", "")
                    �s�ԍ� = �s�ԍ� + 1
                    ���W���[���� = ""
                End If
                �s�f�[�^ = ""
            Wend
            
            ' �t�@�C������č폜
            Close #�t�@�C���ԍ�
            Kill �p�X & �ꎞ�t�@�C��
        Next
    End With
    
    ' �����ݒ�
    �L�[���X�g�����ݒ�
End Sub

Sub �L�[���X�g�����ݒ�()
    ' �w�b�_�����̏����ݒ�
    With Range("A1:C1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    ' �t���[�Y�y�C���̐ݒ�
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    ' �I�[�g�t�B���^�̐ݒ�
    Cells.Select
    Selection.AutoFilter
    
    ' �\���{���̐ݒ�
    ActiveWindow.Zoom = 75
    
    ' �񕝂̐ݒ�
    Columns("A:A").ColumnWidth = 60
    Columns("B:B").ColumnWidth = 10
    
    ' �w�b�_�̐ݒ�
    Range("A1").Value = "���W���[����"
    Range("B1").Value = "�L�[��"
    Range("C1").Value = "�L�[���"
    
    ' �ŏ��̃Z����I��
    Range("A1").Select
End Sub

