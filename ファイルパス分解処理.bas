Attribute VB_Name = "�t�@�C���p�X��������"
Sub �t�@�C���p�X��������M()
    Dim �ŏI�s As Long
    Dim i As Long
    Dim �t�@�C���p�X As String
    Dim �p�X����() As String
    Dim �t�@�C���� As String
    Dim �g���q As String
    Dim j As Long
    Dim �f�[�^ As Variant
    Dim ���� As Variant

    ' ���ږ���ݒ�
    Cells(1, 1).Value = "�t�@�C���p�X"
    Cells(1, 2).Value = "�t�@�C����"
    Cells(1, 3).Value = "�g���q"

    ' A��̍ŏI�s���擾
    �ŏI�s = Cells(Rows.count, 1).End(xlUp).row

    ' 2�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 2 To �ŏI�s
        ' A��̃t�@�C���p�X���擾
        �t�@�C���p�X = Cells(i, 1).Value
        
        ' �t�@�C���p�X�𕪉�
        �p�X���� = Split(�t�@�C���p�X, "\")
        
        ' �t�@�C�����Ɗg���q���擾
        �t�@�C���� = �p�X����(UBound(�p�X����))
        �g���q = Split(�t�@�C����, ".")(1)
        
        ' ���ʂ��V�[�g�ɏ�������
        Cells(i, 2).Value = �t�@�C����
        Cells(i, 3).Value = �g���q
        
        ' �t�H���_�K�w���ɕ������Đݒ�
        For j = LBound(�p�X����) To UBound(�p�X����) - 1
            Cells(1, 4 + j).Value = "F" & (j + 1)
            Cells(i, 4 + j).Value = �p�X����(j)
        Next j
    Next i

    ' A��������N�ɐݒ�
    For i = 2 To �ŏI�s
        Cells(i, 1).Hyperlinks.Add Anchor:=Cells(i, 1), Address:=Cells(i, 1).Value, TextToDisplay:=Cells(i, 1).Value
    Next i

    ' ���ڍs���Œ�
    If ActiveWindow.FreezePanes = False Then
        Rows("2:2").Select
        ActiveWindow.FreezePanes = True
    End If

    ' �S�̂��t�B���^
    If ActiveSheet.AutoFilterMode = False Then
        Cells(1, 1).CurrentRegion.AutoFilter
    End If
End Sub

