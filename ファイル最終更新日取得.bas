Attribute VB_Name = "�t�@�C���ŏI�X�V���擾"
Sub �t�@�C���ŏI�X�V���擾M()
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String

    ' A��̍ŏI�s���擾
    lastRow = Cells(Rows.count, 1).End(xlUp).row

    ' 1�s�ڂ���ŏI�s�܂Ń��[�v
    For i = 1 To lastRow
        filePath = Cells(i, 1).Value ' A��̃t�@�C���p�X���擾

        ' �t�@�C�������݂���ꍇ�̂ݍŏI�X�V����B��ɏ�������
        If Dir(filePath) <> "" Then
            Cells(i, 2).Value = FileDateTime(filePath) ' B��ɍŏI�X�V������������
        Else
            Cells(i, 2).Value = "�t�@�C�������݂��܂���" ' �t�@�C�������݂��Ȃ��ꍇ�̃��b�Z�[�W
        End If
    Next i
End Sub
