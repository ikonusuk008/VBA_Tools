Attribute VB_Name = "��w�艺��"
Sub ��w�艺��M()
    Const �J�n�s As Long = 2

    Dim �Ώۗ� As Long
    Dim �ŏI�s As Long
    Dim ���݂̒l As String
    Dim i As Long

    ' ��ԍ������[�U�[�ɓ��͂�����
    �Ώۗ� = InputBox("��ԍ�����͂��Ă��������i��FA��Ȃ�1�AB��Ȃ�2�j")

    ' �����ȗ�ԍ������͂��ꂽ�ꍇ�A�������I��
    If �Ώۗ� < 1 Then Exit Sub

    ' �ŏI�s���擾
    �ŏI�s = Cells(Rows.count, �Ώۗ�).End(xlUp).row

    ' �����̌��݂̒l��ݒ�
    ���݂̒l = Cells(�J�n�s, �Ώۗ�).Value

    ' �����ߏ��������s
    For i = �J�n�s To �ŏI�s
        If Cells(i, �Ώۗ�).Value <> "" Then
            ���݂̒l = Cells(i, �Ώۗ�).Value
        End If
        Cells(i, �Ώۗ�).Value = ���݂̒l
    Next i
End Sub

