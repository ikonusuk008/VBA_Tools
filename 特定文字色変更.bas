Attribute VB_Name = "���蕶���F�ύX"
Sub ���蕶���F�ύXM()

    Dim rng As Range
    Dim ptr As Integer

    '23: "Numeric value", "Character", "Logical value", and "Error value" are all selected
    Const dataType As Long = 23

    Dim colorChangeText As String
    colorChangeText = InputBox("�F��ύX����e�L�X�g����͂��Ă�������")

    Dim colorIndex As String
    colorIndex = InputBox("�F�ԍ�����͂��Ă�������", "�F�ݒ�", "3")

    Dim boldSetting As String
    boldSetting = InputBox("�����ɂ��܂����H (B/b�ő���)", "�����ݒ�", " ")

    For Each rng In ActiveSheet.Cells.SpecialCells(xlCellTypeConstants, dataType)
        ptr = InStr(rng.Value, colorChangeText)

        ' While���[�v�ŃZ�����̕���������ׂČ�����
        While ptr > 0

            rng.Characters(Start:=ptr, Length:=Len(colorChangeText)).Font.colorIndex = CInt(colorIndex)

            If UCase(boldSetting) = "B" Then
                rng.Characters(Start:=ptr, Length:=Len(colorChangeText)).Font.Bold = True
            End If

            ' ���̈�v������
            ptr = InStr(ptr + Len(colorChangeText), rng.Value, colorChangeText)

        Wend
    Next rng

End Sub
