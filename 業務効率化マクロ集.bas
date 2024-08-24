Attribute VB_Name = "�Ɩ��������}�N���W"
' ���݂̃Z���̍s���폜����}�N��
Sub One_row_delete()
    ' ���݂̃Z���̍s��I��
    ActiveCell.EntireRow.Select
    ' �I�������s���폜
    Rows(ActiveCell.row).Delete
End Sub

' ���݂̃Z���̍s�ɐV�����s��}������}�N��
Sub One_row_insert()
    ' ���݂̃Z���̍s��I��
    ActiveCell.EntireRow.Select
    ' �I�������s�ɐV�����s��}��
    Rows(ActiveCell.row).Insert
End Sub

' ���݂̃Z���̗���폜����}�N��
Sub One_Column_delete()
    ' ���݂̃Z���̗��I��
    ActiveCell.EntireColumn.Select
    ' �I����������폜
    Columns(ActiveCell.Column).Delete
End Sub

' ���݂̃Z���̗�ɐV�������}������}�N��
Sub One_Column_insert()
    ' ���݂̃Z���̗��I��
    ActiveCell.EntireColumn.Select
    ' �I��������ɐV�������}��
    Columns(ActiveCell.Column).Insert
End Sub

' ���ׂĂ̍s�̍�����������������}�N��
Sub Automatic_row_height_adjustment()
    ' ���ׂẴZ����I��
    Cells.Select
    ' ���ׂĂ̍s�̍�������������
    Cells.EntireRow.AutoFit
    ' �Z��A1��I���i�I����Ԃ��������邽�߁j
    Range("A1").Select
End Sub

' �t�H���g�Ɠ��t�񕝂̒������s���}�N��
Sub �J�����̏k�ڒ���()
    ' ���ׂẴZ����I��
    Cells.Select
    ' �񕝂�ݒ�
    Selection.ColumnWidth = 8.29
    ' ���ׂĂ̗�̕�����������
    Cells.EntireColumn.AutoFit
    ' �Z��A1��I���i�I����Ԃ��������邽�߁j
    Range("A1").Select
End Sub

' ���݂̃Z���̉��s�Ƌ󔒂��폜����}�N��
Sub Eliminate_line_breaks_and_blanks_in_the_active_Cell()
    ' ���݂̃Z���̒l���擾
    a = ActiveCell.Value
    ' ���s���폜
    b = Replace(a, vbLf, "")
    ' �󔒂��폜
    b = Replace(b, " ", "")
    ' �C�������l�����݂̃Z���ɐݒ�
    ActiveCell.Value = b
End Sub

' ���݂̃Z���̉��̃Z���̒l���R�s�[����}�N��
Sub Copy_cells_below_active_cell()
    ' ���݂̃Z���ɉ��̃Z���̒l��ݒ�
    ActiveCell.Cells(1, 1).Value = ActiveCell.Cells(2, 1).Value
End Sub

' ���ׂẴV�[�g�����Z���ɕ\������}�N��
Sub �V�[�g���擾�G�N�Z���֐�()
    For i1 = 1 To Worksheets.count - 1
        ' �e�V�[�g��I��
        Worksheets.Select (i1)
        ' B2�Z����I��
        Range("B2").Select
        ' B2�Z�����N���A
        ActiveCell.FormulaR1C1 = ""
        ' A1�Z���ɃV�[�g����ݒ�
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "=RIGHT(CELL(""filename"",RC),LEN(CELL(""filename"",RC))-FIND(""]"",CELL(""filename"",RC)))"
        ' A2�Z����I������1�s�ڂ̍�����ݒ�
        Range("A2").Select
        Rows("1:1").RowHeight = 23.25
    Next i1
End Sub

' ���ׂẴV�[�g�����f�o�b�O�o�͂���}�N��
Sub �V�[�g���̊m�F()
    For i1 = 1 To Worksheets.count - 1
        ' �e�V�[�g�̖��O���f�o�b�O�o��
        Debug.Print Worksheets(i1).name
    Next i1
End Sub
