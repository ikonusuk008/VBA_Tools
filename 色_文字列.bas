Attribute VB_Name = "�F_������"
' # Excel�}�N���œ���̕�����̐F�Ƒ�����ύX������@
' ## �͂��߂�
' ���̃u���O�L���ł́AExcel�V�[�g���̓���̕�����̐F��ύX���A�K�v�ɉ����đ����ɂ���}�N���̍쐬���@��������܂��B<br>���[�U�[�́A�ύX�������e�L�X�g�A�F�A����ё����ݒ����̓{�b�N�X��ʂ��Ďw�肵�܂��B<br>
' ## �F�Ƒ�����ύX����}�N��
' ���̃}�N���́A���[�U�[����̓��͂��󂯎��A�V�[�g���̓���̕�����̐F�Ƒ����ݒ��ύX���܂��B<br>
' <pre><code>
Sub �F�Ƒ�����ύX����()
Attribute �F�Ƒ�����ύX����.VB_ProcData.VB_Invoke_Func = "r\n14"
    Const �f�[�^�^ As Long = 23
    Dim �͈� As Range
    Dim �F�ύX�e�L�X�g As String
    Dim �F�C���f�b�N�X As String
    Dim �����ݒ� As String
    Dim rng As Range
    
    ' ���[�U�[�ɕύX�������e�L�X�g����͂�����
    �F�ύX�e�L�X�g = InputBox("�e�L�X�g����͂��Ă�������")
    
    ' ���[�U�[�ɐF�C���f�b�N�X����͂�����i�f�t�H���g��3�j
    �F�C���f�b�N�X = InputBox("�F����͂��Ă�������", "�F�ݒ�", "3")
    
    ' ���[�U�[�ɑ����ݒ����͂�����iB/b����͂���Ƒ����ɂȂ�j
    �����ݒ� = InputBox("�e�L�X�g�𑾎��ɂ��܂����H (B/b����͂��Ă�������)", "�����ݒ�", " ")
    
    ' �V�[�g���̒萔�Z����͈͂Ƃ��Đݒ�
    Set �͈� = ActiveSheet.Cells.SpecialCells(xlCellTypeConstants, �f�[�^�^)
    
    ' �͈͓��̊e�Z���ɑ΂��ĐF�ύX�����s
    For Each rng In �͈�
        Call �F�ύX(rng, �F�ύX�e�L�X�g, �F�C���f�b�N�X, �����ݒ�)
    Next rng
End Sub
' </code></pre>
' ## �F�ύX�T�u���[�`��
' ���̃T�u���[�`���́A�w�肳�ꂽ�͈͓��̓���̕�����̐F�Ƒ����ݒ��ύX���܂��B<br>
' <pre><code>
Sub �F�ύX(rng As Range, �F�ύX�e�L�X�g As String, �F�C���f�b�N�X As String, �����ݒ� As String)
    Dim �|�C���^ As Integer
    
    ' �w�肳�ꂽ�e�L�X�g�̈ʒu������
    �|�C���^ = InStr(rng.Value, �F�ύX�e�L�X�g)
    
    ' �e�L�X�g����������胋�[�v
    While �|�C���^ > 0
        ' �w�肳�ꂽ�e�L�X�g�̐F��ύX
        rng.Characters(Start:=�|�C���^, Length:=Len(�F�ύX�e�L�X�g)).Font.colorIndex = CInt(�F�C���f�b�N�X)
        
        ' �����ݒ肪"B"�̏ꍇ�A�����ɂ���
        If UCase(�����ݒ�) = "B" Then
            rng.Characters(Start:=�|�C���^, Length:=Len(�F�ύX�e�L�X�g)).Font.Bold = True
        End If
        
        ' ���̈ʒu������
        �|�C���^ = InStr(�|�C���^ + Len(�F�ύX�e�L�X�g), rng.Value, �F�ύX�e�L�X�g)
    Wend
End Sub
' </code></pre>

