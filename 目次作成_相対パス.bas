Attribute VB_Name = "�ڎ��쐬_���΃p�X"
Const magnification = 90 ' �\���{����90%�ɐݒ�

' ================================
' �T�}��
' �쐬��: �{��
' �w�i: Excel�V�[�g�̖ڎ��������������邽�߂̃}�N��
' �ړI: �S���[�N�V�[�g���ꗗ�ɂ��Ėڎ����쐬���A�n�C�p�[�����N��ݒ�
' ���L����: ����̃V�[�g���i"�y" �܂��� "��" �Ŏn�܂�j�̏ꍇ�A�V������Ɉړ����郍�W�b�N���܂�
' ================================

' �ڎ��V�[�g���쐬����T�u���[�`��
Sub �ڎ��쐬_���΃p�XM()
    Dim sheetCount As Integer ' ���[�N�V�[�g�̐�
    Dim worksheetName As String ' ���[�N�V�[�g�̖��O
    Dim Worksheet As Worksheet ' ���[�N�V�[�g�I�u�W�F�N�g
    Dim col_num As Integer ' ��ԍ�
    col_num = 1 ' ������ԍ���1�ɐݒ�
    Dim j As Integer ' ���[�v�J�E���^
    Dim row_num As Integer ' �s�ԍ�
    row_num = 1 ' �����s�ԍ���1�ɐݒ�
    
    ' ����"INDEX"�Ƃ������O�̃��[�N�V�[�g�����݂���ꍇ�A������폜
    If ExistsWorksheet("INDEX") Then
        Worksheets("INDEX").Select
        Application.DisplayAlerts = False ' �폜