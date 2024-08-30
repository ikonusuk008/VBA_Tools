Attribute VB_Name = "�ڎ��쐬_���΃p�X"
' ================================
' �T�}��
' �쐬��: �{��
' �w�i: Excel�V�[�g�̖ڎ��������������邽�߂̃}�N��
' �ړI: �S���[�N�V�[�g���ꗗ�ɂ��Ėڎ����쐬���A�n�C�p�[�����N��ݒ�
' ���L����: ����̃V�[�g���i"�y" �܂��� "��" �Ŏn�܂�j�̏ꍇ�A�V������Ɉړ����郍�W�b�N���܂�
' ================================

' �ڎ��V�[�g���쐬����T�u���[�`��
Sub �ڎ��쐬_���΃p�XM()
Attribute �ڎ��쐬_���΃p�XM.VB_ProcData.VB_Invoke_Func = "U\n14"

    ' ���[�U�[�Ɋg�嗦����͂�����_�C�A���O��\��
    Dim userMagnification As Variant
    Dim magnification As Integer ' �����Ŋg�嗦�̕ϐ���錾

    userMagnification = InputBox("�\���{������͂��Ă������� (��: 100)", "�\���{���ݒ�", "90")
    
    ' ���͂����l���ǂ������m�F
    If IsNumeric(userMagnification) Then
        magnification = CInt(userMagnification) ' ���͂��ꂽ�g�嗦��ݒ�
    Else
        MsgBox "�L���Ȑ��l����͂��Ă��������B", vbExclamation
        Exit Sub
    End If




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
        Application.DisplayAlerts = False ' �폜�̌x����\�����Ȃ�
        ActiveSheet.Delete ' ������INDEX�V�[�g���폜
        Application.DisplayAlerts = True ' �폜�̌x�������ɖ߂�
    End If

    ' �V�������[�N�V�[�g��ǉ����A�����"INDEX"�Ɩ���
    Worksheets.Add
    ActiveSheet.name = "INDEX"

    ' �p�t�H�[�}���X����̂��߂̐ݒ���Ăяo��
    accelerate
    
    ' �S���[�N�V�[�g�����[�v���Ėڎ����쐬
    For i = 1 To Worksheets.count
        worksheetName = Worksheets(i).name
        
        ' ���[�N�V�[�g��������̕����Ŏn�܂�ꍇ�A�V������Ɉړ�
        If Mid(worksheetName, 1, 1) = "�y" Or Mid(worksheetName, 1, 1) = "��" Then
            row_num = 1
            col_num = col_num + 1
        End If
        
        ' �n�C�p�[�����N�̃T�u�A�h���X���쐬�i���΃p�X�j
        Dim subAddress_ As String
        subAddress_ = "'" & worksheetName & "'!A1" ' �n�C�p�[�����N�̑��΃p�X��ݒ�
        
        ' INDEX�V�[�g�̓K�؂ȃZ���Ɉړ����ăn�C�p�[�����N��ǉ�
        With Worksheets("INDEX").Cells(row_num, col_num)
            .Hyperlinks.Add Anchor:=.Cells(row_num, col_num), Address:="", SubAddress:=subAddress_, TextToDisplay:=worksheetName
            .Value = worksheetName
            .Interior.colorIndex = Worksheets(i).Tab.colorIndex
            .EntireColumn.ColumnWidth = 30
        End With
        
        row_num = row_num + 1 ' �s�ԍ��𑝂₷
    Next i
    
    ' �p�t�H�[�}���X�ݒ�����ɖ߂�
    clearAccelerate
     
    ' INDEX�V�[�g�̗񕝂𒲐�
    Columns("B:BB").ColumnWidth = 10
    Columns("A:A").EntireColumn.AutoFit
    Cells.EntireColumn.AutoFit
    Cells(1, 1).Select
    
    ' �E�B���h�E�𕪊����ČŒ�
    FreezePanes

    ' ���ׂẴ��[�N�V�[�g�̕\���{����ݒ�
    For i = 1 To Worksheets.count
        Worksheets(i).Select
        ActiveWindow.Zoom = magnification
    Next i
    
    ' �ŏ��̃V�[�g��I�����ďI��
    Worksheets(1).Select
End Sub

' �E�B���h�E�𕪊����ČŒ肷��T�u���[�`��
Sub FreezePanes()
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub

' ���[�N�V�[�g�����݂��邩�m�F����֐�
Function ExistsWorksheet(ByVal name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In sheets
        If ws.name = name Then
            ExistsWorksheet = True
            Exit Function
        End If
    Next
    ExistsWorksheet = False
End Function

' �p�t�H�[�}���X����̂��߂̐ݒ���s���T�u���[�`��
Sub accelerate()
    With Application
        .ScreenUpdating = False ' ��ʂ̍X�V���~
        .DisplayAlerts = False ' �x����\�����Ȃ�
        .EnableEvents = False ' �C�x���g�𖳌���
        .Calculation = xlCalculationManual ' �蓮�v�Z���[�h�ɐݒ�
    End With
End Sub

' �p�t�H�[�}���X�ݒ�����ɖ߂��T�u���[�`��
Sub clearAccelerate()
    With Application
        .ScreenUpdating = True ' ��ʂ̍X�V���ĊJ
        .DisplayAlerts = True ' �x����\��
        .EnableEvents = True ' �C�x���g��L����
        .Calculation = xlCalculationAutomatic ' �����v�Z���[�h�ɐݒ�
    End With
End Sub


