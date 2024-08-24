Attribute VB_Name = "Excel�t�@�C��������"
' �O���[�o���ϐ��ōs�J�E���g���Ǘ�
Dim r As Long

Sub Excel�t�@�C��������M()
    Dim folderPath As String
    Dim outputWb As Workbook
    Dim outputWs As Worksheet
    
    ' �������F�ŏ��̃f�[�^�̏������݊J�n�s���w��
    r = 1
    
    ' �t�H���_�̃p�X���w��
    folderPath = "C:\YourFolderPath\" ' ��΃p�X���w�肵�Ă�������
    
    ' �o�͐�̃��[�N�u�b�N�ƃV�[�g���w��
    Set outputWb = Workbooks.Open("C:\YourOutputWorkbookPath\OutputWorkbook.xlsx") ' �o�͐�̃��[�N�u�b�N���w��
    Set outputWs = outputWb.sheets("Sheet1") ' �o�͐�̃V�[�g���w��
    
    ' �T�u�t�H���_�T�����ċA�I�ɏ������邽�߂ɁA���C���t�H���_����J�n
    Call ProcessFolderRecursively(folderPath, outputWs)
    
    MsgBox "�f�[�^�̌������������܂����B", vbInformation
End Sub

Sub �t�H���_���ċA�I�ɏ���(folderPath As String, outputWs As Worksheet)
    Dim fileName As String
    Dim FSO As Object
    Dim Folder As Object
    Dim SubFolder As Object

    ' FileSystemObject���g�p���ăt�H���_���̃t�@�C���ƃT�u�t�H���_������
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(folderPath)

    ' �t�H���_����Excel�t�@�C��������
    Call ProcessFilesInFolder(Folder.path, outputWs)

    ' �T�u�t�H���_���ċA�I�ɏ���
    For Each SubFolder In Folder.SubFolders
        Call ProcessFolderRecursively(SubFolder.path, outputWs)
    Next SubFolder
End Sub

Sub �t�H���_���̃t�@�C��������(folderPath As String, outputWs As Worksheet)
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetExists As Boolean

    ' �t�H���_����Excel�t�@�C�������ԂɊJ��
    fileName = Dir(folderPath & "\*.xls*")

    Do While fileName <> ""
        ' 'AAA' ���܂ރt�@�C�������`�F�b�N
        If InStr(fileName, "AAA") > 0 Then
            ' �t�@�C�����J��
            Set wb = Workbooks.Open(folderPath & "\" & fileName)
            sheetExists = False
            
            ' "AAA" �Ƃ������O�̃V�[�g�����݂��邩���m�F
            On Error Resume Next
            Set ws = wb.sheets("AAA")
            If Not ws Is Nothing Then
                sheetExists = True
            End If
            On Error GoTo 0

            ' �V�[�g�����݂���ꍇ�A�f�[�^������
            If sheetExists Then
                ' �f�[�^�̍ŏI�s���擾
                lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

                ' �e�s�����Ԃɏ�������
                For i = 2 To lastRow ' A2����ŏI�s�܂ł������i�w�b�_�[�s�͏����j
                    outputWs.Cells(r, 1).Value = ws.Cells(i, 1).Value ' ���Ԃ��R�s�[
                    outputWs.Cells(r, 2).Value = ws.Cells(i, 2).Value ' ���O���R�s�[
                    r = r + 1 ' ���̏������ݍs��
                Next i
            Else
                ' �V�[�g�����݂��Ȃ��ꍇ�A�t�@�C������ "�V�[�g����" ���L�^
                outputWs.Cells(r, 1).Value = fileName
                outputWs.Cells(r, 2).Value = "�V�[�g����"
                r = r + 1
            End If

            ' �t�@�C�������
            wb.Close SaveChanges:=False
        End If

        ' ���̃t�@�C����
        fileName = Dir
    Loop
End Sub
