Attribute VB_Name = "�S�ʎ擾_�R�[�h_�t�H���_�w��"
Sub �S�ʎ擾_�R�[�h_�t�H���_�w��M()
    Dim wsOutput As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim filePath As String
    Dim lastModified As Date
    Dim fileContent As String
    Dim fileNum As Integer
    Dim lineData As String
    Dim outputRow As Long
    Dim subroutineName As String
    Dim inSubroutine As Boolean
    Dim lineNumber As Long
    Dim overallLineNumber As Long
    Dim targetWorkbook As Workbook

    ' ����̃u�b�N��ݒ�i�����ł́A���݃A�N�e�B�u�ȃu�b�N��ΏۂƂ��܂��j
    Set targetWorkbook = Application.ActiveWorkbook

    ' �t�H���_�p�X���w��
    folderPath = InputBox("�t�H���_�p�X����͂��Ă�������:", "�t�H���_�I��")
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' �o�̓V�[�g���쐬�܂��͎擾
    On Error Resume Next
    Set wsOutput = targetWorkbook.sheets("VBA�S��")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Se