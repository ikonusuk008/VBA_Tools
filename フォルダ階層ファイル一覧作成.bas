Attribute VB_Name = "�t�H���_�K�w�t�@�C���ꗗ�쐬"
Sub �t�H���_�K�w�t�@�C���ꗗ�쐬M()

    Dim mainFolder As String
    Dim subFolder1 As Object
    Dim subFolder2 As Object
    Dim subFolder3 As Object
    Dim file As Object
    Dim fileSystem As Object
    Dim currentRow As Long
    Dim maxExcelCount As Long, maxTextCount As Long, maxCsvCount As Long, maxGrandchildCount As Long
    Dim excelFileCount As Long, textFileCount As Long, csvFileCount As Long, grandchildFileCount As Long
    Dim excelFiles() As String, textFiles() As String, csvFiles() As String, grandchildFiles() As String
    
    ' ���C���t�H���_�̃p�X���w��
    mainFolder = "C:\path\to\your\folder" ' �t�H���_�̃p�X���w��
    
    ' �t�@�C���V�X�e���I�u�W�F�N�g���쐬
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' �����ݒ�
    currentRow = 2 ' �\�̏������݊J�n�s�i1�s�ڂɃw�b�_�[�j
    
    ' �T�u�t�H���_2����т��̔z���̃t�H���_�i�T�u�t�H���_3�j���̍ő�t�@�C�������擾
    maxExcelCount = 0
 