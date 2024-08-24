Attribute VB_Name = "�S�ʎ擾_�R�[�h_�V�[�g�p�X"
Sub �S�ʎ擾_�R�[�h_�V�[�g�p�XM()
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim filePath As String
    Dim fileName As String
    Dim lastModified As Date
    Dim i As Long
    Dim fileContent As String
    Dim fileNum As Integer
    Dim testFlag As Boolean
    Dim lineData As String
    Dim outputRow As Long
    Dim targetWorkbook As Workbook
    Dim subroutineName As String
    Dim inSubroutine As Boolean
    Dim lineNumber As Long
    Dim overallLineNumber As Long

    ' ����̃u�b�N��ݒ�i�����ł́A���݃A�N�e�B�u�ȃu�b�N��ΏۂƂ��܂��j
    Set targetWorkbook = Application.ActiveWorkbook

    ' �e�X�g�t���O��ݒ�
    testFlag = True ' �e�X�g�t���O��True�̏ꍇ�A2�s��������

    ' �o�̓V�[�g���쐬�܂��͎擾
    On Error Resume Next
    Set wsOutput = targetWorkbook.sheets("VBA�S��")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Set wsOutput = targetWorkbook.sheets.Add(After:=targetWorkbook.sheets(targetWorkbook.sheets.count))
        wsOutput.name = "VBA�S��"
    End If
    
    ' 2�s�ڈȍ~���N���A�i�t�B���^���ێ��j
    wsOutput.Rows("2:" & wsOutput.Rows.count).ClearContents
    
    ' �J��������ݒ�i���ɂ���ꍇ�ł��ݒ肵�����܂��j
    wsOutput.Cells(1, "A").Value = "path filename"
    wsOutput.Cells(1, "B").Value = "filename"
    wsOutput.Cells(1, "C").Value = "last modified"
    wsOutput.Cells(1, "D").Value = "code"
    wsOutput.Cells(1, "E").Value = "subroutine"
    wsOutput.Cells(1, "F").Value = "line number"
    wsOutput.Cells(1, "G").Value = "overall line number"

    ' �t�B���^�ݒ�i�t�B���^���ێ��j
    If wsOutput.AutoFilterMode Then
        wsOutput.AutoFilterMode = False
    End If
    wsOutput.Rows(1).AutoFilter

    ' 1�s�ڂ��Œ�
    wsOutput.Activate
    wsOutput.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' �t�@�C���p�X���L�ڂ��ꂽ�V�[�g���w��
    Set wsSource = targetWorkbook.sheets("VBA_path")
    
    ' �ŏI�s���擾
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).row
    
    ' �e�X�g�t���O��True�Ȃ�A�����s���𐧌�
    If testFlag Then
        lastRow = WorksheetFunction.Min(lastRow, 2 + 1) ' �w�b�_�s�����邽�߁A2�s�̏ꍇ�͎���3�s�ڂ܂ŏ���
    End If
    
    ' ������
    outputRow = 2 ' �J�������̉�����o��
    overallLineNumber = 1 ' �S�̘̂A�Ԃ�������

    ' 2�s�ڂ��烋�[�v
    For i = 2 To lastRow
        filePath = wsSource.Cells(i, "B").Value ' B�񂩂�t�@�C���p�X���擾
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1) ' �t�@�C�������擾
        lastModified = FileDateTime(filePath) ' �t�@�C���̍ŏI�X�V�����擾
        Debug.Print "Processing file path: " & filePath ' �t�@�C���p�X���f�o�b�O�o��

        If filePath <> "" And Right(filePath, 4) = ".bas" Then
            ' �t�@�C����ǂݍ���
            fileNum = FreeFile
            On Error Resume Next
            Open filePath For Input As fileNum
            If Err.Number <> 0 Then
                MsgBox "�t�@�C�����J���܂���ł���: " & filePath, vbExclamation
                On Error GoTo 0
                Close fileNum
                GoTo NextFile
            End If
            On Error GoTo 0
            
            ' ������
            inSubroutine = False
            subroutineName = ""
            lineNumber = 1

            ' �t�@�C���̓��e��1�s���o��
            Do Until EOF(fileNum)
                Line Input #fileNum, lineData

                ' �T�u���[�`���̊J�n�����o
                If InStr(1, lineData, "Sub ", vbTextCompare) > 0 Or InStr(1, lineData, "Function ", vbTextCompare) > 0 Then
                    subroutineName = Trim(Split(lineData, " ")(1))
                    inSubroutine = True
                End If

                ' �T�u���[�`���̏I�������o
                If inSubroutine And (InStr(1, lineData, "End Sub", vbTextCompare) > 0 Or InStr(1, lineData, "End Function", vbTextCompare) > 0) Then
                    inSubroutine = False
                End If

                ' �o��
                wsOutput.Cells(outputRow, "A").Value = filePath ' A��Ƀt���p�X���o��
                wsOutput.Cells(outputRow, "B").Value = fileName ' B��Ƀt�@�C�������o��
                wsOutput.Cells(outputRow, "C").Value = lastModified ' C��ɍŏI�X�V�����o��
                wsOutput.Cells(outputRow, "D").Value = lineData ' D���1�s���R�[�h���o��
                wsOutput.Cells(outputRow, "E").Value = subroutineName ' E��ɃT�u���[�`�������o�́i�Y���s���T�u���[�`�����̏ꍇ�j
                wsOutput.Cells(outputRow, "F").Value = lineNumber ' F��Ƀt�@�C�����̘A�Ԃ��o��
                wsOutput.Cells(outputRow, "G").Value = overallLineNumber ' G��ɑS�̘̂A�Ԃ��o��

                outputRow = outputRow + 1
                lineNumber = lineNumber + 1
                overallLineNumber = overallLineNumber + 1
            Loop
            Close fileNum
            
            Debug.Print "File content output complete for: " & filePath
        Else
            Debug.Print "Skipping file: " & filePath ' �����𖞂����Ȃ��t�@�C���̓X�L�b�v
        End If
NextFile:
    Next i
    
    MsgBox "�������������܂����B", vbInformation
End Sub
