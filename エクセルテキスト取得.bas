Attribute VB_Name = "�G�N�Z���e�L�X�g�擾"
Sub ExtractTextFromExcelFiles()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim outputWs As Worksheet
    Dim outputRow As Long
    Dim cell As Range
    Dim columnLetter As String
    
    ' ���̓V�[�g�Əo�̓V�[�g��ݒ�
    Set ws = sheets("IN")
    Set outputWs = Worksheets.Add
    outputWs.name = "ExtractedText"
    outputRow = 1
    
    ' �w�b�_�[��ǉ�
    outputWs.Cells(outputRow, 1).Value = "�p�X"
    outputWs.Cells(outputRow, 2).Value = "�V�[�g��"
    outputWs.Cells(outputRow, 3).Value = "��"
    outputWs.Cells(outputRow, 4).Value = "�Z���e�L�X�g"
    outputRow = outputRow + 1
    
    ' �t�@�C���p�X�̂���Ō�̍s���擾
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    ' �e�t�@�C��������
    For i = 2 To lastRow
        filePath = ws.Cells(i, 1).Value
        
        ' �G�N�Z���t�@�C���݂̂�����
        If LCase(Right(filePath, 4)) = ".xls" Or LCase(Right(filePath, 5)) = ".xlsx" Then
            ' �t�@�C�����J��
            Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
            
            ' �e�V�[�g������
            For Each sourceWs In sourceWb.Worksheets
                ' �e�Z��������
                For Each cell In sourceWs.UsedRange
                    If cell.Value <> "" Then
                        columnLetter = Split(cell.Address, "$")(1)
                        outputWs.Cells(outputRow, 1).Value = filePath
                        outputWs.Cells(outputRow, 2).Value = sourceWs.name
                        outputWs.Cells(outputRow, 3).Value = columnLetter
                        outputWs.Cells(outputRow, 4).Value = cell.Value
                        outputRow = outputRow + 1
                    End If
                Next cell
            Next sourceWs
            
            ' �t�@�C�������
            sourceWb.Close SaveChanges:=False
        End If
    Next i
    
    ' �񕝂̎�������
    outputWs.Columns("A:D").AutoFit
    
    MsgBox "�e�L�X�g���o���������܂����B", vbInformation
End Sub
