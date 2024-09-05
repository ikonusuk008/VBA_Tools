Attribute VB_Name = "エクセルテキスト取得"
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
    
    ' 入力シートと出力シートを設定
    Set ws = sheets("IN")
    Set outputWs = Worksheets.Add
    outputWs.name = "ExtractedText"
    outputRow = 1
    
    ' ヘッダーを追加
    outputWs.Cells(outputRow, 1).Value = "パス"
    outputWs.Cells(outputRow, 2).Value = "シート名"
    outputWs.Cells(outputRow, 3).Value = "列"
    outputWs.Cells(outputRow, 4).Value = "セルテキスト"
    outputRow = outputRow + 1
    
    ' ファイルパスのある最後の行を取得
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    ' 各ファイルを処理
    For i = 2 To lastRow
        filePath = ws.Cells(i, 1).Value
        
        ' エクセルファイルのみを処理
        If LCase(Right(filePath, 4)) = ".xls" Or LCase(Right(filePath, 5)) = ".xlsx" Then
            ' ファイルを開く
            Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
            
            ' 各シートを処理
            For Each sourceWs In sourceWb.Worksheets
                ' 各セルを処理
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
            
            ' ファイルを閉じる
            sourceWb.Close SaveChanges:=False
        End If
    Next i
    
    ' 列幅の自動調整
    outputWs.Columns("A:D").AutoFit
    
    MsgBox "テキスト抽出が完了しました。", vbInformation
End Sub
