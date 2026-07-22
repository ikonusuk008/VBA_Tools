Attribute VB_Name = "フォルダ階層ファイル一覧作成"
Sub フォルダ階層ファイル一覧作成M()

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
    Dim baseColumn As Long
    Dim i As Long

    ' メインフォルダのパスを指定
    mainFolder = "C:\path\to\your\folder" ' フォルダのパスを指定

    ' ファイルシステムオブジェクトを作成
    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    ' 初期設定
    currentRow = 2 ' 表の書き込み開始行（1行目にヘッダー）

    ' サブフォルダ2およびその配下のフォルダ（サブフォルダ3）内の最大ファイル数を取得
    maxExcelCount = 0
    maxTextCount = 0
    maxCsvCount = 0
    maxGrandchildCount = 0

    For Each subFolder1 In fileSystem.GetFolder(mainFolder).SubFolders
        For Each subFolder2 In subFolder1.SubFolders
            excelFileCount = 0
            textFileCount = 0
            csvFileCount = 0
            grandchildFileCount = 0

            ' サブフォルダ2直下のファイルを種類別に数える
            For Each file In subFolder2.Files
                Select Case LCase(fileSystem.GetExtensionName(file.name))
                    Case "xls", "xlsx", "xlsm"
                        excelFileCount = excelFileCount + 1
                    Case "txt"
                        textFileCount = textFileCount + 1
                    Case "csv"
                        csvFileCount = csvFileCount + 1
                End Select
            Next file

            ' サブフォルダ3（孫フォルダ）直下のファイルを数える
            For Each subFolder3 In subFolder2.SubFolders
                grandchildFileCount = grandchildFileCount + subFolder3.Files.count
            Next subFolder3

            If excelFileCount > maxExcelCount Then maxExcelCount = excelFileCount
            If textFileCount > maxTextCount Then maxTextCount = textFileCount
            If csvFileCount > maxCsvCount Then maxCsvCount = csvFileCount
            If grandchildFileCount > maxGrandchildCount Then maxGrandchildCount = grandchildFileCount
        Next subFolder2
    Next subFolder1

    ' ヘッダーを書き込む（1行目）
    Cells(1, 1).Value = "サブフォルダ1"
    Cells(1, 2).Value = "サブフォルダ2"
    baseColumn = 3
    For i = 1 To maxExcelCount
        Cells(1, baseColumn + i - 1).Value = "Excel" & i
    Next i
    baseColumn = baseColumn + maxExcelCount
    For i = 1 To maxTextCount
        Cells(1, baseColumn + i - 1).Value = "テキスト" & i
    Next i
    baseColumn = baseColumn + maxTextCount
    For i = 1 To maxCsvCount
        Cells(1, baseColumn + i - 1).Value = "CSV" & i
    Next i
    baseColumn = baseColumn + maxCsvCount
    For i = 1 To maxGrandchildCount
        Cells(1, baseColumn + i - 1).Value = "サブフォルダ3ファイル" & i
    Next i

    ' 本体を書き込む（2行目以降：サブフォルダ2ごとに1行）
    For Each subFolder1 In fileSystem.GetFolder(mainFolder).SubFolders
        For Each subFolder2 In subFolder1.SubFolders
            excelFileCount = 0
            textFileCount = 0
            csvFileCount = 0
            grandchildFileCount = 0

            If maxExcelCount > 0 Then ReDim excelFiles(1 To maxExcelCount)
            If maxTextCount > 0 Then ReDim textFiles(1 To maxTextCount)
            If maxCsvCount > 0 Then ReDim csvFiles(1 To maxCsvCount)
            If maxGrandchildCount > 0 Then ReDim grandchildFiles(1 To maxGrandchildCount)

            ' サブフォルダ2直下のファイル名を種類別に取得
            For Each file In subFolder2.Files
                Select Case LCase(fileSystem.GetExtensionName(file.name))
                    Case "xls", "xlsx", "xlsm"
                        excelFileCount = excelFileCount + 1
                        excelFiles(excelFileCount) = file.name
                    Case "txt"
                        textFileCount = textFileCount + 1
                        textFiles(textFileCount) = file.name
                    Case "csv"
                        csvFileCount = csvFileCount + 1
                        csvFiles(csvFileCount) = file.name
                End Select
            Next file

            ' サブフォルダ3（孫フォルダ）直下のファイル名を取得
            For Each subFolder3 In subFolder2.SubFolders
                For Each file In subFolder3.Files
                    grandchildFileCount = grandchildFileCount + 1
                    grandchildFiles(grandchildFileCount) = subFolder3.name & "\" & file.name
                Next file
            Next subFolder3

            ' 1行分を書き込む
            Cells(currentRow, 1).Value = subFolder1.name
            Cells(currentRow, 2).Value = subFolder2.name
            baseColumn = 3
            For i = 1 To excelFileCount
                Cells(currentRow, baseColumn + i - 1).Value = excelFiles(i)
            Next i
            baseColumn = baseColumn + maxExcelCount
            For i = 1 To textFileCount
                Cells(currentRow, baseColumn + i - 1).Value = textFiles(i)
            Next i
            baseColumn = baseColumn + maxTextCount
            For i = 1 To csvFileCount
                Cells(currentRow, baseColumn + i - 1).Value = csvFiles(i)
            Next i
            baseColumn = baseColumn + maxCsvCount
            For i = 1 To grandchildFileCount
                Cells(currentRow, baseColumn + i - 1).Value = grandchildFiles(i)
            Next i

            currentRow = currentRow + 1
        Next subFolder2
    Next subFolder1

    MsgBox "ファイル一覧の作成が完了しました。", vbInformation

End Sub
