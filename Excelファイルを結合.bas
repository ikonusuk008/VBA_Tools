Attribute VB_Name = "Excelファイルを結合"
' グローバル変数で行カウントを管理
Dim r As Long

Sub Excelファイルを結合M()
    Dim folderPath As String
    Dim outputWb As Workbook
    Dim outputWs As Worksheet
    
    ' 初期化：最初のデータの書き込み開始行を指定
    r = 1
    
    ' フォルダのパスを指定
    folderPath = "C:\YourFolderPath\" ' 絶対パスを指定してください
    
    ' 出力先のワークブックとシートを指定
    Set outputWb = Workbooks.Open("C:\YourOutputWorkbookPath\OutputWorkbook.xlsx") ' 出力先のワークブックを指定
    Set outputWs = outputWb.sheets("Sheet1") ' 出力先のシートを指定
    
    ' サブフォルダ探索も再帰的に処理するために、メインフォルダから開始
    Call ProcessFolderRecursively(folderPath, outputWs)
    
    MsgBox "データの結合が完了しました。", vbInformation
End Sub

Sub フォルダを再帰的に処理(folderPath As String, outputWs As Worksheet)
    Dim fileName As String
    Dim FSO As Object
    Dim Folder As Object
    Dim SubFolder As Object

    ' FileSystemObjectを使用してフォルダ内のファイルとサブフォルダを処理
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(folderPath)

    ' フォルダ内のExcelファイルを処理
    Call ProcessFilesInFolder(Folder.path, outputWs)

    ' サブフォルダを再帰的に処理
    For Each SubFolder In Folder.SubFolders
        Call ProcessFolderRecursively(SubFolder.path, outputWs)
    Next SubFolder
End Sub

Sub フォルダ内のファイルを処理(folderPath As String, outputWs As Worksheet)
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetExists As Boolean

    ' フォルダ内のExcelファイルを順番に開く
    fileName = Dir(folderPath & "\*.xls*")

    Do While fileName <> ""
        ' 'AAA' を含むファイル名をチェック
        If InStr(fileName, "AAA") > 0 Then
            ' ファイルを開く
            Set wb = Workbooks.Open(folderPath & "\" & fileName)
            sheetExists = False
            
            ' "AAA" という名前のシートが存在するかを確認
            On Error Resume Next
            Set ws = wb.sheets("AAA")
            If Not ws Is Nothing Then
                sheetExists = True
            End If
            On Error GoTo 0

            ' シートが存在する場合、データを処理
            If sheetExists Then
                ' データの最終行を取得
                lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

                ' 各行を順番に書き込む
                For i = 2 To lastRow ' A2から最終行までを処理（ヘッダー行は除く）
                    outputWs.Cells(r, 1).Value = ws.Cells(i, 1).Value ' 時間をコピー
                    outputWs.Cells(r, 2).Value = ws.Cells(i, 2).Value ' 名前をコピー
                    r = r + 1 ' 次の書き込み行へ
                Next i
            Else
                ' シートが存在しない場合、ファイル名と "シート無し" を記録
                outputWs.Cells(r, 1).Value = fileName
                outputWs.Cells(r, 2).Value = "シート無し"
                r = r + 1
            End If

            ' ファイルを閉じる
            wb.Close SaveChanges:=False
        End If

        ' 次のファイルへ
        fileName = Dir
    Loop
End Sub
