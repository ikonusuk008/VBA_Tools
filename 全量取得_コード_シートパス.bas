Attribute VB_Name = "全量取得_コード_シートパス"
Sub 全量取得_コード_シートパスM()
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

    ' 特定のブックを設定（ここでは、現在アクティブなブックを対象とします）
    Set targetWorkbook = Application.ActiveWorkbook

    ' テストフラグを設定
    testFlag = True ' テストフラグがTrueの場合、2行だけ処理

    ' 出力シートを作成または取得
    On Error Resume Next
    Set wsOutput = targetWorkbook.sheets("VBA全量")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Set wsOutput = targetWorkbook.sheets.Add(After:=targetWorkbook.sheets(targetWorkbook.sheets.count))
        wsOutput.name = "VBA全量"
    End If
    
    ' 2行目以降をクリア（フィルタを維持）
    wsOutput.Rows("2:" & wsOutput.Rows.count).ClearContents
    
    ' カラム名を設定（既にある場合でも設定し直します）
    wsOutput.Cells(1, "A").Value = "path filename"
    wsOutput.Cells(1, "B").Value = "filename"
    wsOutput.Cells(1, "C").Value = "last modified"
    wsOutput.Cells(1, "D").Value = "code"
    wsOutput.Cells(1, "E").Value = "subroutine"
    wsOutput.Cells(1, "F").Value = "line number"
    wsOutput.Cells(1, "G").Value = "overall line number"

    ' フィルタ設定（フィルタを維持）
    If wsOutput.AutoFilterMode Then
        wsOutput.AutoFilterMode = False
    End If
    wsOutput.Rows(1).AutoFilter

    ' 1行目を固定
    wsOutput.Activate
    wsOutput.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' ファイルパスが記載されたシートを指定
    Set wsSource = targetWorkbook.sheets("VBA_path")
    
    ' 最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.count, "B").End(xlUp).row
    
    ' テストフラグがTrueなら、処理行数を制限
    If testFlag Then
        lastRow = WorksheetFunction.Min(lastRow, 2 + 1) ' ヘッダ行があるため、2行の場合は実質3行目まで処理
    End If
    
    ' 初期化
    outputRow = 2 ' カラム名の下から出力
    overallLineNumber = 1 ' 全体の連番を初期化

    ' 2行目からループ
    For i = 2 To lastRow
        filePath = wsSource.Cells(i, "B").Value ' B列からファイルパスを取得
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1) ' ファイル名を取得
        lastModified = FileDateTime(filePath) ' ファイルの最終更新日を取得
        Debug.Print "Processing file path: " & filePath ' ファイルパスをデバッグ出力

        If filePath <> "" And Right(filePath, 4) = ".bas" Then
            ' ファイルを読み込む
            fileNum = FreeFile
            On Error Resume Next
            Open filePath For Input As fileNum
            If Err.Number <> 0 Then
                MsgBox "ファイルを開けませんでした: " & filePath, vbExclamation
                On Error GoTo 0
                Close fileNum
                GoTo NextFile
            End If
            On Error GoTo 0
            
            ' 初期化
            inSubroutine = False
            subroutineName = ""
            lineNumber = 1

            ' ファイルの内容を1行ずつ出力
            Do Until EOF(fileNum)
                Line Input #fileNum, lineData

                ' サブルーチンの開始を検出
                If InStr(1, lineData, "Sub ", vbTextCompare) > 0 Or InStr(1, lineData, "Function ", vbTextCompare) > 0 Then
                    subroutineName = Trim(Split(lineData, " ")(1))
                    inSubroutine = True
                End If

                ' サブルーチンの終了を検出
                If inSubroutine And (InStr(1, lineData, "End Sub", vbTextCompare) > 0 Or InStr(1, lineData, "End Function", vbTextCompare) > 0) Then
                    inSubroutine = False
                End If

                ' 出力
                wsOutput.Cells(outputRow, "A").Value = filePath ' A列にフルパスを出力
                wsOutput.Cells(outputRow, "B").Value = fileName ' B列にファイル名を出力
                wsOutput.Cells(outputRow, "C").Value = lastModified ' C列に最終更新日を出力
                wsOutput.Cells(outputRow, "D").Value = lineData ' D列に1行ずつコードを出力
                wsOutput.Cells(outputRow, "E").Value = subroutineName ' E列にサブルーチン名を出力（該当行がサブルーチン内の場合）
                wsOutput.Cells(outputRow, "F").Value = lineNumber ' F列にファイル内の連番を出力
                wsOutput.Cells(outputRow, "G").Value = overallLineNumber ' G列に全体の連番を出力

                outputRow = outputRow + 1
                lineNumber = lineNumber + 1
                overallLineNumber = overallLineNumber + 1
            Loop
            Close fileNum
            
            Debug.Print "File content output complete for: " & filePath
        Else
            Debug.Print "Skipping file: " & filePath ' 条件を満たさないファイルはスキップ
        End If
NextFile:
    Next i
    
    MsgBox "処理が完了しました。", vbInformation
End Sub
