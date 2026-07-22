Attribute VB_Name = "全量取得_コード_フォルダ指定"
Sub 全量取得_コード_フォルダ指定M()
    Dim wsOutput As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim filePath As String
    Dim lastModified As Date
    Dim fileNum As Integer
    Dim lineData As String
    Dim outputRow As Long
    Dim subroutineName As String
    Dim inSubroutine As Boolean
    Dim lineNumber As Long
    Dim overallLineNumber As Long
    Dim targetWorkbook As Workbook

    ' 特定のブックを設定（ここでは、現在アクティブなブックを対象とします）
    Set targetWorkbook = Application.ActiveWorkbook

    ' フォルダパスを指定
    folderPath = InputBox("フォルダパスを入力してください:", "フォルダ選択")
    If folderPath = "" Then Exit Sub ' キャンセルまたは未入力の場合は終了
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

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

    ' 初期化
    outputRow = 2           ' カラム名の下から出力
    overallLineNumber = 1   ' 全体の連番を初期化

    ' フォルダ直下の .bas ファイルを順番に処理
    fileName = Dir(folderPath & "*.bas")
    Do While fileName <> ""
        filePath = folderPath & fileName

        ' ファイルの最終更新日を取得
        On Error Resume Next
        lastModified = FileDateTime(filePath)
        If Err.Number <> 0 Then
            Debug.Print "ファイルが見つからない可能性があります: " & filePath
            Err.Clear
            On Error GoTo 0
            GoTo NextFile
        End If
        On Error GoTo 0

        Debug.Print "Processing file path: " & filePath ' ファイルパスをデバッグ出力

        ' ファイルを開いて1行ずつ読み込む
        fileNum = FreeFile()
        Open filePath For Input As #fileNum

        ' 初期化
        inSubroutine = False
        subroutineName = ""
        lineNumber = 1

        Do While Not EOF(fileNum)
            Line Input #fileNum, lineData

            ' サブルーチンや関数の開始を検出
            If InStr(1, lineData, "Sub ", vbTextCompare) > 0 Or InStr(1, lineData, "Function ", vbTextCompare) > 0 Then
                subroutineName = Trim(Split(lineData, " ")(1))
                inSubroutine = True
            End If

            ' サブルーチンの終了を検出
            If inSubroutine And (InStr(1, lineData, "End Sub", vbTextCompare) > 0 Or InStr(1, lineData, "End Function", vbTextCompare) > 0) Then
                inSubroutine = False
            End If

            ' 出力
            wsOutput.Cells(outputRow, "A").Value = filePath          ' フルパス
            wsOutput.Cells(outputRow, "B").Value = fileName          ' ファイル名
            wsOutput.Cells(outputRow, "C").Value = lastModified      ' 最終更新日
            wsOutput.Cells(outputRow, "D").Value = lineData          ' 1行分のコード
            wsOutput.Cells(outputRow, "E").Value = subroutineName    ' サブルーチン名
            wsOutput.Cells(outputRow, "F").Value = lineNumber        ' ファイル内の連番
            wsOutput.Cells(outputRow, "G").Value = overallLineNumber ' 全体の連番

            outputRow = outputRow + 1
            lineNumber = lineNumber + 1
            overallLineNumber = overallLineNumber + 1
        Loop

        ' ファイルを閉じる
        Close #fileNum

        Debug.Print "File content output complete for: " & filePath

NextFile:
        ' 次のファイルへ
        fileName = Dir
    Loop

    MsgBox "処理が完了しました。", vbInformation
End Sub
