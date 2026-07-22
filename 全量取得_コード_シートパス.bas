Sub 全量取得_コード_シートパスM()
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim filePath As String
    Dim fileName As String
    Dim lastModified As Date
    Dim i As Long
    Dim testFlag As Boolean
    Dim lineData As String
    Dim outputRow As Long
    Dim targetWorkbook As Workbook
    Dim subroutineName As String
    Dim inSubroutine As Boolean
    Dim lineNumber As Long
    Dim overallLineNumber As Long
    
    ' ADODB.Stream 用（遅延バインディング）
    Dim stm As Object

    ' 特定のブックを設定（ここでは、現在アクティブなブックを対象とします）
    Set targetWorkbook = Application.ActiveWorkbook

    ' テストフラグを設定
    testFlag = False ' True の場合、2行だけ処理

    ' 出力シートを作成または取得
    On Error Resume Next
    Set wsOutput = targetWorkbook.Sheets("VBA全量")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Set wsOutput = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        wsOutput.Name = "VBA全量"
    End If
    
    ' 2行目以降をクリア（フィルタを維持）
    wsOutput.Rows("2:" & wsOutput.Rows.Count).ClearContents
    
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
    Set wsSource = targetWorkbook.Sheets("VBA_path")
    
    ' 最終行を取得
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    
    ' テストフラグがTrueなら、処理行数を制限
    If testFlag Then
        ' ヘッダ行が1行あるため、2行のみ処理するなら行数= 1(ヘッダ) + 2(実際に読む行)
        lastRow = WorksheetFunction.Min(lastRow, 3)
    End If
    
    ' 初期化
    outputRow = 2           ' カラム名の下から出力
    overallLineNumber = 1   ' 全体の連番を初期化

    ' 2行目からループ
    For i = 2 To lastRow
        filePath = wsSource.Cells(i, "B").Value ' B列からファイルパスを取得
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1) ' ファイル名を取得
        If filePath <> "" And Right(filePath, 4) = ".bas" Then
            ' ファイルの最終更新日を取得
            On Error Resume Next
            lastModified = FileDateTime(filePath)
            If Err.Number <> 0 Then
                Debug.Print "ファイルが見つからない可能性があります: " & filePath
                Err.Clear
                GoTo NextFile
            End If
            On Error GoTo 0
            
            Debug.Print "Processing file path: " & filePath ' ファイルパスをデバッグ出力

            '===== ここから ADODB.Stream で UTF-8 読み込み =====
            Set stm = CreateObject("ADODB.Stream")
            With stm
                .Type = 2         ' adTypeText
                .Charset = "UTF-8"
                .Open
                .LoadFromFile filePath
            End With

            ' 初期化
            inSubroutine = False
            subroutineName = ""
            lineNumber = 1
            
            ' ファイルを最後まで 1行ずつ読み込む
            Do Until stm.EOS
                lineData = stm.ReadText(-2)  ' -2: 一行ずつ読み込み

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

            ' ストリームをクローズ
            stm.Close
            Set stm = Nothing
            
            Debug.Print "File content output complete for: " & filePath
        Else
            Debug.Print "Skipping file: " & filePath ' 条件を満たさない場合はスキップ
        End If
NextFile:
    Next i
    
    MsgBox "処理が完了しました。", vbInformation
End Sub
