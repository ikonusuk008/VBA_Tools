Attribute VB_Name = "ファイル最終更新日取得"
Sub ファイル最終更新日取得M()
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String

    ' A列の最終行を取得
    lastRow = Cells(Rows.count, 1).End(xlUp).row

    ' 1行目から最終行までループ
    For i = 1 To lastRow
        filePath = Cells(i, 1).Value ' A列のファイルパスを取得

        ' ファイルが存在する場合のみ最終更新日をB列に書き込み
        If Dir(filePath) <> "" Then
            Cells(i, 2).Value = FileDateTime(filePath) ' B列に最終更新日を書き込み
        Else
            Cells(i, 2).Value = "ファイルが存在しません" ' ファイルが存在しない場合のメッセージ
        End If
    Next i
End Sub
