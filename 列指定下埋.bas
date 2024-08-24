Attribute VB_Name = "列指定下埋"
Sub 列指定下埋M()
    Const 開始行 As Long = 2

    Dim 対象列 As Long
    Dim 最終行 As Long
    Dim 現在の値 As String
    Dim i As Long

    ' 列番号をユーザーに入力させる
    対象列 = InputBox("列番号を入力してください（例：A列なら1、B列なら2）")

    ' 無効な列番号が入力された場合、処理を終了
    If 対象列 < 1 Then Exit Sub

    ' 最終行を取得
    最終行 = Cells(Rows.count, 対象列).End(xlUp).row

    ' 初期の現在の値を設定
    現在の値 = Cells(開始行, 対象列).Value

    ' 下埋め処理を実行
    For i = 開始行 To 最終行
        If Cells(i, 対象列).Value <> "" Then
            現在の値 = Cells(i, 対象列).Value
        End If
        Cells(i, 対象列).Value = 現在の値
    Next i
End Sub

