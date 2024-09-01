Attribute VB_Name = "ファイルパス分解処理"
Sub ファイルパス分解処理M()
    Dim 最終行 As Long
    Dim i As Long
    Dim ファイルパス As String
    Dim パス部分() As String
    Dim ファイル名 As String
    Dim 拡張子 As String
    Dim j As Long
    Dim データ As Variant
    Dim 結果 As Variant

    ' 項目名を設定
    Cells(1, 1).Value = "ファイルパス"
    Cells(1, 2).Value = "ファイル名"
    Cells(1, 3).Value = "拡張子"

    ' A列の最終行を取得
    最終行 = Cells(Rows.count, 1).End(xlUp).row

    ' 2行目から最終行までループ
    For i = 2 To 最終行
        ' A列のファイルパスを取得
        ファイルパス = Cells(i, 1).Value
        
        ' ファイルパスを分解
        パス部分 = Split(ファイルパス, "\")
        
        ' ファイル名と拡張子を取得
        ファイル名 = パス部分(UBound(パス部分))
        拡張子 = Split(ファイル名, ".")(1)
        
        ' 結果をシートに書き込む
        Cells(i, 2).Value = ファイル名
        Cells(i, 3).Value = 拡張子
        
        ' フォルダ階層を列に分割して設定
        For j = LBound(パス部分) To UBound(パス部分) - 1
            Cells(1, 4 + j).Value = "F" & (j + 1)
            Cells(i, 4 + j).Value = パス部分(j)
        Next j
    Next i

    ' A列をリンクに設定
    For i = 2 To 最終行
        Cells(i, 1).Hyperlinks.Add Anchor:=Cells(i, 1), Address:=Cells(i, 1).Value, TextToDisplay:=Cells(i, 1).Value
    Next i

    ' 項目行を固定
    If ActiveWindow.FreezePanes = False Then
        Rows("2:2").Select
        ActiveWindow.FreezePanes = True
    End If

    ' 全体をフィルタ
    If ActiveSheet.AutoFilterMode = False Then
        Cells(1, 1).CurrentRegion.AutoFilter
    End If
End Sub

