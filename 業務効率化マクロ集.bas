Attribute VB_Name = "業務効率化マクロ集"
' 現在のセルの行を削除するマクロ
Sub One_row_delete()
    ' 現在のセルの行を選択
    ActiveCell.EntireRow.Select
    ' 選択した行を削除
    Rows(ActiveCell.row).Delete
End Sub

' 現在のセルの行に新しい行を挿入するマクロ
Sub One_row_insert()
    ' 現在のセルの行を選択
    ActiveCell.EntireRow.Select
    ' 選択した行に新しい行を挿入
    Rows(ActiveCell.row).Insert
End Sub

' 現在のセルの列を削除するマクロ
Sub One_Column_delete()
    ' 現在のセルの列を選択
    ActiveCell.EntireColumn.Select
    ' 選択した列を削除
    Columns(ActiveCell.Column).Delete
End Sub

' 現在のセルの列に新しい列を挿入するマクロ
Sub One_Column_insert()
    ' 現在のセルの列を選択
    ActiveCell.EntireColumn.Select
    ' 選択した列に新しい列を挿入
    Columns(ActiveCell.Column).Insert
End Sub

' すべての行の高さを自動調整するマクロ
Sub Automatic_row_height_adjustment()
    ' すべてのセルを選択
    Cells.Select
    ' すべての行の高さを自動調整
    Cells.EntireRow.AutoFit
    ' セルA1を選択（選択状態を解除するため）
    Range("A1").Select
End Sub

' フォントと日付列幅の調整を行うマクロ
Sub カラムの縮尺調整()
    ' すべてのセルを選択
    Cells.Select
    ' 列幅を設定
    Selection.ColumnWidth = 8.29
    ' すべての列の幅を自動調整
    Cells.EntireColumn.AutoFit
    ' セルA1を選択（選択状態を解除するため）
    Range("A1").Select
End Sub

' 現在のセルの改行と空白を削除するマクロ
Sub Eliminate_line_breaks_and_blanks_in_the_active_Cell()
    ' 現在のセルの値を取得
    a = ActiveCell.Value
    ' 改行を削除
    b = Replace(a, vbLf, "")
    ' 空白を削除
    b = Replace(b, " ", "")
    ' 修正した値を現在のセルに設定
    ActiveCell.Value = b
End Sub

' 現在のセルの下のセルの値をコピーするマクロ
Sub Copy_cells_below_active_cell()
    ' 現在のセルに下のセルの値を設定
    ActiveCell.Cells(1, 1).Value = ActiveCell.Cells(2, 1).Value
End Sub

' すべてのシート名をセルに表示するマクロ
Sub シート名取得エクセル関数()
    For i1 = 1 To Worksheets.count - 1
        ' 各シートを選択
        Worksheets.Select (i1)
        ' B2セルを選択
        Range("B2").Select
        ' B2セルをクリア
        ActiveCell.FormulaR1C1 = ""
        ' A1セルにシート名を設定
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "=RIGHT(CELL(""filename"",RC),LEN(CELL(""filename"",RC))-FIND(""]"",CELL(""filename"",RC)))"
        ' A2セルを選択して1行目の高さを設定
        Range("A2").Select
        Rows("1:1").RowHeight = 23.25
    Next i1
End Sub

' すべてのシート名をデバッグ出力するマクロ
Sub シート名の確認()
    For i1 = 1 To Worksheets.count - 1
        ' 各シートの名前をデバッグ出力
        Debug.Print Worksheets(i1).name
    Next i1
End Sub
