Attribute VB_Name = "色_文字列"
' # Excelマクロで特定の文字列の色と太字を変更する方法
' ## はじめに
' このブログ記事では、Excelシート内の特定の文字列の色を変更し、必要に応じて太字にするマクロの作成方法を解説します。<br>ユーザーは、変更したいテキスト、色、および太字設定を入力ボックスを通じて指定します。<br>
' ## 色と太字を変更するマクロ
' このマクロは、ユーザーからの入力を受け取り、シート内の特定の文字列の色と太字設定を変更します。<br>
' <pre><code>
Sub 色と太字を変更する()
Attribute 色と太字を変更する.VB_ProcData.VB_Invoke_Func = "r\n14"
    Const データ型 As Long = 23
    Dim 範囲 As Range
    Dim 色変更テキスト As String
    Dim 色インデックス As String
    Dim 太字設定 As String
    Dim rng As Range
    
    ' ユーザーに変更したいテキストを入力させる
    色変更テキスト = InputBox("テキストを入力してください")
    
    ' ユーザーに色インデックスを入力させる（デフォルトは3）
    色インデックス = InputBox("色を入力してください", "色設定", "3")
    
    ' ユーザーに太字設定を入力させる（B/bを入力すると太字になる）
    太字設定 = InputBox("テキストを太字にしますか？ (B/bを入力してください)", "太字設定", " ")
    
    ' シート内の定数セルを範囲として設定
    Set 範囲 = ActiveSheet.Cells.SpecialCells(xlCellTypeConstants, データ型)
    
    ' 範囲内の各セルに対して色変更を実行
    For Each rng In 範囲
        Call 色変更(rng, 色変更テキスト, 色インデックス, 太字設定)
    Next rng
End Sub
' </code></pre>
' ## 色変更サブルーチン
' このサブルーチンは、指定された範囲内の特定の文字列の色と太字設定を変更します。<br>
' <pre><code>
Sub 色変更(rng As Range, 色変更テキスト As String, 色インデックス As String, 太字設定 As String)
    Dim ポインタ As Integer
    
    ' 指定されたテキストの位置を検索
    ポインタ = InStr(rng.Value, 色変更テキスト)
    
    ' テキストが見つかる限りループ
    While ポインタ > 0
        ' 指定されたテキストの色を変更
        rng.Characters(Start:=ポインタ, Length:=Len(色変更テキスト)).Font.colorIndex = CInt(色インデックス)
        
        ' 太字設定が"B"の場合、太字にする
        If UCase(太字設定) = "B" Then
            rng.Characters(Start:=ポインタ, Length:=Len(色変更テキスト)).Font.Bold = True
        End If
        
        ' 次の位置を検索
        ポインタ = InStr(ポインタ + Len(色変更テキスト), rng.Value, 色変更テキスト)
    Wend
End Sub
' </code></pre>

