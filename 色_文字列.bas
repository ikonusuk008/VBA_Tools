Attribute VB_Name = "色_文字列"
' シート内の特定の文字列の色と太字を変更するマクロ
' （特定文字色変更.bas と同機能を、メイン処理とサブルーチンの2つに分けた構成）
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
    If 色変更テキスト = "" Then Exit Sub ' キャンセルまたは未入力の場合は終了

    ' ユーザーに色インデックスを入力させる（デフォルトは3）
    色インデックス = InputBox("色を入力してください", "色設定", "3")
    If 色インデックス = "" Then Exit Sub ' キャンセルまたは未入力の場合は終了

    ' ユーザーに太字設定を入力させる（B/bを入力すると太字になる）
    太字設定 = InputBox("テキストを太字にしますか？ (B/bを入力してください)", "太字設定", " ")

    ' シート内の定数セルを範囲として設定（定数セルが1つも無い場合は実行時エラーになるため事前チェック）
    On Error Resume Next
    Set 範囲 = ActiveSheet.Cells.SpecialCells(xlCellTypeConstants, データ型)
    On Error GoTo 0
    If 範囲 Is Nothing Then
        MsgBox "対象となるセル（定数セル）が見つかりませんでした。", vbInformation
        Exit Sub
    End If

    ' 範囲内の各セルに対して色変更を実行
    For Each rng In 範囲
        Call 色変更(rng, 色変更テキスト, 色インデックス, 太字設定)
    Next rng
End Sub

' 指定された範囲内の特定の文字列の色と太字設定を変更するサブルーチン
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
