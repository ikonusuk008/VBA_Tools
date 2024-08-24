Attribute VB_Name = "目次作成_相対パス"
Const magnification = 90 ' 表示倍率を90%に設定

' ================================
' サマリ
' 作成日: 本日
' 背景: Excelシートの目次を自動生成するためのマクロ
' 目的: 全ワークシートを一覧にして目次を作成し、ハイパーリンクを設定
' 特記事項: 特定のシート名（"【" または "★" で始まる）の場合、新しい列に移動するロジックを含む
' ================================

' 目次シートを作成するサブルーチン
Sub 目次作成_相対パスM()
    Dim sheetCount As Integer ' ワークシートの数
    Dim worksheetName As String ' ワークシートの名前
    Dim Worksheet As Worksheet ' ワークシートオブジェクト
    Dim col_num As Integer ' 列番号
    col_num = 1 ' 初期列番号を1に設定
    Dim j As Integer ' ループカウンタ
    Dim row_num As Integer ' 行番号
    row_num = 1 ' 初期行番号を1に設定
    
    ' 既に"INDEX"という名前のワークシートが存在する場合、それを削除
    If ExistsWorksheet("INDEX") Then
        Worksheets("INDEX").Select
        Application.DisplayAlerts = False ' 削除