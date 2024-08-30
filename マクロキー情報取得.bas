Attribute VB_Name = "マクロキー情報取得"
Sub マクロキー情報取得M()
    ' シート「キー」を選択
    On Error Resume Next
    sheets("キー一覧").Select
    If Err.Number <> 0 Then
        MsgBox "シート「キー」が見つかりません。シート名を確認してください。", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' 変数の初期化
    Dim パス As String
    Dim ファイル番号 As Integer
    Dim 行データ As String
    Dim 行番号 As Integer
    行番号 = 1
    Dim 一時配列() As String
    Dim モジュール名 As String
    Dim キー名 As String
    Dim コンポーネント As Object
    
    ' 定数の定義
    Const 属性1 As String = "Attribute "
    Const 属性2 As String = "VB_Invoke_Func ="
    Const 一時ファイル As String = "Temp1.bas"
    
    ' パスの取得
    パス = ThisWorkbook.path & "\"
    
    ' VBProjectから各コンポーネントを取得し、処理
    With ThisWorkbook.VBProject
        For Each コンポーネント In .VBComponents
            ' コンポーネントを一時ファイルとしてエクスポート
            .VBComponents(コンポーネント.name).Export fileName:=パス & 一時ファイル
            ファイル番号 = FreeFile()
            Open パス & 一時ファイル For Input As #ファイル番号
            
            ' ファイル内容を読み取り
            While Not EOF(ファイル番号)
                Line Input #ファイル番号, 行データ
                
                ' Sub名を取得
                If InStr(1, 行データ, "Sub", vbTextCompare) = 1 Then
                    モジュール名 = Mid$(行データ, InStr(行データ, "Sub") + 4)
                End If
                
                ' Attribute情報を取得
                If InStr(行データ, 属性1) = 1 And InStr(行データ, 属性2) > 0 Then
                    ReDim Preserve 一時配列(行番号)
                    キー名 = ":" & Mid$(行データ, InStrRev(行データ, "=") + 3, 1)
                    一時配列(行番号) = モジュール名 & キー名
                    
                    ' 結果をシートに出力
                    Cells(行番号 + 1, 1) = モジュール名
                    Cells(行番号 + 1, 2) = Replace(キー名, ":", "")
                    行番号 = 行番号 + 1
                    モジュール名 = ""
                End If
                行データ = ""
            Wend
            
            ' ファイルを閉じて削除
            Close #ファイル番号
            Kill パス & 一時ファイル
        Next
    End With
    
    ' 書式設定
    キーリスト書式設定
End Sub

Sub キーリスト書式設定()
    ' ヘッダ部分の書式設定
    With Range("A1:C1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    ' フリーズペインの設定
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    ' オートフィルタの設定
    Cells.Select
    Selection.AutoFilter
    
    ' 表示倍率の設定
    ActiveWindow.Zoom = 75
    
    ' 列幅の設定
    Columns("A:A").ColumnWidth = 60
    Columns("B:B").ColumnWidth = 10
    
    ' ヘッダの設定
    Range("A1").Value = "モジュール名"
    Range("B1").Value = "キー名"
    Range("C1").Value = "キー候補"
    
    ' 最初のセルを選択
    Range("A1").Select
End Sub

