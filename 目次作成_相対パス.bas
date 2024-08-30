Attribute VB_Name = "目次作成_相対パス"
' ================================
' サマリ
' 作成日: 本日
' 背景: Excelシートの目次を自動生成するためのマクロ
' 目的: 全ワークシートを一覧にして目次を作成し、ハイパーリンクを設定
' 特記事項: 特定のシート名（"【" または "★" で始まる）の場合、新しい列に移動するロジックを含む
' ================================

' 目次シートを作成するサブルーチン
Sub 目次作成_相対パスM()
Attribute 目次作成_相対パスM.VB_ProcData.VB_Invoke_Func = "U\n14"

    ' ユーザーに拡大率を入力させるダイアログを表示
    Dim userMagnification As Variant
    Dim magnification As Integer ' ここで拡大率の変数を宣言

    userMagnification = InputBox("表示倍率を入力してください (例: 100)", "表示倍率設定", "90")
    
    ' 入力が数値かどうかを確認
    If IsNumeric(userMagnification) Then
        magnification = CInt(userMagnification) ' 入力された拡大率を設定
    Else
        MsgBox "有効な数値を入力してください。", vbExclamation
        Exit Sub
    End If




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
        Application.DisplayAlerts = False ' 削除の警告を表示しない
        ActiveSheet.Delete ' 既存のINDEXシートを削除
        Application.DisplayAlerts = True ' 削除の警告を元に戻す
    End If

    ' 新しいワークシートを追加し、それを"INDEX"と命名
    Worksheets.Add
    ActiveSheet.name = "INDEX"

    ' パフォーマンス向上のための設定を呼び出し
    accelerate
    
    ' 全ワークシートをループして目次を作成
    For i = 1 To Worksheets.count
        worksheetName = Worksheets(i).name
        
        ' ワークシート名が特定の文字で始まる場合、新しい列に移動
        If Mid(worksheetName, 1, 1) = "【" Or Mid(worksheetName, 1, 1) = "★" Then
            row_num = 1
            col_num = col_num + 1
        End If
        
        ' ハイパーリンクのサブアドレスを作成（相対パス）
        Dim subAddress_ As String
        subAddress_ = "'" & worksheetName & "'!A1" ' ハイパーリンクの相対パスを設定
        
        ' INDEXシートの適切なセルに移動してハイパーリンクを追加
        With Worksheets("INDEX").Cells(row_num, col_num)
            .Hyperlinks.Add Anchor:=.Cells(row_num, col_num), Address:="", SubAddress:=subAddress_, TextToDisplay:=worksheetName
            .Value = worksheetName
            .Interior.colorIndex = Worksheets(i).Tab.colorIndex
            .EntireColumn.ColumnWidth = 30
        End With
        
        row_num = row_num + 1 ' 行番号を増やす
    Next i
    
    ' パフォーマンス設定を元に戻す
    clearAccelerate
     
    ' INDEXシートの列幅を調整
    Columns("B:BB").ColumnWidth = 10
    Columns("A:A").EntireColumn.AutoFit
    Cells.EntireColumn.AutoFit
    Cells(1, 1).Select
    
    ' ウィンドウを分割して固定
    FreezePanes

    ' すべてのワークシートの表示倍率を設定
    For i = 1 To Worksheets.count
        Worksheets(i).Select
        ActiveWindow.Zoom = magnification
    Next i
    
    ' 最初のシートを選択して終了
    Worksheets(1).Select
End Sub

' ウィンドウを分割して固定するサブルーチン
Sub FreezePanes()
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub

' ワークシートが存在するか確認する関数
Function ExistsWorksheet(ByVal name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In sheets
        If ws.name = name Then
            ExistsWorksheet = True
            Exit Function
        End If
    Next
    ExistsWorksheet = False
End Function

' パフォーマンス向上のための設定を行うサブルーチン
Sub accelerate()
    With Application
        .ScreenUpdating = False ' 画面の更新を停止
        .DisplayAlerts = False ' 警告を表示しない
        .EnableEvents = False ' イベントを無効化
        .Calculation = xlCalculationManual ' 手動計算モードに設定
    End With
End Sub

' パフォーマンス設定を元に戻すサブルーチン
Sub clearAccelerate()
    With Application
        .ScreenUpdating = True ' 画面の更新を再開
        .DisplayAlerts = True ' 警告を表示
        .EnableEvents = True ' イベントを有効化
        .Calculation = xlCalculationAutomatic ' 自動計算モードに設定
    End With
End Sub


