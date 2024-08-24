Attribute VB_Name = "全量取得_コード_フォルダ指定"
Sub 全量取得_コード_フォルダ指定M()
    Dim wsOutput As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim filePath As String
    Dim lastModified As Date
    Dim fileContent As String
    Dim fileNum As Integer
    Dim lineData As String
    Dim outputRow As Long
    Dim subroutineName As String
    Dim inSubroutine As Boolean
    Dim lineNumber As Long
    Dim overallLineNumber As Long
    Dim targetWorkbook As Workbook

    ' 特定のブックを設定（ここでは、現在アクティブなブックを対象とします）
    Set targetWorkbook = Application.ActiveWorkbook

    ' フォルダパスを指定
    folderPath = InputBox("フォルダパスを入力してください:", "フォルダ選択")
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' 出力シートを作成または取得
    On Error Resume Next
    Set wsOutput = targetWorkbook.sheets("VBA全量")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Se