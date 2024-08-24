Attribute VB_Name = "フォルダ階層ファイル一覧作成"
Sub フォルダ階層ファイル一覧作成M()

    Dim mainFolder As String
    Dim subFolder1 As Object
    Dim subFolder2 As Object
    Dim subFolder3 As Object
    Dim file As Object
    Dim fileSystem As Object
    Dim currentRow As Long
    Dim maxExcelCount As Long, maxTextCount As Long, maxCsvCount As Long, maxGrandchildCount As Long
    Dim excelFileCount As Long, textFileCount As Long, csvFileCount As Long, grandchildFileCount As Long
    Dim excelFiles() As String, textFiles() As String, csvFiles() As String, grandchildFiles() As String
    
    ' メインフォルダのパスを指定
    mainFolder = "C:\path\to\your\folder" ' フォルダのパスを指定
    
    ' ファイルシステムオブジェクトを作成
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' 初期設定
    currentRow = 2 ' 表の書き込み開始行（1行目にヘッダー）
    
    ' サブフォルダ2およびその配下のフォルダ（サブフォルダ3）内の最大ファイル数を取得
    maxExcelCount = 0
 