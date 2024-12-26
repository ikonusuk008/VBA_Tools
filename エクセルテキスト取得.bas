Option Explicit

Sub ExtractTextFromVariousFiles()
    Dim wsIn As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long
    Dim filePath As String
    Dim i As Long
    Dim outputRow As Long
    
    ' 入力を格納しているシートを取得
    Set wsIn = ThisWorkbook.Sheets("IN")
    
    ' 出力先シートを作成 or 既存なら再利用（初期化）
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets("ExtractedText")
    On Error GoTo 0
    
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add
        wsOut.Name = "ExtractedText"
    Else
        wsOut.Cells.Clear
    End If
    
    ' 出力シートにヘッダを設定
    outputRow = 1
    wsOut.Cells(outputRow, 1).Value = "ファイルパス"
    wsOut.Cells(outputRow, 2).Value = "区分(シート/スライド名)"
    wsOut.Cells(outputRow, 3).Value = "列/Shape名"
    wsOut.Cells(outputRow, 4).Value = "テキスト内容"
    outputRow = outputRow + 1
    
    ' 最終行を取得（A列）
    lastRow = wsIn.Cells(wsIn.Rows.Count, "A").End(xlUp).Row
    
    ' 2行目から順番に処理
    For i = 2 To lastRow
        filePath = wsIn.Cells(i, 1).Value
        
        If LCase(Right(filePath, 4)) = ".pdf" Then
            ' PDF → Acrobatなしなのでスキップ
            Debug.Print "Skipping PDF (no Acrobat library): " & filePath
            
        ElseIf LCase(Right(filePath, 4)) = ".xls" Or _
               LCase(Right(filePath, 5)) = ".xlsx" Or _
               LCase(Right(filePath, 5)) = ".xlsm" Then
            ' Excelファイル
            ExtractTextFromExcel filePath, wsOut, outputRow
            
        ElseIf LCase(Right(filePath, 5)) = ".pptx" Then
            ' PowerPointファイル
            ExtractTextFromPowerPoint filePath, wsOut, outputRow
            
        Else
            ' その他の拡張子 → スキップ
            Debug.Print "Skipping unsupported file: " & filePath
        End If
    Next i
    
    ' 見やすいように列幅調整
    wsOut.Columns("A:D").AutoFit
    
    MsgBox "抽出が完了しました。", vbInformation
End Sub

'===============================================================================
' Excelファイル(xls, xlsx, xlsm) のテキスト抽出
' 非表示シートはスキップして UsedRange をすべて走査
'===============================================================================
Private Sub ExtractTextFromExcel(ByVal filePath As String, ByVal wsOut As Worksheet, ByRef outputRow As Long)
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim c As Range
    Dim colLetter As String
    
    On Error Resume Next
    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
    On Error GoTo 0
    
    If sourceWb Is Nothing Then
        Debug.Print "Could not open Excel file: " & filePath
        Exit Sub
    End If
    
    For Each sourceWs In sourceWb.Worksheets
        ' 非表示シートはスキップ
        If sourceWs.Visible = xlSheetVisible Then
            Dim usedRg As Range
            Set usedRg = sourceWs.UsedRange
            For Each c In usedRg
                If Not IsEmpty(c.Value) Then
                    ' 列名(A,B,C...)を抜き出したいだけなので Split
                    colLetter = Split(c.Address, "$")(1)
                    
                    wsOut.Cells(outputRow, 1).Value = filePath
                    wsOut.Cells(outputRow, 2).Value = sourceWs.Name
                    wsOut.Cells(outputRow, 3).Value = colLetter
                    wsOut.Cells(outputRow, 4).Value = c.Value
                    outputRow = outputRow + 1
                End If
            Next c
        Else
            Debug.Print "Skipping hidden sheet: " & sourceWs.Name
        End If
    Next sourceWs
    
    sourceWb.Close SaveChanges:=False
End Sub

'===============================================================================
' PowerPointファイル(pptx) のテキスト抽出
' 参照設定不要(遅延バインディング)で Shapes を走査
'===============================================================================
Private Sub ExtractTextFromPowerPoint(ByVal filePath As String, ByVal wsOut As Worksheet, ByRef outputRow As Long)
    Dim pptApp As Object        ' PowerPoint.Application
    Dim pptPres As Object       ' PowerPoint.Presentation
    Dim sld As Object           ' PowerPoint.Slide
    Dim shp As Object           ' PowerPoint.Shape
    
    On Error Resume Next
    Set pptApp = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    
    If pptApp Is Nothing Then
        Debug.Print "PowerPoint not installed or could not be started."
        Exit Sub
    End If
    
    On Error Resume Next
    Set pptPres = pptApp.Presentations.Open(filePath, Untitled:=msoFalse, WithWindow:=msoFalse)
    On Error GoTo 0
    
    If pptPres Is Nothing Then
        Debug.Print "Could not open PowerPoint file: " & filePath
        pptApp.Quit
        Exit Sub
    End If
    
    ' 各スライドを巡回
    For Each sld In pptPres.Slides
        ' 各スライド上の Shape を確認
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    Dim txt As String
                    txt = shp.TextFrame.TextRange.Text
                    If Len(txt) > 0 Then
                        wsOut.Cells(outputRow, 1).Value = filePath
                        wsOut.Cells(outputRow, 2).Value = "Slide " & sld.SlideIndex
                        wsOut.Cells(outputRow, 3).Value = shp.Name
                        wsOut.Cells(outputRow, 4).Value = txt
                        outputRow = outputRow + 1
                    End If
                End If
            End If
        Next shp
    Next sld
    
    ' PowerPoint終了
    pptPres.Close
    pptApp.Quit
End Sub
