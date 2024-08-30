Attribute VB_Name = "目次作成_相対パス"
Const magnification = 90

Sub 目次作成M()
    Dim sheetCount As Integer
    Dim worksheetName As String
    Dim Worksheet As Worksheet
    Dim col_num As Integer
    col_num = 1
    Dim j As Integer
    row_num = 1
    
    If ExistsWorksheet("INDEX") Then
        Worksheets("INDEX").Select
        Application.DisplayAlerts = False
        ActiveSheet.Delete
        Application.DisplayAlerts = True
    End If

    Worksheets.Add
    ActiveSheet.name = "INDEX"

    accelerate
    
    For i = 1 To Worksheets.count
        worksheetName = Worksheets(i).name
        
        If Mid(worksheetName, 1, 1) = "【" Or Mid(worksheetName, 1, 1) = "★" Then
            row_num = 1
            col_num = col_num + 1
        End If
       
        Dim subAddress_ As String
        subAddress_ = "'" & worksheetName & "'" & "!" & Worksheets(i).Cells(1, 1).Address
        
        Worksheets("INDEX").Cells(row_num, col_num).Activate
        
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=subAddress_
        ActiveCell.Value = worksheetName
        Cells(row_num, col_num) = Worksheets(i).name
        Cells(row_num, col_num).Interior.colorIndex = Worksheets(i).Tab.colorIndex
        Cells(row_num, col_num).EntireColumn.ColumnWidth = 30
        
        row_num = row_num + 1
    Next i
    
    clearAccelerate
     
    Columns("B:BB").ColumnWidth = 10
    Columns("A:A").EntireColumn.AutoFit
    Cells.EntireColumn.AutoFit
    Cells(1, 1).Select
    
    FreezePanes

    ' ズーム倍率をユーザーに入力してもらう
    Dim magnification As Integer
    magnification = InputBox("ズーム倍率を入力してください（例: 90）", "ズーム設定", 90)
    
    For i = 1 To Worksheets.count
        Worksheets(i).Select
        ActiveWindow.Zoom = magnification
    Next i

    
    Worksheets(1).Select
End Sub

Sub FreezePanes()
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub

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

Sub accelerate()
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
End Sub

Sub clearAccelerate()
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

