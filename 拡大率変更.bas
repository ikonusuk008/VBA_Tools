Attribute VB_Name = "拡大率変更"
Sub 拡大率変更M()

    Dim buf As String
    Dim magnification As Long
    magnification = 80 ' デフォルトのズーム倍率
    buf = InputBox("input size", "Sheet size setting", magnification)

    For i = 1 To Worksheets.count
        Worksheets(i).Select
        ActiveWindow.Zoom = buf
    Next i
    
    Worksheets(1).Select
    
End Sub
