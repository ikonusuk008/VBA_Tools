Attribute VB_Name = "gċĤÏX"
Sub gċĤÏXM()

    Dim buf As String
    Dim magnification As Long
    magnification = 80 ' ftHgÌY[{Ĥ
    buf = InputBox("input size", "Sheet size setting", magnification)

    For i = 1 To Worksheets.count
        Worksheets(i).Select
        ActiveWindow.Zoom = buf
    Next i
    
    Worksheets(1).Select
    
End Sub
