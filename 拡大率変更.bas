Attribute VB_Name = "�g�嗦�ύX"
Sub �g�嗦�ύXM()

    Dim buf As String
    Dim magnification As Long
    magnification = 80 ' �f�t�H���g�̃Y�[���{��
    buf = InputBox("input size", "Sheet size setting", magnification)

    For i = 1 To Worksheets.count
        Worksheets(i).Select
        ActiveWindow.Zoom = buf
    Next i
    
    Worksheets(1).Select
    
End Sub
