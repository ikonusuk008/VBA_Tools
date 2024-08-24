Attribute VB_Name = "特定文字色変更"
Sub 特定文字色変更M()

    Dim rng As Range
    Dim ptr As Integer

    '23: "Numeric value", "Character", "Logical value", and "Error value" are all selected
    Const dataType As Long = 23

    Dim colorChangeText As String
    colorChangeText = InputBox("色を変更するテキストを入力してください")

    Dim colorIndex As String
    colorIndex = InputBox("色番号を入力してください", "色設定", "3")

    Dim boldSetting As String
    boldSetting = InputBox("太字にしますか？ (B/bで太字)", "太字設定", " ")

    For Each rng In ActiveSheet.Cells.SpecialCells(xlCellTypeConstants, dataType)
        ptr = InStr(rng.Value, colorChangeText)

        ' Whileループでセル内の文字列をすべて見つける
        While ptr > 0

            rng.Characters(Start:=ptr, Length:=Len(colorChangeText)).Font.colorIndex = CInt(colorIndex)

            If UCase(boldSetting) = "B" Then
                rng.Characters(Start:=ptr, Length:=Len(colorChangeText)).Font.Bold = True
            End If

            ' 次の一致を検索
            ptr = InStr(ptr + Len(colorChangeText), rng.Value, colorChangeText)

        Wend
    Next rng

End Sub
