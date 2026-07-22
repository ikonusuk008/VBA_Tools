Attribute VB_Name = "特定文字色変更"
Sub 特定文字色変更M()

    Dim rng As Range
    Dim targetRange As Range
    Dim ptr As Integer

    '23: "Numeric value", "Character", "Logical value", and "Error value" are all selected
    Const dataType As Long = 23

    Dim colorChangeText As String
    colorChangeText = InputBox("色を変更するテキストを入力してください")
    If colorChangeText = "" Then Exit Sub ' キャンセルまたは未入力の場合は終了

    Dim colorIndex As String
    colorIndex = InputBox("色番号を入力してください", "色設定", "3")
    If colorIndex = "" Then Exit Sub ' キャンセルまたは未入力の場合は終了

    Dim boldSetting As String
    boldSetting = InputBox("太字にしますか？ (B/bで太字)", "太字設定", " ")

    ' 定数セルを取得（定数セルが1つも無い場合は実行時エラーになるため事前チェック）
    On Error Resume Next
    Set targetRange = ActiveSheet.Cells.SpecialCells(xlCellTypeConstants, dataType)
    On Error GoTo 0
    If targetRange Is Nothing Then
        MsgBox "対象となるセル（定数セル）が見つかりませんでした。", vbInformation
        Exit Sub
    End If

    For Each rng In targetRange
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
