' Part II: Modify the script such that it changes the word Hornets to "Bugs"

Sub ReplaceHornets()

Dim HornetCount As Integer
HornetCount = 0

'iterate over rows
For row = 1 To 6

    'for each row iterate over columns
    For col = 1 To 7

        'Condtionals
        If Cells(row, col).value = "Hornets" Then
            Cells(row, col).value = "Bugs"
        End If

    Next col
    
Next row

End Sub
