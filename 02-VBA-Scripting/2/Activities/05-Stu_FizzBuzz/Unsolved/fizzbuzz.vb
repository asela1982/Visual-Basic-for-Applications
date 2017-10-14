Sub fizzbuzz()

'declare variables
'loop through column1 and grab the numbers
'conditionals
'set output

    Dim i As Integer

    For i = 2 To 100

        If (Cells(i, 1).Value Mod 3 = 0 And Cells(i, 1).Value Mod 5 = 0) Then
        Cells(i, 2).Value = "Fizzbuzz"

        ElseIf (Cells(i, 1).Value Mod 3 = 0) Then
        Cells(i, 2).Value = "Fizz"

        ElseIf (Cells(i, 1).Value Mod 5 = 0) Then
        Cells(i, 2).Value = "Buzz"

        End If

    Next i

End Sub
