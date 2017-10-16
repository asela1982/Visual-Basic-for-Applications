Sub fizzbuzz()

    ' loop through the value 
    For i = 2 To 100

        dim num as integer
        num = Cells(i, 1).Value

        ' check if the number is divisible by 3 and 5...
        If num Mod 3 = 0 And num Mod 5 = 0 Then
            'if so print Fizzbuzz
            Cells(i, 2).Value = "Fizzbuzz"

        ' check if the number is divisible by just 3...
        ElseIf num Mod 3 = 0 Then
            'if so print Fizz
            Cells(i, 2).Value = "Fizz"

        ' check if the number is divisible by just 5
        ElseIf num Mod 5 = 0 Then
            'if so print Buzz
            Cells(i, 2).Value = "Buzz"

        End If

    Next i

End Sub
