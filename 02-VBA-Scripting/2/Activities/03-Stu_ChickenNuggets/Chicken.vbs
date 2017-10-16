Sub forloop()

    Dim i As Integer

    'loop through first 10 rows
    For i = 1 To 10

        'set values in column 1 to "I will eat"
        Cells(i, 1).Value = "I will eat"

        'set values in column 2 to the sum of i + 10
        Cells(i, 2).Value = i + 10

        'set values in column 3 to "Chicken Nuggets"
        Cells(i, 3).Value = "Chicken Nuggets"

    'Call the next iteration
    Next i

End Sub
