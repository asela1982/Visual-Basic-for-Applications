Sub lotto_winner()

    ' create variables to hold the winning numbers 
    Dim first, second, third As Long

    ' assign the winning numbers to the variables created above
    first = 3957481
    second = 5865187
    third = 2817729

   ' create variables to hold the winning numbers 
    Dim run1, run2, run3 As Long
    
   ' assign the winning numbers to the variables created above
    run1 = 2275339
    run2 = 5868182
    run3 = 1841402

    ' loop through each of the lotto tickets
    For i = 2 To 1001

        ' check if the lotto numbers match the first place winner ...
        If Cells(i, 3).Value = first Then

            'retrive the values associated with the winner enter them into the winner's box
            Cells(2, 6).Value = Cells(i, 1).Value
            Cells(2, 7).Value = Cells(i, 2).Value
            Cells(2, 8).Value = Cells(i, 3).Value

            ' create a message specifiying the first place winner
            MsgBox "Congrats! " + Cells(2, 6).Value + " " + Cells(2, 7).Value

         ' check if the lotto numbers match the second place winner ...
        ElseIf Cells(i, 3).Value = second Then

            'retrive the values associated with the winner enter them into the winner's box
            Cells(3, 6).Value = Cells(i, 1).Value
            Cells(3, 7).Value = Cells(i, 2).Value
            Cells(3, 8).Value = Cells(i, 3).Value

         ' check if the lotto numbers match the third place winner ...
        ElseIf Cells(i, 3).Value = third Then
            'retrive the values associated with the winner enter them into the winner's box
            Cells(4, 6).Value = Cells(i, 1).Value
            Cells(4, 7).Value = Cells(i, 2).Value
            Cells(4, 8).Value = Cells(i, 3).Value

        ' end this series of IF/ELSE conditionals
        End If
    
    Next i

    ' next loop through the lotto tickets second time to find the first instance of a "runner-up" winner
    For i = 2 To 1001

        ' check for runner ups with an OR operator
        If Cells(i, 3).Value = run1  or Cells(i, 3).Value = run2 or Cells(i, 3).Value = run3 Then

            ' retrive the values associated with the winner and enter the values into the winner's box
            runner_up = Cells(i,3).value
            Cells(5, 6).Value = Cells(i, 1).Value
            Cells(5, 7).Value = Cells(i, 2).Value
            Cells(5, 8).Value = runner_up

            ' if a match is found, exit the for loop
            exit for

        End if

    Next i 

End Sub
