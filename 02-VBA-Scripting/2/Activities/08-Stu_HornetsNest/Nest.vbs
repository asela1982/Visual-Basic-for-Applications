' Part I: Count the number of Hornets found and display the number to your user in the form of a message box.

Sub HorentCounter()

    ' activate Sheet1
    Worksheets("Sheet2").Activate

    Dim Hornets As Integer
    counter = 0

    'iterate over rows
    For Row = 1 To 6

        'for each row iterate over columns
        For col = 1 To 7

            'if cell value matches a horent, increment the counter by 1.
            If Cells(Row, col).Value = "Hornets" Then
                counter = counter + 1
            End If

        Next col
        
    Next Row

    MsgBox (counter & " Hornets Found")

End Sub



' Part II: Modify the script such that it changes the word Hornets to "Bugs".

Sub HorentsToBugs()

    ' activate Sheet2
    Worksheets("Sheet2").Activate

    'iterate over the rows
    For Row = 1 To 6

        'for each row, iterate over the columns
        For col = 1 To 7

            'if cell value matches a horent, change the cell value to "bugs"
            If Cells(Row, col).Value = "Hornets" Then
                Cells(Row, col).Value = "Bugs"
            End If

        Next col
        
    Next Row

End Sub



  ' Part III: Modify the script a third time, this time keeping in mind that you
  ' have a limited number of Bugs and Bees. Use the full set of Bugs and Bees you
  ' have available to replace the Hornets. If you run out of Bugs or Bees provide
  ' the user with the message: "Oh no! We still have hornets..."

Sub ReplaceHornets()

    ' activate sheet3
    Worksheets("Sheet3").Activate

    'create two variables to hold the values of bugs and bees
    Dim bugs, bees As Integer

    'initialize the the above variables by the available bugs and bees
    bugs = Range("l2").Value
    bees = Range("r2").Value

    'iterate over the rows
    For Row = 1 To 6

        ' for each row, iterate over the columns
        For col = 1 To 7

            If Cells(Row, col).Value = "Hornets" And bugs > 0 Then
                'replace the hornets with bugs
                Cells(Row, col) = "Bugs"
                'decrement available bugs by 1
                bugs = bugs - 1

            ElseIf Cells(Row, col).Value = "Hornets" And bees > 0 Then
                'replace the hornets with bees
                Cells(Row, col) = "Bees"
                'decrement available bees by 1
                bees = bees - 1

            ElseIf Cells(Row, col).Value = "Hornets" And bugs = 0 And bees = 0 Then
                MsgBox "Oh no! We still have hornets..."

            End If

        Next col
    
    Next Row

MsgBox (bugs & "Number of Bugs left")
MsgBox (bees & "Number of Bees left")

End Sub


