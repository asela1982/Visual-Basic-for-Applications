' StarCounter
' 1. Create a nested for loop that iterates through each student.
' 2. For each loop count the number of instances of the word "Full-Star" using a counter
' 3. Save the counter value to the total cell
' 4. BONUS: Instead of hard-coding the last number of the loop, use VBA to determine the last row.
' 5. BONUS: Create two charts:
     ' One to see if there is a relationship between Program type and Rating
     ' One to see if there is a relationship between Date and Rating

Sub StarCounter()

  ' Create a variable to hold the StarCounter. We will repeatedly use this.
  Dim StarCount As Integer


  ' Loop through each row
  For i = 2 To 51

    ' Initially set the StarCounter to be 0 for each row
      StarCount = 0

    ' While in each row, loop through each star column
        For j = 1 To 5

      ' If a column contains the word "Full-Star"...
          If Cells(i, j + 3).Value = "Full-Star" Then
        ' Add 1 to the StarCounter
            StarCount = StarCount + 1
          End If
        Next j
    ' Once we've completed all rows, print the value in the total column
        Cells(i, 9).Value = StarCount
  Next i

End Sub
