' StarCounter
' 1. Create a nested for loop that iterates through each student.
' 2. For each loop count the number of instances of the word "Full-Star" using a counter
' 3. Save the counter value to the total cell
' 4. BONUS: Instead of hard-coding the last number of the loop, use VBA to determine the last row.
' 5. BONUS: Create two charts:
     ' One to see if there is a relationship between Program type and Rating
     ' One to see if there is a relationship between Date and Rating

Sub fullstar()

Dim counter As Integer
Dim lastrow As Integer
Dim lastcol As Integer

'get the last row and last col
lastrow = Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row
lastcol = Sheets("Sheet1").Cells(1, Columns.Count).End(xlToLeft).Column

    'loop through each student
    For i = 2 To lastrow
        
        'initialize a counter object to count the number of instances of the word "Full-Star"
        counter = 0

        'for each student through the star columns
        For j = 1 To 5

            'execute a if conditional statment
            If Cells(i, j + 3).Value = "Full-Star" Then
                counter = counter + 1
            End If

        Next j
        'assign the counter value to the total column
        Cells(i, lastcol).Value = counter
    Next i

End Sub