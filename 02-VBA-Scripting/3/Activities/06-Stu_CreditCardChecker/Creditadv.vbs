Sub NextCells()

  ' Set a variable for specifying the column of interest
  Dim column, columncc, index, cc_sum As Integer
  column = 1
  columncc = 7
  index = 2
  cc_sum = 0

  ' Loop through rows in the column
  For rows = 2 To 101

    ' Searches for when the value of the next cell is different than that of the current cell
    If Cells(rows + 1, column).Value <> Cells(i, column).Value Then

      ' Message Box the value of the current cell and value of the next cell
      Cells(index, columncc).Value = Cells(rows, column).Value
      cc_sum = Cells(rows, column + 2).Value + cc_sum
      index = index + 1
      
    End If

  Next rows

End Sub