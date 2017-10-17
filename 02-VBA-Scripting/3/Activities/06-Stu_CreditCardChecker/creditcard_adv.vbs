Sub creditcard_basic()

'derive the row index of the last row in the worksheet
lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
j = 2
    ' loop through each of the credit cards
    For i = 2 To lastrow
    
        ' initate an if conditional to check if the cc type of the current cell differs to that of
        ' the following cell
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Cells(j, 7).Value = Cells(i, 1).Value
            j = j + 1

        End If
    
    Next i

' derive the row index of the last in the summary table
lastrow_table = ActiveSheet.Cells(Rows.Count, 7).End(xlUp).Row

    ' loop through each of the distinct credit card types in the summary table
    for x = 2 to lastrow_table
    ' set an initial value for holding the total per credit card brand
    sum = 0

        ' for each of the distinct credit cary brands, loop through the credit card purchases in the worksheet
        ' to find matching values
        For y = 2 To lastrow
            if cells(x, 7).value = cells(y, 1).value Then
                sum = sum + cells(y, 3).value
            end if
        next y
        cells(x , 8).value = sum

    next x
End Sub

