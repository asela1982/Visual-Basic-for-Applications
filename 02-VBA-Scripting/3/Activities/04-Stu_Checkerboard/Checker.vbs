Sub Checker()

Dim cellnumber As Integer
Dim rows As Integer
Dim cols As Integer

cellnumber = 1

    For rows = 1 To 8
        For cols = 1 To 8
            If (cellnumber Mod 2 = 0) Then
                Cells(rows, cols).Interior.ColorIndex = 3
            Else
                Cells(rows, cols).Interior.ColorIndex = 1
            End If
        
        cellnumber = cellnumber + 1
        Next cols

        cellnumber = cellnumber + 1
        
    Next rows

End Sub

