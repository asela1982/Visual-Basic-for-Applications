Sub creditcard_basic()

'derive the row index of the last row in the worksheet
lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

    ' loop through each of the credit cards
    For i = 2 To lastrow

        ' initate an if conditional to check if the cc type of the current cell differs to that of
        ' the following cell
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value And Cells(i + 1, 1).Value <> "" Then
            'if differs, pop-up a message box with the names of the different credit card types
            MsgBox (Cells(i, 1).Value + " then " + Cells(i + 1, 1).Value)

        End If

    Next i

End Sub
