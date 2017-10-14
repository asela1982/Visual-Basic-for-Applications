' Part I: Count the number of Hornets found and display the number to your user in the form of a message box.

Sub HorentCount()

Dim HornetCount As Integer
HornetCount = 0

'iterate over rows
For row = 1 To 6

    'for each row iterate over columns
    For col = 1 To 7

        'Condtionals
        If Cells(row, col).value = "Hornets" Then
            HornetCount = HornetCount + 1
        End If

    Next col
    
Next row

MsgBox (HornetCount & " Hornets Found")

End Sub
