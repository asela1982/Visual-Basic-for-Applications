' Easy
' Create a script that will loop through each year of stock data and grab the total amount of 
' volume each stock had over the year.
' You will also need to display the ticker symbol to coincide with the total volume.

Sub easy()

    'loop through all the worksheets
    for each ws in worksheets

        ' Set an initial variable for holding the ticker name
        Dim ticker_name As String

        ' Set an initial variable for holding the total volume per ticker name
        Dim ticker_volume As Double
        ticker_volume = 0

        ' Keep track of the location for each ticker name in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

            ' derive the row index of the last row of the worksheet
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            ' print the headers to the summary table
            ws.Range("i1") = "Ticker Name"
            ws.Range("j1") = "Ticker Volume"


            ' Loop through all the days
            For i = 2 To lastrow

                    ' check if we are still within the same ticker, if it is not...
                    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                        
                        ' set the ticker name
                        ticker_name = ws.Cells(i, 1).Value
                        
                        ' add to the ticker volume
                        ticker_volume = ticker_volume + ws.Cells(i, 7).Value

                        ' print the ticker name to the summary table
                        ws.Range("i" & Summary_Table_Row) = ticker_name

                        ' print the ticker volume to the summary table
                        ws.Range("j" & Summary_Table_Row) = ticker_volume

                        'add one to the summary table
                        Summary_Table_Row = Summary_Table_Row + 1

                        'reset the ticket volume
                        ticker_volume = 0


                    Else
                        ' add to the ticket volume
                        ticker_volume = ticker_volume + ws.Cells(i, 7).Value

                    End If

            Next i

        next

End Sub
