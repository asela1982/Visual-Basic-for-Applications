
' Hard
' Your solution will include everything from the moderate challenge.
' Your solution will also be able to locate the stock with the "Greatest % increase", 
' "Greatest % Decrease" and "Greatest total volume".


Sub hard()

    'loop through each worksheet
    for each ws in worksheets

        'declare variables
        Dim closeprice As Double
        Dim openprice As Double
        Dim ticker As String
        Dim summary_table_row As Integer
        Dim ticker_volume As Double
        Dim counter As Integer
        Dim index As Long
        Dim yearlychange As Double
        Dim percentchange As Double

        ' initialize values to the below counter objects
        summary_table_row = 2
        ticker_volume = 0
        counter = 0

        ' print headers in the summary table
        ws.Range("i1").Value = "ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"

        ' derive the last row index
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            'loop through the daily ticker volumes
            For i = 2 To lastrow
                'counter variable stores the number of iterations carried out for each ticker name
                'this value is required to find the first row index(open price) for each ticker name
                counter = counter + 1

                ' if contitional to check if two consecutive cells have the same ticker name
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    'index is the first row index for each ticker
                    index = (i - counter) + 1
                    'use index to arrive at the open price for each ticker
                    openprice = ws.Cells(index, 3).Value
                    closeprice = ws.Cells(i, 6).Value
                    yearlychange = closeprice - openprice
                    ticker = ws.Cells(i, 1).Value

                    ' program will throw a division error if it encounters a zero value as denominator
                    If openprice = 0 then
                        percentchange = 0
                    Else
                        percentchange = closeprice / openprice
                    
                    end if

                    ' assign values to the summary table
                    ws.Range("i" & summary_table_row) = ticker
                    ws.Range("j" & summary_table_row) = yearlychange
                    ws.Range("k" & summary_table_row) = percentchange
                    ws.Range("l" & summary_table_row) = ticker_volume + ws.Cells(i, 7).Value

                    ' increment the summary table row variable by 1 to add the next ticker name
                    summary_table_row = summary_table_row + 1

                    ' reset the counter values to 0
                    ticker_volume = 0
                    counter = 0
                    index = 0

                Else
                    ticker_volume = ticker_volume + ws.Cells(i, 7).Value

                End If

            Next i

        ' applying conditional formatting to the summary table

        ' derive the last row index of the summary table
        lastrow_summary = ws.Cells(Rows.Count, 11).End(xlUp).Row

        ' apply the percent numberformat to the the percentchange column
            For Row = 2 To lastrow_summary
                ws.Cells(Row, 11).NumberFormat = "0.00%"
            Next Row

        ' apply conditional formatting to the yearlychange column
            For Row = 2 to lastrow_summary
                
                if ws.cells(row,10) < 0 then 
                    ws.cells(row,10).interior.colorindex = 3 ' negative change in red

                elseif ws.cells(row,10) > 0 then
                    ws.cells(row,10).interior.colorindex = 4 ' positive change in green

                end if

            next Row 


        ' final summary table
        ' declare new variables for the final summary table
        dim max_value as Double 
        dim max_ticker as String
        dim min_value as Double
        dim min_ticker as String
        dim max_volume as Double
        dim max_volume_ticker as String

        ' assign values the newly created variables
        max_value = ws.cells(2, 11).value
        max_ticker = ws.cells(2, 9).value
        min_value = ws.cells(2, 11).value
        min_ticker = ws.cells(2, 9).value
        max_volume = ws.cells(2, 12).value
        max_volume_ticker = ws.cells(2, 9).value

        ' print headers and description in the final summary table
        ws.Range("p1").Value = "ticker"
        ws.Range("q1").Value = "value"
        ws.Range("o2").value = "Greatest % Increase"
        ws.Range("o3").value = "Greatest % Decrease"
        ws.Range("o4").value = "Greates Total Volume"
 

            ' find the ticker that had the greatest % increase
            for row = 2 to lastrow_summary

                if ws.cells(row, 11).value > max_value then
                    max_value = ws.cells(row, 11).value
                    max_ticker = ws.cells(row, 9).value
                end if
            next row

            ' find the ticker that had the greatest % decrease
            for row = 2 to lastrow_summary

                if ws.cells(row, 11).value < min_value then
                    min_value = ws.cells(row, 11).value
                    min_ticker = ws.cells(row, 9).value
                end if
            next row


            'print the values in the final summary table
            ws.range("p2").value = max_ticker
            ws.range("q2").value = max_value
            ws.range("p3").value = min_ticker
            ws.range("q3").value = min_value


            ' find the ticker that had the greatest volume
            for row = 2 to lastrow_summary

                if ws.cells(row, 12).value > max_volume then
                    max_volume = ws.cells(row, 12).value
                    max_volume_ticker = ws.cells(row, 9).value

                end if

            next row

            'print the values in the final summary table
            ws.range("p4").value = max_volume_ticker
            ws.range("q4").value = max_volume

            'finally format the summary table
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(4, 17).NumberFormat = "#,##0.00_);(#,##0.00)"

    next

  
End Sub

