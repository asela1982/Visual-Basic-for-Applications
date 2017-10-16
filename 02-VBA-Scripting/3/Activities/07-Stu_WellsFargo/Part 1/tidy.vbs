
Sub tidy()

        '-----------------------------------------------------
        ' loop through each worksheet
        '-----------------------------------------------------

    for each ws in worksheets  

        '-----------------------------------------------------
        ' insert the state
        '-----------------------------------------------------

        ' Extract words before the phrase "_Wells_Fargo" to figure out which State.
        Dim word() As String
        Dim headers() As String
        Dim currentName As String

        lastrow = ws.Range("A" & Rows.Count).End(xlUp).Row

        currentName = ws.Name
        word = Split(currentName, "_")

        ' Inserting a Column at Column A
        ws.Range("A1").EntireColumn.Insert

        ' Add the State to the first column of the worksheet.
        ws.Cells(1, 1) = "State"
        ws.Range("A2:A" & lastrow) = word(0)

        '-----------------------------------------------------
        ' extract the year
        '-----------------------------------------------------

        ' Convert the headers of each row to simply say the year.
        lastcol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        for j = 3 to lastcol
            header = split(ws.cells(1,j).value, " ")
            ws.cells(1,j).value = header(3)
        next j

        '-----------------------------------------------------
        ' Convert the numbers to currency values for all cells
        '-----------------------------------------------------

        ' re-calcuate the index for lastrow and lastcol
        lastrow = ws.Range("A" & Rows.Count).End(xlUp).Row
        lastcol = ws.Cells(1, Columns.Count).End(xlToLeft).Column

        for row = 2 to lastrow
            for col = 3 to lastcol
                'convert the format of the cell to currency
                ws.cells(row,col).style = "Currency"
            next col
        next row

    next ws

    Msgbox ("Tidy Completed!")

End Sub
