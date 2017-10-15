Sub Chess()

    Worksheets("Sheet1").Activate
    
   ' Top-half of the chess board use Ranges
   
    Range("A1").Value = "Rook"
    Range("H1").Value = "Rook"
    
    Range("B1").Value = "Knight"
    Range("G1").Value = "Knight"
    
    Range("C1").Value = "Bishop"
    Range("F1").Value = "Bishop"

    Range("D1").Value = "Queen"
    Range("E1").Value = "King"

    Range("A2", "H2").Value = "Pawn"
    
    ' Bottom-half of the chess board use Cells

    Cells(8, 1).Value = "Rook"
    Cells(8, 8).Value = "Rook"

    Cells(8, 2).Value = "Knight"
    Cells(8, 7).Value = "Knight"

    Cells(8, 3).Value = "Bishop"
    Cells(8, 6).Value = "Bishop"

    Cells(8, 4).Value = "Queen"
    Cells(8, 5).Value = "King"
    
    Range(Cells(7, 1), Cells(7, 8)).Value = "Pawn"
    

End Sub