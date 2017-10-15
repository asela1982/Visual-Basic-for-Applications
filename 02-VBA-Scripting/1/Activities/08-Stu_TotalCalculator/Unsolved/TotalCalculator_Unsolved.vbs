Sub Calculator()

    ' DIM is short for 'declare in memory'
    
    ' variable declaration
    
    Dim Inputprice As Double
    Dim TaxRate As Double
    Dim Quantity As Double
    Dim Total As Double
    
    ' variable assingment
    
    Inputprice = Range("B2").Value
    TaxRate = Range("C2").Value
    Quantity = Range("D2").Value
    
    Total = Inputprice * (1 + TaxRate) * Quantity
    
    ' populate the derived Total in the cell
    
    Cells(2, 5).Value = Total

End Sub
