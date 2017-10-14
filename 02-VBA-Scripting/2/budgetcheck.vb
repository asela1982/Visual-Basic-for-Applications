Sub BudgetCheck()

    'part 1
    'variable declaration
    
    Dim price As Double
    Dim fees As Double
    Dim budget As Double
    Dim total As Double
    
    ' variable assignment
    price = Range("f3").Value
    fees = Range("h3").Value
    budget = Range("c3").Value

    'Calculate the total after fees
    total = price * (1 + fees)
    
    ' Assign the total back to worksheet
    Range("l3").Value = total

    'part 2

    If (budget > total) Then
        MsgBox ("Under budget")
    Else
        MsgBox ("Over budget")
    End If
    
    'part 3
    Dim new_price As Double
    new_price = budget / (1 + fees)
    Range("f3").Value = new_price
    Range("l3").Value = new_price * (1 + fees)


End Sub
