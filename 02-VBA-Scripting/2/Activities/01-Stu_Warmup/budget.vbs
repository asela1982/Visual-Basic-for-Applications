
Sub budget()

    ' Part 1: Calculate the total after fees and enter the value in the "Total" cell.
    
    ' declare the variables
    Dim budget As Double
    Dim price As Double
    Dim fees As Double
    Dim total As Double
    
    ' assign values to the variables
    
    budget = Range("c3").Value
    price = Range("f3").Value
    fees = Range("h3").Value
    
    ' calculate the total after fees
    
    total = price + (price * fees)
    
    ' enter the value in the "Total" cell.
    
    Range("l3").Value = total
    
    ' Part II: Create a Message Box for the user to designate whether the amount including fees is
    ' within or over budget.

    ' Part III (Challenge): If the total is over budget, correct the price such that it fits within
    ' the max of the user's budget. Be sure to round down! (Example: If the user's budget is 100 and
    ' the fees are 15%, the max price should 86)
    
    If total > budget Then
        MsgBox "over budget"
        Dim newprice as Double
        newprice = Application.WorksheetFunction.RoundDown(budget / (1 + fees), 0)
        Range("l3").Value = newprice + (newprice * fees)
    
    Else
        MsgBox "under budget"
    
    End If
    
End Sub

