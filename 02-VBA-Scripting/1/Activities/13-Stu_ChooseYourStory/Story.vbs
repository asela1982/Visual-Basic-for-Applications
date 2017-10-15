Sub Game()

    'declare variables
    Dim InputValue As Integer

    'assign values to the variables
    InputValue = InputBox("Enter a value between 1-4")
    
    'declare a conditional statement
    If (InputValue = 1) Then
    MsgBox ("You choose to enter the wooded forest of doom!")
    
    ElseIf (InputValue = 2) Then
    MsgBox ("You choose to enter the fiery volcano of doom!")
    
    ElseIf (InputValue = 3) Then
    MsgBox ("You choose to enter the terrifying jungle of doom!")

    Else
    MsgBox ("You decide to stay home instead")
    
    End If

End Sub
