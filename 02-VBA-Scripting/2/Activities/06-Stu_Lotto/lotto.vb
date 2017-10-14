Sub lotto()

Dim first, second, third As Long
Dim run1, run2, run3 As Long
Dim flag As integer

first = 3957481
second = 5865187
third = 2817729

run1 = 2275339
run2 = 5868182
run3 = 1841402

flag = 0

For i = 2 To 1001

    If Cells(i, 3).Value = first Then
        Cells(2, 6).Value = Cells(i, 1).Value
        Cells(2, 7).Value = Cells(i, 2).Value
        Cells(2, 8).Value = Cells(i, 3).Value

    ElseIf Cells(i, 3).Value = second Then
        Cells(3, 6).Value = Cells(i, 1).Value
        Cells(3, 7).Value = Cells(i, 2).Value
        Cells(3, 8).Value = Cells(i, 3).Value

    ElseIf Cells(i, 3).Value = third Then
        Cells(4, 6).Value = Cells(i, 1).Value
        Cells(4, 7).Value = Cells(i, 2).Value
        Cells(4, 8).Value = Cells(i, 3).Value

    ElseIf (Cells(i, 3).Value = run1 or Cells(i, 3).Value = run2 or Cells(i, 3).Value = run3) and flag = 0 Then
        Cells(5, 6).Value = Cells(i, 1).Value
        Cells(5, 7).Value = Cells(i, 2).Value
        Cells(5, 8).Value = Cells(i, 3).Value 
        flag = 1      

    End If

Next i

MsgBox "Congrats! " + Cells(2, 6).Value + " " + Cells(2, 7).Value




End Sub