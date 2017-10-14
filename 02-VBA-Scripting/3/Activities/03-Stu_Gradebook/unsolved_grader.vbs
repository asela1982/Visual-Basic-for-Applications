Sub Grader()

Dim grade As Integer

grade = Range("B2").Value


If (grade >= 90) Then
    Range("D2").Value = "A"
    Range("C2").Value = "Pass"
    Range("C2").Interior.ColorIndex = 4

ElseIf (grade >= 80) Then
    Range("D2").Value = "B"
    Range("C2").Value = "Pass"
    Range("C2").Interior.ColorIndex = 4

ElseIf (grade >= 70) Then
    Range("D2").Value = "C"
    Range("C2").Value = "Warning"
    Range("C2").Interior.ColorIndex = 6

ElseIf (grade < 70) Then
    Range("D2").Value = "F"
    Range("C2").Value = "Fail"
    Range("C2").Interior.ColorIndex = 3

End If

End Sub
