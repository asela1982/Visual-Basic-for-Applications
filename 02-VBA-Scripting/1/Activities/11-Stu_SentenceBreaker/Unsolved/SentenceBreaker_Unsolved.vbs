Sub SentenceBreaker()

    ' Retrieve the user sentence and store in variable
     Dim sentence As String
     sentence = Range("B1").Value

     MsgBox (sentence)
  

    ' Retrieve the user word numbers and store in variables

     Dim Word1 As Integer
     Dim Word2 As Integer
     Dim Word3 As Integer

     Word1 = Range("A4").Value
     Word2 = Range("A5").Value
     Word3 = Range("A6").Value

     MsgBox (Word1)
     MsgBox (Word2)
     MsgBox (Word3)

    
    ' Split the user's sentence into words
    Dim word() As String
    word = Split(sentence)

    ' Use the word numbers to retrieve the specific words in the sentence
    ' Remember to offset by the 0 index
    
    Cells(4, 2).Value = word(Word1 - 1)
    Cells(5, 2).Value = word(Word2 - 1)
    Cells(6, 2).Value = word(Word3 - 1)

    
End Sub
