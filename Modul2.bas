Attribute VB_Name = "Modul2"
Dim Something As Boolean
Sub Eatsom()
  Call InputECG
End Sub
Private Sub InputECG()
'Inputboxen um das Kommen, �berstunden, Mittagessen und Wunschgehen einzutragen
Range("D20") = InputBox("Wann gehst du essen?")
Range("G21") = InputBox("Wann willst du gehen ?")
Range("N2") = InputBox("Wie viele �berstunden hast du?")
End Sub

