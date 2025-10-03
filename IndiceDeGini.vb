Sub IndiceDeGini()
    Dim feuille As Worksheet
    Dim Fi As Range, Di As Range
    Dim i As Long, n As Long
    Dim GINI As Double, Somme As Double
    
    Set feuille = ThisWorkbook.Sheets("Feuil1")   ' adapter le nom de la feuille

    Set Fi = feuille.Range("D14:D19")   ' Fi
    Set Di = feuille.Range("E14:E19")   ' Di
    
    n = Fi.Rows.Count
    Somme = 0

    For i = 1 To n - 1
        Somme = Somme + (Fi.Cells(i, 1).Value * Di.Cells(i + 1, 1).Value) _
                       - (Fi.Cells(i + 1, 1).Value * Di.Cells(i, 1).Value)
    Next i
    
    GINI = Somme

    feuille.Range("G4").Value = "Indice de Gini" ' modifiable
    feuille.Range("H4").Value = GINI             ' modifiable
End Sub


