Function VARIANCE_SERIE_BRUTE(Xi As Range) As Double
    Dim somme_ecart As Double
    Dim moyenne_X As Double
    Dim n As Long
    Dim i As Long
    
    somme_ecart = 0
    n = Xi.Rows.Count
  
    For i = 1 To n
        moyenne_X = moyenne_X + Xi.Cells(i, 1).Value
    Next i
    moyenne_X = moyenne_X / n
   
    For i = 1 To n
        somme_ecart = somme_ecart + (Xi.Cells(i, 1).Value - moyenne_X) ^ 2
    Next i
    
    
    VARIANCE_SERIE_BRUTE = somme_ecart / n
End Function


