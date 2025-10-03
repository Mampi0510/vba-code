Function FISHER_SG(Ni As Range, Xi As Range) As Double
    Dim somme_Ni As Double
    Dim somme_NiXi As Double
    Dim somme_NiXi2 As Double
    Dim moyenne_X As Double
    Dim variance As Double
    Dim ecart_type As Double
    Dim mu3 As Double
    Dim n As Long
    Dim i As Long
    
    somme_Ni = 0
    somme_NiXi = 0
    somme_NiXi2 = 0
    mu3 = 0
    n = Xi.Rows.Count

    For i = 1 To n
        somme_Ni = somme_Ni + Ni.Cells(i, 1).Value
        somme_NiXi = somme_NiXi + Ni.Cells(i, 1).Value * Xi.Cells(i, 1).Value
    Next i
    moyenne_X = somme_NiXi / somme_Ni

    For i = 1 To n
        somme_NiXi2 = somme_NiXi2 + Ni.Cells(i, 1).Value * (Xi.Cells(i, 1).Value - moyenne_X) ^ 2
    Next i
    variance = somme_NiXi2 / somme_Ni
    ecart_type = Sqr(variance)

    For i = 1 To n
        mu3 = mu3 + Ni.Cells(i, 1).Value * (Xi.Cells(i, 1).Value - moyenne_X) ^ 3
    Next i
    mu3 = mu3 / somme_Ni
    
    FISHER_SG = mu3 / (ecart_type ^ 3)
End Function

