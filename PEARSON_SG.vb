Function PEARSON_SG(Ni As Range, Xi As Range) As Double
    Dim n As Long, i As Long
    Dim somme_Ni As Double, somme_NiXi As Double
    Dim moyenne As Double, var As Double, ecart_type As Double
    Dim Mediane As Double
    Dim addition As Double, cible As Double
    
    n = Xi.Rows.Count

    For i = 1 To n
        somme_Ni = somme_Ni + Ni.Cells(i, 1).Value
        somme_NiXi = somme_NiXi + Ni.Cells(i, 1).Value * Xi.Cells(i, 1).Value
    Next i
    
    moyenne = somme_NiXi / somme_Ni

    For i = 1 To n
        var = var + Ni.Cells(i, 1).Value * (Xi.Cells(i, 1).Value - moyenne) ^ 2
    Next i
    ecart_type = Sqr(var / somme_Ni)
    addition = 0
    cible = somme_Ni / 2
    
    For i = 1 To n
        addition = addition + Ni.Cells(i, 1).Value
        If addition >= cible Then
            Mediane = Xi.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    PEARSON_SG = 3 * (moyenne - Mediane) / ecart_type
End Function

