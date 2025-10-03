Function VARIANCE_SERIE_GROUPEE(Xi As Range, Ni As Range) As Double
    Dim somme_Ni As Double
    Dim somme_NiXi As Double
    Dim somme_NiXi2 As Double
    Dim moyenne_X As Double
    Dim n As Long
    Dim i As Long
    
    somme_Ni = 0
    somme_NiXi = 0
    somme_NiXi2 = 0
    n = Xi.Rows.Count
    
    For i = 1 To n
        somme_Ni = somme_Ni + Ni.Cells(i, 1).Value
        somme_NiXi = somme_NiXi + Ni.Cells(i, 1).Value * Xi.Cells(i, 1).Value
        somme_NiXi2 = somme_NiXi2 + Ni.Cells(i, 1).Value * (Xi.Cells(i, 1).Value ^ 2)
    Next i
    
    moyenne_X = somme_NiXi / somme_Ni
    
    VARIANCE_SERIE_GROUPEE = (somme_NiXi2 / somme_Ni) - (moyenne_X ^ 2)
End Function

