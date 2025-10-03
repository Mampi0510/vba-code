Function FISHER_SC(Bi As Range, Bs As Range, Ni As Range) As Double
    Dim n As Long, i As Long
    Dim ci As Double, moyenne As Double, var As Double, mu3 As Double, d As Double
    Dim somme_Ni As Double, somme_Nici As Double
    
    n = Ni.Rows.Count

    For i = 1 To n
        ci = (Bi.Cells(i, 1).Value + Bs.Cells(i, 1).Value) / 2
        somme_Ni = somme_Ni + Ni.Cells(i, 1).Value
        somme_Nici = somme_Nici + Ni.Cells(i, 1).Value * ci
    Next i
    moyenne = somme_Nici / somme_Ni

    For i = 1 To n
        ci = (Bi.Cells(i, 1).Value + Bs.Cells(i, 1).Value) / 2
        d = ci - moyenne
        var = var + Ni.Cells(i, 1).Value * d ^ 2
        mu3 = mu3 + Ni.Cells(i, 1).Value * d ^ 3
    Next i
    
    var = var / somme_Ni
    mu3 = mu3 / somme_Ni
    
    FISHER_SC = mu3 / (var ^ 1.5)
End Function

