Function APPLATISSEMENT_SG(Ni As Range, Xi As Range) As Double
    Dim n As Long, i As Long
    Dim sommeNi As Double, sommeNiXi As Double
    Dim moyenne As Double, var As Double, mu4 As Double
    
    sommeNi = 0
    sommeNiXi = 0
    n = Xi.Rows.Count
    
    For i = 1 To n
        sommeNi = sommeNi + Ni.Cells(i, 1).Value
        sommeNiXi = sommeNiXi + Ni.Cells(i, 1).Value * Xi.Cells(i, 1).Value
    Next i
    moyenne = sommeNiXi / sommeNi
    
    var = 0
    mu4 = 0
    For i = 1 To n
        var = var + Ni.Cells(i, 1).Value * (Xi.Cells(i, 1).Value - moyenne) ^ 2
        mu4 = mu4 + Ni.Cells(i, 1).Value * (Xi.Cells(i, 1).Value - moyenne) ^ 4
    Next i
    
    var = var / sommeNi
    mu4 = mu4 / sommeNi
    
    APPLATISSEMENT_SG = mu4 / (var ^ 2)
End Function

