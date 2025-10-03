Function APPLATISSEMENT_SB(X As Range) As Double
    Dim n As Long, i As Long
    Dim moyenne As Double, var As Double, mu4 As Double
    Dim somme As Double
    
    n = X.Count
    somme = 0
    For i = 1 To n
        somme = somme + X.Cells(i, 1).Value
    Next i
    moyenne = somme / n
    
    var = 0
    mu4 = 0
    For i = 1 To n
        var = var + (X.Cells(i, 1).Value - moyenne) ^ 2
        mu4 = mu4 + (X.Cells(i, 1).Value - moyenne) ^ 4
    Next i
    
    var = var / n
    mu4 = mu4 / n
    
    APPLATISSEMENT_SB = mu4 / (var ^ 2)
End Function

