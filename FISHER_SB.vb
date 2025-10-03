Function FISHER_SB(Xi As Range) As Double
    Dim n As Long, i As Long
    Dim moyenne As Double, var As Double, mu3 As Double, d As Double
    
    n = Xi.Count
    
    For i = 1 To n
        moyenne = moyenne + Xi(i, 1).Value
    Next i
    moyenne = moyenne / n
    
    For i = 1 To n
        d = Xi(i, 1).Value - moyenne
        var = var + d ^ 2
        mu3 = mu3 + d ^ 3
    Next i
    
    var = var / n
    mu3 = mu3 / n

    FISHER_SB = mu3 / (var ^ 1.5)
End Function

