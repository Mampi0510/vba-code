Function APPLATISSEMENT_SC(Bi As Range, Bs As Range, Ni As Range) As Variant
    Dim n As Long, i As Long
    Dim SommeEffectifs As Double, moyenne As Double, var As Double, mu4 As Double
    Dim ci As Double, d As Double
    
    n = Ni.Rows.Count

    For i = 1 To n
        ci = (Bi(i, 1).Value + Bs(i, 1).Value) / 2
        SommeEffectifs = SommeEffectifs + Ni(i, 1).Value
        moyenne = moyenne + Ni(i, 1).Value * ci
    Next i
    
    If SommeEffectifs = 0 Then APPLATISSEMENT_SC = CVErr(xlErrDiv0): Exit Function
    moyenne = moyenne / SommeEffectifs

    For i = 1 To n
        ci = (Bi(i, 1).Value + Bs(i, 1).Value) / 2
        d = ci - moyenne
        var = var + Ni(i, 1).Value * d ^ 2
        mu4 = mu4 + Ni(i, 1).Value * d ^ 4
    Next i
    
    var = var / SommeEffectifs
    mu4 = mu4 / SommeEffectifs

    APPLATISSEMENT_SC = mu4 / (var ^ 2)
End Function

