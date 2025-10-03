Function PEARSON_SC(Bi As Range, Bs As Range, Ni As Range) As Double
    Dim i As Long, n As Long
    Dim sommeNi As Double, sommeNiCi As Double, sommeNiCi2 As Double
    Dim moyenne As Double, variance As Double, ecart_type As Double
    Dim ci As Double, L As Double, h As Double
    Dim cum As Double, CFprec As Double, Mediane As Double
    
    n = Ni.Rows.Count

    For i = 1 To n
        ci = (Bi(i, 1).Value + Bs(i, 1).Value) / 2
        sommeNi = sommeNi + Ni(i, 1).Value
        sommeNiCi = sommeNiCi + Ni(i, 1).Value * ci
        sommeNiCi2 = sommeNiCi2 + Ni(i, 1).Value * ci ^ 2
    Next i
    
    moyenne = sommeNiCi / sommeNi
    variance = (sommeNiCi2 / sommeNi) - moyenne ^ 2
    If variance <= 0 Then
        PEARSON_SC = CVErr(xlErrDiv0)
        Exit Function
    End If
    ecart_type = Sqr(variance)

    cum = 0
    For i = 1 To n
        cum = cum + Ni(i, 1).Value
        If cum >= sommeNi / 2 Then
            L = Bi(i, 1).Value
            h = Bs(i, 1).Value - Bi(i, 1).Value
            CFprec = cum - Ni(i, 1).Value
            Mediane = L + ((sommeNi / 2 - CFprec) / Ni(i, 1).Value) * h
            Exit For
        End If
    Next i

    PEARSON_SC = 3 * (moyenne - Mediane) / ecart_type
End Function

