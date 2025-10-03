Function COEFFICIENT_YULE_SG(plageValeurs As Range, plageEffectifs As Range) As Double
    Dim n As Long, i As Long
    Dim total As Long, addition As Long
    Dim Q1 As Double, Q2 As Double, Q3 As Double
    
    ' Calculer effectif total
    For i = 1 To plageEffectifs.Count
        total = total + plageEffectifs.Cells(i, 1).Value
    Next i
    
    addition = 0
    For i = 1 To plageValeurs.Count
        addition = addition + plageEffectifs.Cells(i, 1).Value
        
        If addition >= total / 4 And Q1 = 0 Then Q1 = plageValeurs.Cells(i, 1).Value
        If addition >= total / 2 And Q2 = 0 Then Q2 = plageValeurs.Cells(i, 1).Value
        If addition >= 3 * total / 4 And Q3 = 0 Then Q3 = plageValeurs.Cells(i, 1).Value
    Next i
    
    COEFFICIENT_YULE_SG = (Q1 + Q3 - 2 * Q2) / (Q3 - Q1)
End Function

