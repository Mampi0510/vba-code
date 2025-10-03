Function COEFFICIENT_YULE_SC(plageBi As Range, plageBs As Range, plageEffectifs As Range) As Double
    Dim n As Long, i As Long
    Dim total As Long, addition As Long
    Dim Q1 As Double, Q2 As Double, Q3 As Double
    Dim position1 As Double, position2 As Double, position3 As Double
    
    ' Effectif total
    For i = 1 To plageEffectifs.Count
        total = total + plageEffectifs.Cells(i, 1).Value
    Next i
    
    position1 = total / 4
    position2 = total / 2
    position3 = 3 * total / 4
    
    addition = 0
    For i = 1 To plageEffectifs.Count
        addition = addition + plageEffectifs.Cells(i, 1).Value
        If addition >= position1 Then
            Q1 = plageBi.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    addition = 0
    For i = 1 To plageEffectifs.Count
        addition = addition + plageEffectifs.Cells(i, 1).Value
        If addition >= position2 Then
            Q2 = plageBi.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    addition = 0
    For i = 1 To plageEffectifs.Count
        addition = addition + plageEffectifs.Cells(i, 1).Value
        If addition >= position3 Then
            Q3 = plageBi.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    COEFFICIENT_YULE_SC = (Q1 + Q3 - 2 * Q2) / (Q3 - Q1)
End Function

