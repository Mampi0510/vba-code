Function COEFFICIENT_YULE(plage As Range) As Double
    Dim n As Long, i As Long
    Dim valeurs() As Double
    Dim Q1 As Double, Q2 As Double, Q3 As Double
    Dim total As Long, cumule As Long
    
    For Each cellule In plage
        If IsNumeric(cellule.Value) Then
            n = n + 1
            ReDim Preserve valeurs(1 To n)
            valeurs(n) = cellule.Value
        End If
    Next cellule
    
    Dim j As Long, t As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If valeurs(i) > valeurs(j) Then
                t = valeurs(i)
                valeurs(i) = valeurs(j)
                valeurs(j) = t
            End If
        Next j
    Next i
    
    Q1 = valeurs(Int((n + 1) / 4))
    Q2 = valeurs(Int((n + 1) / 2))
    Q3 = valeurs(Int(3 * (n + 1) / 4))
    
    COEFFICIENT_YULE = (Q1 + Q3 - 2 * Q2) / (Q3 - Q1)
End Function

