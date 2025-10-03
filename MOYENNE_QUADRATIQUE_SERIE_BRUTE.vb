Function MOYENNE_QUADRATIQUE_SERIE_BRUTE(plage1 As Range, cellule2 As Range)
    Dim cellule1 As Range
    Dim xi As Integer
    Dim effectifTotal As Integer
    
    For Each cellule1 In plage1
        If IsNumeric(cellule1.Value) Then
            xi = xi + ((cellule1.Value) ^ 2)
        End If
    Next cellule1
    
    If IsNumeric(cellule2.Value) Then
        effectifTotal = cellule2.Value
    End If
    
    MOYENNE_QUADRATIQUE_SERIE_BRUTE = Sqr(xi / effectifTotal)
End Function

