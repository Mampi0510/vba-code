Function MOYENNE_GEOMETRIQUE_SERIE_GROUPEE(plage1 As Range, plage2 As Range, cellule3 As Range) As Double
    Dim cellule1 As Range ' xi
    Dim cellule2 As Range ' ni
    Dim puissance As Double
    Dim effectifTotal As Integer
    Dim xi As Double
    Dim ni As Double
    puissance = 1
    
    Set cellule2 = plage2.Cells(1)
    
    For Each cellule1 In plage1
    xi = cellule1.Value
    ni = cellule2.Value
        If IsNumeric(xi) And IsNumeric(ni) Then
            puissance = puissance * ((xi) ^ (ni))
        End If
        Set cellule2 = cellule2.Offset(1, 0)
    Next cellule1
    
        If IsNumeric(cellule3.Value) Then
            effectifTotal = cellule3.Value
        End If
    
    MOYENNE_GEOMETRIQUE_SERIE_GROUPEE = (puissance) ^ (1 / effectifTotal)
End Function
