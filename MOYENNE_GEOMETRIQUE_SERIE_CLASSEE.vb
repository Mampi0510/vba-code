Function MOYENNE_GEOMETRIQUE_SERIE_CLASSEE(plage1 As Range, plage2 As Range, plage3 As Range) As Double
    Dim cellule1 As Range ' bi
    Dim cellule2 As Range ' bs
    Dim cellule3 As Range ' ni
    Dim ci As Double
    Dim ni As Long
    Dim puissance As Double
    Dim effectifTotal As Long
    puissance = 1
    
    Set cellule2 = plage2.Cells(1)
    Set cellule3 = plage3.Cells(1)
    
    For Each cellule1 In plage1
        ci = ((cellule1.Value) + (cellule2.Value)) / 2
        ni = cellule3.Value
        If IsNumeric(ci) And IsNumeric(ni) Then
            puissance = puissance * ((ci) ^ (ni))
        End If
        Set cellule2 = cellule2.Offset(1, 0)
        Set cellule3 = cellule3.Offset(1, 0)
    Next cellule1
    
    For Each cellule3 In plage3
        If IsNumeric(cellule3.Value) Then
            effectifTotal = effectifTotal + cellule3.Value
        End If
    Next cellule3
    
    MOYENNE_GEOMETRIQUE_SERIE_CLASSEE = (puissance) ^ (1 / effectifTotal)
End Function
