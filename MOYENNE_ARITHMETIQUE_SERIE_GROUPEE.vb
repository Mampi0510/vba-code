Function MOYENNE_SERIE_GROUPEE(plage1 As Range, plage2 As Range, cellule3 As Range) As Double
    Dim cellule1 As Range
    Dim cellule2 As Range
    Dim produitxini As Double
    Dim effectifTotal As Double

    Set cellule2 = plage2.Cells(1)

    For Each cellule1 In plage1
        If IsNumeric(cellule1.Value) And IsNumeric(cellule2.Value) Then
            produitxini = produitxini + (cellule1.Value * cellule2.Value)
        End If
        Set cellule2 = cellule2.Offset(1, 0) ' Avancer cellule2 à la cellule suivante
    Next cellule1
    
        If IsNumeric(cellule3.Value) Then
            effectifTotal = effectifTotal + cellule3.Value
        End If
        
    MOYENNE_SERIE_GROUPEE = produitxini / effectifTotal
End Function


