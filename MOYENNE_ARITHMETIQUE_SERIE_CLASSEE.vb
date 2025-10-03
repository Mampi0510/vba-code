Function MOYENNE_ARITHMETIQUE_SERIE_CLASSEE(plage1 As Range, plage2 As Range, plage3 As Range, cellule4 As Range) As Double
    Dim cellule1 As Range ' borne inférieure
    Dim cellule2 As Range ' borne supérieure
    Dim cellule3 As Range ' effectif ni
    Dim ci As Double
    Dim somme_cini As Double
    Dim effectifTotal As Integer

    Set cellule2 = plage2.Cells(1)
    Set cellule3 = plage3.Cells(1)

    For Each cellule1 In plage1
        If IsNumeric(cellule1.Value) And IsNumeric(cellule2.Value) And IsNumeric(cellule3.Value) Then
            ci = (cellule1.Value + cellule2.Value) / 2
            somme_cini = somme_cini + (ci * cellule3.Value)
        End If
        Set cellule2 = cellule2.Offset(1, 0) ' Avancer cellule2 à la cellule suivante
        Set cellule3 = cellule3.Offset(1, 0)
    Next cellule1

        If IsNumeric(cellule4.Value) Then
            effectifTotal = effectifTotal + cellule4.Value
        End If

    MOYENNE_ARITHMETIQUE_SERIE_CLASSEE = somme_cini / effectifTotal
End Function


