Function MOYENNE_HARMONIQUE_SERIE_CLASSEE(plage1 As Range, plage2 As Range, plage3 As Range, cellule4 As Range) As Double
    Dim cellule1 As Range ' bi
    Dim cellule2 As Range ' bs
    Dim cellule3 As Range ' ni
    Dim effectifTotal As Integer
    Dim ci As Double
    Dim ni As Long
    Dim somme_nici As Double
    
    Set cellule2 = plage2.Cells(1)
    Set cellule3 = plage3.Cells(1)
    
    For Each cellule1 In plage1
        If IsNumeric(cellule1.Value) And IsNumeric(cellule2.Value) And IsNumeric(cellule3.Value) Then
            ni = cellule3.Value
            ci = (cellule1.Value + cellule2.Value) / 2
            somme_nici = somme_nici + (ni / ci)
        End If
        Set cellule2 = cellule2.Offset(1, 0)
        Set cellule3 = cellule3.Offset(1, 0)
    Next cellule1
    
        If IsNumeric(cellule4.Value) Then
            effectifTotal = cellule4.Value
        End If
    
    MOYENNE_HARMONIQUE_SERIE_CLASSEE = effectifTotal / somme_nici
End Function
                     


