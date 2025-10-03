Function MOYENNE_HARMONIQUE_SERIE_GROUPEE(plage1 As Range, plage2 As Range, cellule3 As Range) As Double
    Dim cellule1 As Range ' xi
    Dim cellule2 As Range ' ni
    Dim effectifTotal As Integer
    Dim xi As Double
    Dim ni As Long
    Dim somme_nixi As Double
    
    Set cellule2 = plage2.Cells(1)
    
    For Each cellule1 In plage1
    xi = cellule1.Value
    ni = cellule2.Value
        If IsNumeric(xi) And IsNumeric(ni) Then
            somme_nixi = somme_nixi + (ni / xi)
        End If
        Set cellule2 = cellule2.Offset(1, 0)
    Next cellule1
    
        If IsNumeric(cellule3.Value) Then
            effectifTotal = cellule3.Value
        End If
    
    MOYENNE_HARMONIQUE_SERIE_GROUPEE = effectifTotal / somme_nixi
End Function
                     

