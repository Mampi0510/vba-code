Function MOYENNE_QUADRATIQUE_SERIE_CLASSEE(plage1 As Range, plage2 As Range, plage3 As Range, cellule4 As Range)
    Dim cellule1 As Range ' bi
    Dim cellule2 As Range ' bs
    Dim cellule3 As Range ' ni
    Dim ni As Integer
    Dim ci As Double
    Dim somme_cini As Double
    Dim effectifTotal As Integer
    
    Set cellule2 = plage2.Cells(1)
    Set cellule3 = plage3.Cells(1)
    
    For Each cellule1 In plage1
        If IsNumeric(cellule1.Value) And IsNumeric(cellule2.Value) And IsNumeric(cellule3.Value) Then
            ci = ((cellule1.Value + cellule2.Value) / 2) ^ 2
            ni = cellule3.Value
            somme_cini = somme_cini + (ci * ni)
        End If
        Set cellule2 = cellule2.Offset(1, 0)
        Set cellule3 = cellule3.Offset(1, 0)
    Next cellule1
    
    If IsNumeric(cellule4.Value) Then
        effectifTotal = cellule4.Value
    End If
    
    MOYENNE_QUADRATIQUE_SERIE_CLASSEE = Sqr(somme_cini / effectifTotal)
End Function



