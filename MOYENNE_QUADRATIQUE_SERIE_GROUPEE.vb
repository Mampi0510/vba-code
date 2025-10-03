Function MOYENNE_QUADRATIQUE_SERIE_GROUPEE(plage1 As Range, plage2 As Range, cellule3 As Range)
    Dim cellule1 As Range ' xi
    Dim cellule2 As Range ' ni
    Dim ni As Integer
    Dim xi As Integer
    Dim somme_xini As Long
    Dim effectifTotal As Integer
    
    Set cellule2 = plage2.Cells(1)
    
    For Each cellule1 In plage1
        If IsNumeric(cellule1.Value) And IsNumeric(cellule2.Value) Then
            xi = (cellule1.Value) ^ 2
            ni = cellule2.Value
            somme_xini = somme_xini + (xi * ni)
        End If
        Set cellule2 = cellule2.Offset(1, 0)
    Next cellule1
    
    If IsNumeric(cellule3.Value) Then
        effectifTotal = cellule3.Value
    End If
    
    MOYENNE_QUADRATIQUE_SERIE_GROUPEE = Sqr(somme_xini / effectifTotal)
End Function


