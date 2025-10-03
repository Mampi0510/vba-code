Function EFFECTIF(plage As Range, total As Variant) As Integer
    Dim cell As Range
    Dim compteur As Integer
    compteur = 0
    
    For Each cell In plage
        If cell.Value = total Then
        compteur = compteur + 1
        End If
    Next cell
    
    EFFECTIF = compteur
End Function
