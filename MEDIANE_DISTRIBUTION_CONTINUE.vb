Function MEDIANE_DISTRIBUTION_CONTINUE(plage1 As Range, plage2 As Range, plage3 As Range) As Double
    Dim cellule1 As Range ' borne inférieure
    Dim cellule2 As Range ' borne supérieure
    Dim cellule3 As Range ' fréquence cumulée
    Dim L As Double, h As Double
    Dim Fprec As Double, fmed As Double
    Dim Med As Double
    
    Set cellule2 = plage2.Cells(1)
    Set cellule3 = plage3.Cells(1)
    
    For Each cellule1 In plage1
        If IsNumeric(cellule1.Value) And IsNumeric(cellule2.Value) And IsNumeric(cellule3.Value) Then
            If cellule3.Value >= 0.5 Then
                L = cellule1.Value
                h = cellule2.Value - cellule1.Value
                If cellule3.Row > plage3.Cells(1).Row Then
                    Fprec = cellule3.Offset(-1, 0).Value
                Else
                    Fprec = 0
                End If
                fmed = cellule3.Value - Fprec
                
                Med = L + ((0.5 - Fprec) / fmed) * h
                MEDIANE_DISTRIBUTION_CONTINUE = Med
                Exit Function
            End If
        End If
        Set cellule2 = cellule2.Offset(1, 0)
        Set cellule3 = cellule3.Offset(1, 0)
    Next cellule1
End Function

