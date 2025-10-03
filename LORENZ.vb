
Function FrequenceCumule(Ni As Range, n As Long) As Double
    Dim somme_Ni As Long
    Dim i As Long
    
    somme_Ni = 0
    
   
    For i = 1 To Ni.Rows.Count
        somme_Ni = somme_Ni + Ni.Cells(i, 1).Value
    Next i

    FrequenceCumule = somme_Ni / n
End Function


Function VariableCumulee(Ni As Range, Xi As Range, indiceMax As Long) As Double
    Dim somme_NiXi As Double
    Dim somme_cumulee As Double
    Dim i As Long

    
    somme_NiXi = 0
    somme_cumulee = 0

    
    For i = 1 To Ni.Rows.Count
        somme_NiXi = somme_NiXi + Ni.Cells(i, 1).Value * Xi.Cells(i, 1).Value
    Next i

    
    For i = 1 To indiceMax
        somme_cumulee = somme_cumulee + Ni.Cells(i, 1).Value * Xi.Cells(i, 1).Value
    Next i

    VariableCumulee = somme_cumulee / somme_NiXi
End Function

Sub CourbeDeLorenz()
    Dim feuille As Worksheet
    Dim tabObject As ChartObject
    Dim plageX As Range, plageY As Range

    Set feuille = ThisWorkbook.Sheets("Feuil1")   ' adapter le nom

    Set plageX = feuille.Range("D14:D18")   ' Fi (à definir prealablement)
    Set plageY = feuille.Range("E14:E18")   ' Di (à definir prealablement)

    Set tabObject = feuille.ChartObjects.Add(Left:=300, Top:=50, Width:=400, Height:=200)
    tabObject.Chart.ChartType = xlXYScatterLines

    With tabObject.Chart.SeriesCollection.NewSeries
        .XValues = plageX
        .Values = plageY
        .Name = "Courbe de Lorenz"
    End With

    With tabObject.Chart.SeriesCollection.NewSeries
        .XValues = Array(0, 1)
        .Values = Array(0, 1)
        .Name = "Égalité parfaite"
    End With
End Sub

