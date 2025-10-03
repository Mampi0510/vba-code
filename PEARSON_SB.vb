Function PEARSON_SB(X As Range) As Double
    Dim moyenne As Double, ecart_type As Double, Mediane As Double
    
    moyenne = Application.WorksheetFunction.Average(X)
    ecart_type = Application.WorksheetFunction.StDev_P(X)
    Mediane = Application.WorksheetFunction.Median(X)
    
    PEARSON_SB = 3 * (moyenne - Mediane) / ecart_type
End Function

