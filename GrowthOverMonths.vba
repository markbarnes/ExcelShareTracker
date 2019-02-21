Function GrowthOverMonths(ISIN As String, PurchaseDate As Date, EndDate As Date, NumMonths As Integer, CurrentValue As Double) As Variant

    Dim StartDate As Date, StartValue As Double, EndValue As Double
    
    If EndDate = 0 Then
        EndDate = Date
    End If
    StartDate = WorksheetFunction.EDate(EndDate, -NumMonths)
    If (StartDate < PurchaseDate) Then
        GrowthOverMonths = 0
    Else
        StartValue = GetPrice(ISIN, StartDate)
        EndValue = GetPrice(ISIN, EndDate)
        GrowthOverMonths = (EndValue - StartValue) / StartValue
    End If
End Function
