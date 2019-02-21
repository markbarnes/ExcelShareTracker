Function AnnualGrowth(StartValue As Double, EndValue As Double, StartDate As Date, EndDate As Date) As Double
    If EndDate = 0 Then
        EndDate = Date
    End If
    AnnualGrowth = (EndValue / StartValue) ^ (1 / ((EndDate - StartDate) / 365.25)) - 1
End Function
