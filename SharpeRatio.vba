Function SharpeRatio(PricesRange As Range)

    ' Calculate monthly growth
    Dim ActualGrowth As Double, BondGrowth As Double, StartingBondPrice As Double, EndingBondPrice As Double, StartDate As Date, EndDate As Date, EarliestDate As Date, LatestDate As Date
    Dim CellCount As Integer, PreviousValue As Double, cell As Range, sum As Integer, ExcessReturns(), SumOfExcessReturns As Double, RowNum As Integer
    ReDim ExcessReturns(1 To 999)
    CellCount = 0
    SumOfExcessReturns = 0
    RowNum = PricesRange.Row
    EarliestDate = #1/1/1900#
    For Each cell In PricesRange
        If Val(cell.Value) <> 0 Then
            RowNum = cell.Row
            If EarliestDate = #1/1/1900# Then
                EarliestDate = PricesRange.Worksheet.Cells(RowNum, 2)
            End If
            LatestDate = PricesRange.Worksheet.Cells(RowNum, 2)
            CellCount = CellCount + 1
        End If
    Next cell
    StartingBondPrice = GetPrice("IE00B1S75374", EarliestDate)
    EndingBondPrice = GetPrice("IE00B1S75374", LatestDate)
    BondGrowth = (((EndingBondPrice / StartingBondPrice) ^ (1 / CellCount)) - 1) * 100
    CellCount = 0
    For Each cell In PricesRange
        If Val(cell.Value) <> 0 Then
            CellCount = CellCount + 1
            If (CellCount > 1) Then
                RowNum = cell.Row
                StartDate = PricesRange.Worksheet.Cells(RowNum - 1, 2)
                EndDate = PricesRange.Worksheet.Cells(RowNum, 2)
                ActualGrowth = (cell.Value - PreviousValue) / PreviousValue * 100
                ExcessReturns(CellCount - 1) = ActualGrowth - BondGrowth
                SumOfExcessReturns = SumOfExcessReturns + ExcessReturns(CellCount - 1)
            End If
            PreviousValue = cell.Value
        End If
    Next cell
    ReDim Preserve ExcessReturns(1 To CellCount - 1)
    
    'Calculate Standard Deviation
    Dim avg As Double, SumSq As Double, StdDev As Double, i As Integer
    avg = SumOfExcessReturns / (CellCount - 1)
    For i = 1 To (CellCount - 1)
        SumSq = SumSq + (ExcessReturns(i) - avg) ^ 2
    Next i
    StdDev = Sqr(SumSq / (CellCount - 2))
    
    SharpeRatio = avg / StdDev
End Function
