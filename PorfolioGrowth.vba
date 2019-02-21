Function PorfolioGrowth(SharesRange As Range, Optional CalcYear As Integer) As Double
    
    Dim i, j As Integer

    'Count stocks
    Dim NumStocks As Integer, Columns(999)
    NumStocks = 0
    j = 1
    Do While (SharesRange.Cells(1, (j * 2) - 1).Value) <> ""       ' Loop until the stock name is empty
        If (SharesRange.Cells(3, (j * 2)).Value) <> "" Then        ' Only add stocks that have a purchase date
            NumStocks = NumStocks + 1
            Columns(NumStocks) = (j * 2) - 1                       ' Column number of this stock
        End If
        j = j + 1
    Loop
    
    'Create arrays for stock data
    Dim Stocks(), ColNum As Integer
    ReDim Stocks(1 To NumStocks, 1 To 8)
    For j = 1 To NumStocks
        Stocks(j, 1) = SharesRange.Cells(1, (j * 2) - 1).Value 'Name
        Stocks(j, 2) = SharesRange.Cells(5, j * 2).Value 'Purchase Amount
        Stocks(j, 3) = SharesRange.Cells(3, j * 2).Value 'Purchase Date
        Stocks(j, 4) = SharesRange.Cells(7, j * 2).Value 'Sold Date
        Stocks(j, 5) = SharesRange.Cells(9, j * 2).Value 'Current/Final Value
        Stocks(j, 6) = SharesRange.Cells(4, j * 2).Value 'Quantity
        Stocks(j, 7) = SharesRange.Cells(2, (j * 2) - 1).Value 'ISIN
        Stocks(j, 8) = SharesRange.Cells(6, j * 2).Value 'Dividends
        If Stocks(j, 4) = "" Then
            Stocks(j, 4) = Date
        End If
        Stocks(j, 8) = Stocks(j, 8) / (Stocks(j, 4) - Stocks(j, 3)) ' Daily Dividend
    Next j
    Erase Columns
    
    'Throw out invalid shares and create array of all dates
    Dim NewStocks(), NewStocksCounter As Integer, AllDates() As Date, AllDatesCounter As Long
    ReDim NewStocks(1 To NumStocks, 1 To 8)
    ReDim AllDates(1 To (NumStocks * 2) + 3)
    AllDatesCounter = 0
    NewStocksCounter = 0
    For j = 1 To NumStocks
        If Not (IsEmpty(Stocks(j, 2)) Or IsEmpty(Stocks(j, 3))) Then
            AllDatesCounter = AllDatesCounter + 1
            AllDates(AllDatesCounter) = Stocks(j, 3)
            AllDatesCounter = AllDatesCounter + 1
            AllDates(AllDatesCounter) = Stocks(j, 4)
            NewStocksCounter = NewStocksCounter + 1
            For i = 1 To 8
                NewStocks(NewStocksCounter, i) = Stocks(j, i)
            Next i
        End If
    Next j
    If NumStocks <> NewStocksCounter Then
        NumStocks = NewStocksCounter
        ReDim Stocks(1 To NumStocks, 1 To 8)
        For j = 1 To NumStocks
            For i = 1 To 8
                Stocks(j, i) = NewStocks(j, i)
            Next i
        Next j
    End If
    Erase NewStocks
    ReDim Preserve AllDates(1 To AllDatesCounter)
    QuickSort AllDates, 1, AllDatesCounter
    
    'Add beginning and end if calculating a single year
    Dim LastDate As Date
    LastDate = SharesRange.Worksheet.Cells(Rows.Count, 3).End(xlUp).Value
    If (CalcYear <> 0) Then
        If (AllDates(AllDatesCounter) < LastDate And Year(LastDate) = CalcYear) Then
            AllDatesCounter = AllDatesCounter + 1
            ReDim Preserve AllDates(1 To AllDatesCounter)
            AllDates(AllDatesCounter) = LastDate
        End If
        If (AllDates(AllDatesCounter) > DateValue("31 Dec " & CStr(CalcYear))) Then
            AllDatesCounter = AllDatesCounter + 1
            ReDim Preserve AllDates(1 To AllDatesCounter)
            AllDates(AllDatesCounter) = DateValue("31 Dec " & CStr(CalcYear))
        End If
        If (AllDates(1) < DateValue("1 Jan " & CStr(CalcYear))) Then
            AllDatesCounter = AllDatesCounter + 1
            ReDim Preserve AllDates(1 To AllDatesCounter)
            AllDates(AllDatesCounter) = DateValue("1 Jan " & CStr(CalcYear))
        End If
    Else
        AllDatesCounter = AllDatesCounter + 1
        ReDim Preserve AllDates(1 To AllDatesCounter)
        AllDates(AllDatesCounter) = LastDate
    End If
    QuickSort AllDates, 1, AllDatesCounter
    
    'Remove duplicate dates or dates not relevant to this period
    Dim UniqueDates() As Date, OldDate As Date, UniqueDatesCount As Integer
    ReDim UniqueDates(1 To AllDatesCounter)
    UniqueDatesCount = 0
    OldDate = DateValue("1 Jan 1900")
    For i = 1 To AllDatesCounter
        If (OldDate <> AllDates(i)) And AllDates(i) > 0 Then
            If ((CalcYear = 0) Or (Year(AllDates(i)) = CalcYear)) Then
                UniqueDatesCount = UniqueDatesCount + 1
                UniqueDates(UniqueDatesCount) = AllDates(i)
                OldDate = AllDates(i)
            End If
        End If
    Next i
    ReDim Preserve UniqueDates(1 To UniqueDatesCount)
    
    'Calculate fees during this period
    Dim FeeStartDate, FeeEndDate As String
    Dim LastFeeRow As String
    Dim Fees As Double
    LastFeeRow = SharesRange.Worksheet.Cells(Rows.Count, 2).End(xlUp).Row
    If (CalcYear = 0) Then
        FeeStartDate = "01/01/" + CStr(Year(UniqueDates(1)))
        FeeEndDate = "01/01/" + CStr(Year(UniqueDates(UniqueDatesCount)) + 1)
    Else
        FeeStartDate = "01/01/" + CStr(CalcYear)
        FeeEndDate = "01/01/" + CStr(CalcYear + 1)
    End If
    Fees = WorksheetFunction.SumIfs(SharesRange.Worksheet.Range("A22:A" + LastFeeRow), SharesRange.Worksheet.Range("B22:B" + LastFeeRow), "<" + FeeEndDate, SharesRange.Worksheet.Range("B22:B" + LastFeeRow), ">=" + FeeStartDate)
    
    'Calculate growth for each stock during each interval
    Dim ThisDate As Date, PreviousDate As Date, StockValueAtStart As Double, StockValueAtEnd As Double, IntervalGains() As Double, NumDays As Integer
    ReDim IntervalGains(2 To UniqueDatesCount)
    NumDays = UniqueDates(UniqueDatesCount) - UniqueDates(1)
    PreviousDate = UniqueDates(1)
    For i = 2 To UniqueDatesCount
        ThisDate = UniqueDates(i)
        StockValueAtStart = Fees / NumDays * (ThisDate - PreviousDate)
        StockValueAtEnd = 0
        For j = 1 To NumStocks
            If (Stocks(j, 3) < ThisDate) And (Stocks(j, 4) >= ThisDate) Then 'Stock is owned on this date, but wasn't bought on this date, so proceed
                If PreviousDate = Stocks(j, 3) Then ' Bought on this day
                    StockValueAtStart = StockValueAtStart + Stocks(j, 2)
                Else
                    If IsEmpty(Stocks(j, 7)) Then
                        StockValueAtStart = StockValueAtStart + Stocks(j, 2)
                    Else
                        StockValueAtStart = StockValueAtStart + (GetPrice(CStr(Stocks(j, 7)), PreviousDate) * Stocks(j, 6))
                    End If
                End If
                If ThisDate = Stocks(j, 4) Then 'Sold on this day
                    StockValueAtEnd = StockValueAtEnd + Stocks(j, 5)
                Else
                    If IsEmpty(Stocks(j, 7)) Then
                        StockValueAtEnd = StockValueAtEnd + Stocks(j, 2)
                    Else
                        StockValueAtEnd = StockValueAtEnd + (GetPrice(CStr(Stocks(j, 7)), ThisDate) * Stocks(j, 6))
                    End If
                End If
                StockValueAtEnd = StockValueAtEnd + (Stocks(j, 8) * (ThisDate - PreviousDate)) ' Dividend
            End If
        Next j
        IntervalGains(i) = StockValueAtEnd / StockValueAtStart
        PreviousDate = ThisDate
    Next i
    
    'Now calculate the overall annualised growth rate
    Dim Gain As Double
    Gain = 1
    For i = 2 To UniqueDatesCount
        Gain = Gain * IntervalGains(i)
    Next i
    PorfolioGrowth = (Gain ^ (365.25 / (UniqueDates(UniqueDatesCount) - UniqueDates(1)))) - 1

End Function
