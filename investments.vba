Option Explicit
Global DataCache As Dictionary
Global dbConn As ADODB.Connection


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
Function CompetitorGrowth(StartDate As Date, EndDate As Date) As Double
    
    Dim StartCompetitorPrice, EndCompetitorPrice As Double
    
    If EndDate = 0 Then
        EndDate = Date
    End If
    
    StartCompetitorPrice = GetPrice("GB00B7LWFW05", StartDate)
    EndCompetitorPrice = GetPrice("GB00B7LWFW05", EndDate)
    
    CompetitorGrowth = (EndCompetitorPrice - StartCompetitorPrice) / StartCompetitorPrice
    
End Function
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
Function AnnualGrowth(StartValue As Double, EndValue As Double, StartDate As Date, EndDate As Date) As Double
    If EndDate = 0 Then
        EndDate = Date
    End If
    AnnualGrowth = (EndValue / StartValue) ^ (1 / ((EndDate - StartDate) / 365.25)) - 1
End Function

Function GetPrice(ISIN As String, Optional PriceDate As Date)
    Dim URL As String, response As String, MorningStarID As String, PriceCacheKey As String, price As Double, timestamp As Double, StartDate As String, EndDate As String
    
    Call initDataCache
    
    ISIN = Trim(ISIN)
    PriceCacheKey = ISIN & Format(PriceDate, "yyyymmdd")
    EndDate = Format(PriceDate, "yyyy-mm-dd")
    StartDate = Format(DateAdd("d", -7, PriceDate), "yyyy-mm-dd")
    
    If (Not (cacheKeyExists(PriceCacheKey))) Then
        MorningStarID = getMorningstarID(ISIN)
        response = DownloadURL("http://tools.morningstar.co.uk/api/rest.svc/timeseries_price/t92wz0sj7c?currencyId=GBP&idtype=Morningstar&frequency=daily&startDate=" & StartDate & "&endDate=" & EndDate & "&performanceType=&outputType=COMPACTJSON&id=" & MorningStarID & "]8]0]CAALL$$ALL&decPlaces=8&applyTrackRecordExtension=true")
        response = Right(response, Len(response) - InStrRev(response, ",") - 1)
        response = Left(response, Len(response) - 2)
        Call updateCache(PriceCacheKey, Val(Trim(response)))
    End If
    GetPrice = Val(getFromCache(PriceCacheKey))
End Function

Function getMorningstarID(ISIN As String)

    Dim timestamp As String, response As String, MSIDCacheKey As String

    Call initDataCache
    
    MSIDCacheKey = ISIN & "msid"
    If (Not (cacheKeyExists(MSIDCacheKey))) Then
        timestamp = DateDiff("s", "01/01/1970", Now()) * 1000
        response = DownloadURL("http://www.morningstar.co.uk/uk/util/SecuritySearch.ashx?source=nav&moduleId=6&ifIncludeAds=True&q=" & LCase(ISIN) & "&limit=100&timestamp=" & timestamp)
        Call updateCache(MSIDCacheKey, ReturnBetween(response, "|{""i"":""", ""","""))
    End If
    getMorningstarID = getFromCache(MSIDCacheKey)
End Function
Sub updateCache(CacheKey As String, CacheValue As Variant)
    Dim DataCacheRecordSet As ADODB.Recordset, TodaysDate As String

    TodaysDate = Format(Date, "yyyymmdd")
    If Right(CacheKey, Len(TodaysDate)) = TodaysDate Then
        'do nothing
    Else
        Set DataCacheRecordSet = dbConn.Execute("INSERT INTO cache (cache_key, cache_value) VALUES (""" & CacheKey & """, """ & CacheValue & """) ON DUPLICATE KEY UPDATE cache_value=""" & CacheValue & """")
    End If
    If DataCache.Exists(CacheKey) Then
        DataCache.Item(CacheKey) = CacheValue
    Else
        DataCache.Add CacheKey, CacheValue
    End If
End Sub

Function getFromCache(CacheKey As Variant)
    Dim DataCacheRecordSet As ADODB.Recordset, CacheValue As Variant
        
    If Not (DataCache.Exists(CacheKey)) Then
        Set DataCacheRecordSet = dbConn.Execute("SELECT cache_value FROM cache WHERE cache_key LIKE """ & CacheKey & """ LIMIT 1")
        If Not (DataCacheRecordSet.EOF) Then
            CacheValue = DataCacheRecordSet("cache_value")
            If Not (IsEmpty(CacheValue)) Then
                DataCache.Add CacheKey, CacheValue
            End If
        End If
    End If
    getFromCache = DataCache.Item(CacheKey)
 
End Function
Function cacheKeyExists(CacheKey As Variant) As Boolean

    Dim DataCacheRecordSet As ADODB.Recordset, CacheValue As Variant

    If Not (DataCache.Exists(CacheKey)) Then
        Set DataCacheRecordSet = dbConn.Execute("SELECT cache_value FROM cache WHERE cache_key LIKE """ & CacheKey & """ LIMIT 1")
        If Not (DataCacheRecordSet.EOF) Then
            CacheValue = DataCacheRecordSet("cache_value")
            If Not (IsEmpty(CacheValue)) Then
                DataCache.Add CacheKey, CacheValue
            End If
        End If
    End If
    cacheKeyExists = DataCache.Exists(CacheKey)
End Function

Sub initDataCache()
    
    If DataCache Is Nothing Then
        Set DataCache = New Dictionary
    End If
    
    If dbConn Is Nothing Then
        Set dbConn = CreateObject("ADODB.Connection")
        dbConn.Open ("excel_investment_cache")
    End If

End Sub

Function DownloadURL(URL)
    Dim request As WinHttp.WinHttpRequest
    Set request = New WinHttp.WinHttpRequest
    With request
        .Open "GET", URL, True
        .Send
        .WaitForResponse
    End With
    DownloadURL = request.responseText
End Function

Function ReturnAfter(haystack, needle)
    ReturnAfter = Right(haystack, Len(haystack) - InStr(haystack, needle) - Len(needle) + 1)
End Function

Function ReturnBefore(haystack, needle)
    ReturnBefore = Left(haystack, InStr(haystack, needle) - 1)
End Function

Function ReturnBetween(haystack, StartString, EndString)
    ReturnBetween = ReturnBefore(ReturnAfter(haystack, StartString), EndString)
End Function



Sub QuickSort(arr, Lo As Long, Hi As Long)
  Dim varPivot As Variant
  Dim varTmp As Variant
  Dim tmpLow As Long
  Dim tmpHi As Long
  tmpLow = Lo
  tmpHi = Hi
  varPivot = arr((Lo + Hi) \ 2)
  Do While tmpLow <= tmpHi
    Do While arr(tmpLow) < varPivot And tmpLow < Hi
      tmpLow = tmpLow + 1
    Loop
    Do While varPivot < arr(tmpHi) And tmpHi > Lo
      tmpHi = tmpHi - 1
    Loop
    If tmpLow <= tmpHi Then
      varTmp = arr(tmpLow)
      arr(tmpLow) = arr(tmpHi)
      arr(tmpHi) = varTmp
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
  Loop
  If Lo < tmpHi Then QuickSort arr, Lo, tmpHi
  If tmpLow < Hi Then QuickSort arr, tmpLow, Hi
End Sub
