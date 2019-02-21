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
