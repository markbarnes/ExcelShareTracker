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
