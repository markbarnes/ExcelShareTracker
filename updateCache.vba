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
