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
