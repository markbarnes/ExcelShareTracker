Sub initDataCache()
    
    If DataCache Is Nothing Then
        Set DataCache = New Dictionary
    End If
    
    If dbConn Is Nothing Then
        Set dbConn = CreateObject("ADODB.Connection")
        dbConn.Open ("excel_investment_cache")
    End If

End Sub
