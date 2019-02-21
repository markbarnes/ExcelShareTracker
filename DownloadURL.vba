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
