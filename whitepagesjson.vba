

Sub Test_DownloadTextFile()
    Dim text As String
    Dim sJSONString As String
    Dim vJSON As Variant
    Dim sState As String
    text = DownloadTextFile("https://www.whitepages.com.au/api/r/search?location=4209&name=hello")
    
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(text)
    'Debug.Print JsonConverter.ConvertToJson(Json)
    
    Dim Value As Dictionary
    Dim i As Long
    
    i = 0
    For Each Value In Json("results")
        Debug.Print Value("id")
      i = i + 1
    Next Value
    
End Sub

'Tool.References... Add a reference to Microsoft WinHTTPServices
Public Function DownloadTextFile(url As String) As String
    Dim oHTTP As WinHttp.WinHttpRequest
    Set oHTTP = New WinHttp.WinHttpRequest
    oHTTP.Open method:="GET", url:=url, async:=False
    oHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    oHTTP.setRequestHeader "Content-Type", "multipart/form-data; "
    oHTTP.Option(WinHttpRequestOption_EnableRedirects) = True
    oHTTP.send

    Dim success As Boolean
    success = oHTTP.waitForResponse()
    If Not success Then
        Debug.Print "DOWNLOAD FAILED!"
        Exit Function
    End If

    Dim responseText As String
    responseText = oHTTP.responseText

    Set oHTTP = Nothing

    DownloadTextFile = responseText
End Function










