

Sub Test_DownloadTextFile()
    Dim text As String
    Dim sJSONString As String
    Dim vJSON As Variant
    Dim sState As String
    text = DownloadTextFile("https://www.whitepages.com.au/api/r/search?location=4209&name=hello")
    
    Dim Json As Object
    Dim Item As Variant
    Set Json = ParseJson(text)
    'Debug.Print JsonConverter.ConvertToJson(Json)
    
    Dim Parsed As Dictionary
    Set Parsed = JsonConverter.ParseJson(text)
    Dim i As Long

    i = 1
    
    Dim Key As Variant
    
    
    i = 2
    For Each Item In Json("results")(1)
        Dim Value As Dictionary
        For Each Value In Parsed("name")
          Value ("surname")

          i = i + 1
        Next Value
        'Debug.Print Parsed(Item)(1).("surname")
    
    Next
'          ^ Need Variant/Object for For Each

    'For Each Key In Parsed("results")(1).keys
    '    Debug.Print Parsed("results")(1).Key("id")
   ' Next Key

   '' For Each Value In Parsed("results")
    '  Debug.Print Value("physicalAddress")(1)("streetNumber")
  '  Next Value
    'For Each Value In Json("results")
     '   Debug.Print Value("id")(1).keys
      '  For Each Key In Json("results")(i).keys
           
       ' Next Key
        'i = i + 1
    'Next Value
    
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










