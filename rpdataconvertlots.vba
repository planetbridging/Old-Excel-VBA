Dim row As Long
Dim i As Long

Dim IE As Object
Dim objElement As Object
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim objCollection As Object
Dim Doc As HTMLDocument

Dim addedDetails As Long

'3343/8766 done
Sub Automate_IE_Load_Page()

    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True

    addedDetails = 2

    row = 2
    
    '6082
    
    runRequest ("https://rpp.rpdata.com/rpp/search/address/property/summary.html?q=Bluetail+Crescent+Upper+Coomera+QLD+4209&qt=address&view=property&newSearch=true&searchWindowId=")
    
   
End Sub


Function runRequest(URL)
If IE Is Nothing Then
        Set IE = CreateObject("InternetExplorer.Application")
        IE.Visible = True
        Else

        End If
        IE.navigate URL
        

        Do While IE.readyState = 4: DoEvents: Loop   'Do While
        Do Until IE.readyState = 4: DoEvents: Loop   'Do Until
        
        Set Doc = IE.document
        

        If Not Doc.getElementById("username")(0) Is Nothing Then

                Doc.getElementsByTagName("form")(0).getElementsByTagName("button")(0).Click
                
                'Doc.getElementById("searchAddressSimple").getElementById("addressLink").getElementByTagName("span").Click
                'IE.navigate URLaddressLink
                Do While IE.readyState = 4: DoEvents: Loop   'Do While
                Do Until IE.readyState = 4: DoEvents: Loop   'Do Until
        End If
        
        Doc.getElementById("searchAddressSimple").getElementsByTagName("input")(0).Value = "Arcadia Boulevard Pimpama QLD 4209"
        Doc.getElementById("searchAddressSimple").getElementsByTagName("a")(0).Click
        
        
        
      
        'Doc.getElementById("propertySummaryList").getElementsByTagName("div")(0).getElementsByTagName("div")(8).getElementsByTagName("li")(4).Click
        
        '(4).Click
        
        Do While IE.readyState = 4: DoEvents: Loop   'Do While
        Do Until IE.readyState = 4: DoEvents: Loop
        
        Dim button As Boolean
        button = True
        
        Dim detect As String
        detect = CStr(Doc.getElementsByClassName("summaryListItem ")(0).innerText)
        
        
        Dim tr As MSHTML.IHTMLElementCollection
        Set tr = Doc.getElementById("propertySummaryList").getElementsByTagName("div")(0).getElementsByTagName("div")(7).getElementsByTagName("li")
        For Each trObj In tr
            Debug.Print trObj.outerText
            If InStr(CStr(trObj.outerText), "Next") > 0 Then
                trObj.Click
            End If
        Next

        Do While IE.readyState = 4: DoEvents: Loop   'Do While
        Do Until IE.readyState = 4: DoEvents: Loop
        
        
        
        
       ' If Not Doc.getElementsByTagName("operationPanel") Is Nothing Then
     '       Dim tr As MSHTML.IHTMLElementCollection
       '     Set tr = Doc.getElementsByTagName("operationPanel").getElementsByTagName("li")
       '     For Each trObj In tr
       '         Debug.Print trObj.outerText
      '      Next
      '  End If
        
       ' btn-primary btn-block
        
        row = row + 1
End Function











