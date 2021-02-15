Dim row As Long
Dim i As Long
Dim URL As String
Dim IE As Object
Dim objElement As Object
Dim objCollection As Object

Sub Automate_IE_Load_Page()
'This will load a webpage in IE
    
 
    'Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
 
    'Set IE.Visible = True to make IE visible, or False for IE to run in the background
    IE.Visible = True
 
    'Define URL
    'URL = "https://fol.flick.com.au/location/detail.asp?LocationID=1940463"
    'URL = "https://fol.flick.com.au/location/list.asp?ListID=141940&ListSort=LocationCode&OutputFormat=&Page=1&Count="
    URL = ""
    row = 31
    Dim rng As Range, cell As Range
    Set rng = Range("b31:b29408")
    For Each cell In rng
    URL = "https://fol.flick.com.au/location/detail.asp?LocationID=" + CStr(cell.Value)
    Call grabLinks
    Next cell
    
    'Navigate to URL
   Set IE = Nothing
    Set objElement = Nothing
    Set objCollection = Nothing

    
End Sub

Sub grabLinks()
    IE.Navigate URL
 
    ' Statusbar let's user know website is loading
   ' Application.StatusBar = URL &amp; " is loading. Please wait..."
 
    ' Wait while IE loading...
    'IE ReadyState = 4 signifies the webpage has loaded (the first loop is set to avoid inadvertently skipping over the second loop)
    Do While IE.ReadyState = 4: DoEvents: Loop   'Do While
    Do Until IE.ReadyState = 4: DoEvents: Loop   'Do Until
    
    Dim Doc As HTMLDocument
    Dim myPoints As String
    Dim LArray() As String
    Dim extractedLoc As String
    extractedLoc = ""
    Set Doc = IE.document
    
    'extractedLoc = IE.getElementsByTagName("tr")(1).getElementsByTagName("td")(2).innerText
    'IE.getElementById("FlickCopiedLocationLink").Value
    If Not Doc.getElementById("FlickCopiedLocationLink") Is Nothing Then
        myPoints = Trim(Doc.getElementById("FlickCopiedLocationLink").outerText)
        LArray = Split(myPoints, " ")
        extractedLoc = LArray(2)
        extractedLoc = Replace(extractedLoc, "#", "")
        Cells(row, 3).Value = extractedLoc
    Else
        
    End If
    row = row + 1
   ' Debug.Print extractedLoc
    'Webpage Loaded
   'Application.StatusBar = URL &amp; " Loaded"
    
    'Unload IE
    
End Sub




