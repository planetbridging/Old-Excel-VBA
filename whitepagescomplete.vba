Dim row As Long
Dim i As Long

Dim IE As Object
Dim objElement As Object
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim objCollection As Object
Dim Doc As HTMLDocument

Dim addedDetails As Long
Dim lngRow As Long
Dim running As Boolean

'3343/8766 done
Sub WhitePagesStart()
    running = True
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    
    Dim postcode As Variant
    postcode = InputBox("Postcode")
    
    lngRow = Worksheets("loop").Cells(Rows.Count, "A").End(xlUp).row
    MsgBox "found: " + CStr(lngRow)
    addedDetails = 2

    row = 2
    Dim rng As Range, cell As Range
    Dim URL As String
    Set rng = Worksheets("loop").Range("a2:a" + CStr(lngRow))
    For Each cell In rng
        URL = "https://www.whitepages.com.au/residential/results?name=" + CStr(cell.Value) + "&location=" + CStr(postcode)
        runRequest (URL)
        If running = False Then Exit For
    Next cell
   
End Sub

Sub WhitePagesStop()
running = False
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
        
        
        Dim button As Boolean
        button = True
        
        Dim capture As Boolean
        capture = True
        
        While capture
            Dim notRobot As MSHTML.IHTMLElementCollection
            Set notRobot = Doc.getElementsByTagName("wp-captcha")
            Dim i As Long
            Dim EmptyCounter As Long
            EmptyCounter = 0
            i = 0
            For Each trObj In notRobot
                    EmptyCounter = EmptyCounter + 1
            Next
            If EmptyCounter = 0 Then
                capture = False
            Else
                Beep
                MsgBox "Capture"
            End If
        Wend
        
        
        
        
        
        Dim divs As Object
        Dim div As Object
        
        Set divs = Doc.getElementsByTagName("button")
        
        While button
            Dim btnCount As Integer
            btnCount = 0
            For Each div In divs
                Debug.Print div.outerText
                If InStr(div.outerText, "Show") > 0 Then
                  div.Click
                  btnCount = btnCount + 1
                End If
            Next div
  
            If btnCount = 0 Then
                button = False
            End If
            
        Wend
                
        Sleep 2000
        Dim tr As MSHTML.IHTMLElementCollection
        Dim td As MSHTML.IHTMLElementCollection
        Dim rowNumber As Long

        rowNumber = 1
         Dim links As Variant
        Set tr = Doc.getElementsByTagName("li")
        For Each trObj In tr
            If Not trObj.getElementsByTagName("div") Is Nothing Then
                For Each d In trObj.getElementsByTagName("wp-residential-search-result")
                   For Each nameNStreet In trObj.getElementsByTagName("a")
                        Dim Split As Integer
                        Split = 0
                        
                        Dim n As String
                        Dim a1 As String
                        Dim a2 As String
                        Dim a3 As String
                        Dim p As String
                        
                        For Each divs In nameNStreet.getElementsByTagName("div")
                            If Split = 0 Then
                                If Len(divs.outerText) > 0 Then
                                   ' Debug.Print "Names" + divs.outerText + "Names"
                                    'Worksheets("spit").Cells(addedDetails, 1).Value = CStr(divs.outerText)
                                    n = CStr(divs.outerText)
                                End If
                            End If
                            
                            If Split = 2 Then
                                'Debug.Print "Street" + divs.outerText + "Street"
                               ' Worksheets("spit").Cells(addedDetails, 2).Value = CStr(divs.outerText)
                               a1 = CStr(divs.outerText)
                            End If
                            
                            If Split = 3 Then
                                a2 = CStr(divs.outerText)
                                'Debug.Print "sub" + divs.outerText + "sub"
                               ' Worksheets("spit").Cells(addedDetails, 3).Value = CStr(divs.outerText)
                            End If
                            
                            If Split = 4 Then
                                a3 = CStr(divs.outerText)
                               ' Debug.Print "sub" + divs.outerText + "sub"
                                'Worksheets("spit").Cells(addedDetails, 3).Value = CStr(divs.outerText)
                            End If
                            
                             If Split = 5 Then
                                 p = CStr(divs.outerText)
                                'Debug.Print "idk" + divs.outerText + "idk"
                                'Worksheets("spit").Cells(addedDetails, 3).Value = CStr(divs.outerText)
                                'addedDetails = addedDetails + 1
                            End If
                        
                            Split = Split + 1
                        Next
                        Worksheets("results").Cells(addedDetails, 1).Value = CStr(n)
                        Worksheets("results").Cells(addedDetails, 2).Value = CStr(a1)
                        Worksheets("results").Cells(addedDetails, 3).Value = CStr(a2)
                        Worksheets("results").Cells(addedDetails, 4).Value = CStr(a3)
                        Worksheets("results").Cells(addedDetails, 5).Value = CStr(p)
                        addedDetails = addedDetails + 1
                   Next
                Next
            End If
        Next

        Debug.Print CStr(row) + "/" + CStr(lngRow) + " done"
        row = row + 1
End Function












