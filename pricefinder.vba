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
'https://www.pricefinder.com.au/
Sub PFStart()
    addedDetails = 2
    running = True
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    IE.navigate "https://www.pricefinder.com.au/"
    
    Do While IE.readyState = 4: DoEvents: Loop   'Do While
    Do Until IE.readyState = 4: DoEvents: Loop   'Do Until
        
    Set Doc = IE.document

    
    lngRow = Worksheets("loop").Cells(Rows.Count, "A").End(xlUp).row
    MsgBox "found: " + CStr(lngRow)

    row = 2
    Dim rng As Range, cell As Range
    Dim URL As String
    Set rng = Worksheets("loop").Range("a2:a" + CStr(lngRow))
    For Each cell In rng
        Worksheets("loop").Cells(row, 3).Value = "Opening"
        runRequest CStr(row), CStr(Worksheets("loop").Cells(row, 1).Value), CStr(Worksheets("loop").Cells(row, 2).Value)
        If running = False Then Exit For
        row = row + 1
        Debug.Print CStr(row) + "/" + CStr(lngRow)
    Next cell
   
End Sub

Sub PFStop()
running = False
End Sub

Function runRequest(NumberOfStreet, StreetName, Suburb)
    
    
    'streetNameInput
    'suburbPostcodeSuggest
    
    'getElementById

    
    Dim tr As MSHTML.IHTMLElementCollection
    Dim td As MSHTML.IHTMLElementCollection
    Dim findSurenameCompay As MSHTML.IHTMLElementCollection
    
    Dim BtnSearch As HTMLInputElement
    
    Set tr = Doc.getElementsByTagName("form")
    For Each trObj In tr
        
        Dim found As Boolean
        found = False
        Set findSurenameCompay = trObj.getElementsByTagName("div")
        
        For Each divObj In findSurenameCompay
            If InStr(divObj.outerText, "Surname / Company") > 0 Then
                found = True
                Exit For
            End If
        Next
    
        If found Then
            Dim inputCount As Integer
            inputCount = 0
            'formContents(0).Value = StreetName
            'trObj.contentDocument.getElementsByName("suburbPostcodeSuggest")(0).Value = Suburb
            Set td = trObj.getElementsByTagName("input")
            For Each tdObj In td
                If InStr(tdObj.Value, "SEARCH") > 0 Then
                    Set BtnSearch = tdObj
                ElseIf inputCount = 9 Then
                    tdObj.Value = StreetName
                ElseIf inputCount = 10 Then
                    tdObj.Value = Suburb
                End If
                'Debug.Print tdObj.Value
                inputCount = inputCount + 1
            Next
        End If
        
        
        'Debug.Print tr.outText
    Next
    
    BtnSearch.Click
    'Doc.getElementById("Form_0").Click
    
    Do While IE.readyState = 4: DoEvents: Loop   'Do While
    Do Until IE.readyState = 4: DoEvents: Loop   'Do Until
        
    Set Doc = IE.document
    
    Dim trTBL As MSHTML.IHTMLElementCollection
    Dim tdTBL As MSHTML.IHTMLElementCollection
    
    If Not Doc.getElementById("resultsTable") Is Nothing Then
        
        
        Set trTBL = Doc.getElementsByTagName("tr")
        For Each trTBLObj In trTBL
            Dim rowCount As Integer
            rowCount = 2
            Set tdTBL = trTBLObj.getElementsByTagName("td")
            Worksheets("dump").Cells(addedDetails, 1).Value = CStr(NumberOfStreet)
            For Each tdTBLObj In tdTBL
                Worksheets("dump").Cells(addedDetails, rowCount).Value = CStr(tdTBLObj.outerText)
                rowCount = rowCount + 1
            Next
            addedDetails = addedDetails + 1
        Next
        Worksheets("loop").Cells(NumberOfStreet, 3).Value = "Done"
    Else
        Worksheets("loop").Cells(NumberOfStreet, 3).Value = "not found"
    End If


End Function


