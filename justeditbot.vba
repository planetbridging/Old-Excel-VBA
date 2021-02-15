Dim rowCount As Long
Dim i As Long

Dim IE As Object
Dim objElement As Object
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal wClassName As Any, ByVal wWindowName As String) As Long


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long

Public Const BM_CLICK = &HF5&


Dim objCollection As Object

Dim locDoc As HTMLDocument
Dim locIE As Object

Dim locViewer As String
Dim locEdit As String




Sub setupBot()
    locViewer = "https://fol.flick.com.au/location/detail.asp?LocationID="
    locEdit = "https://fol.flick.com.au/location/edit.asp?Mode=Detail&LocationID="
    Set locIE = CreateObject("InternetExplorer.Application")
    locIE.Visible = True
    loopLocations
End Sub

Function loopLocations()
    rowCount = 2
    Set Rng = Range("b2:b4")
    For Each cell In Rng
        Dim tmpURL As String
        Dim tmpEditURL As String
        Dim tmpPostCollection As Collection
        Dim TmpPreDate As String
        Dim foundService As Boolean
        foundService = False
        tmpURL = locViewer + CStr(cell.Value)
        tmpEditURL = locEdit + CStr(cell.Value)
        'tmpPreURL = locViewer + CStr(Cells(rowCount, 4).Value)
        'tmpPreHisURL = locHistory + CStr(Cells(rowCount, 4).Value)
        'TmpPreDate = getPreconHistoryLocationInformation(tmpPreHisURL)
        Set tmpPostCollection = getLocationInformation(CStr(Cells(rowCount, 1).Value), tmpURL)
        'Set tmpPreCollection = getPreconLocationInformation(CStr(Cells(rowCount, 3).Value), tmpPreURL)
       
        For Each tmpServices In tmpPostCollection
            If StrComp(tmpServices, "FLIXPRECON-REIN ") = 0 Then
                foundService = True
            End If
        Next
        
        If tmpPostCollection(1) = "MATCH" Then
            getLocationEdit (tmpEditURL)
        End If
        
        If foundService = True Then
           ' MsgBox "Found Setup"
        Else
            'MsgBox "No Setup"
        End If
        Debug.Print rowCount
        rowCount = rowCount + 1
    Next cell
End Function

Function getLocationInformation(locCheck, locURL) As Collection
    'location number...billtonumber ...number in billto...fname,lastname,number,email...service setup...preconlink...
    Dim locationInformation As New Collection
    Dim locationServices As New Collection
    
    If locIE Is Nothing Then
        Set locIE = CreateObject("InternetExplorer.Application")
        locIE.Visible = True
    End If
    
    locIE.navigate locURL
    Do While locIE.readyState = 4: DoEvents: Loop   'Do While
    Do Until locIE.readyState = 4: DoEvents: Loop
    Set locDoc = locIE.document
    
    locDoc.getElementsByTagName("Table")(0).Click
    
   While locIE.Busy Or locIE.readyState <> READYSTATE_COMPLETE
            DoEvents
            If locIE.Busy Then
                DoEvents
                hWND = FindWindow(vbNullString, "POPPP!!")
                If hWND <> 0 Then childHWND = FindWindowEx(hWND, ByVal 0&, "Button", "OK")
                If childHWND <> 0 Then SendMessage childHWND, BM_CLICK, 0, 0
                SendKeys "~"
            End If
    Wend
    
    If StrComp(CStr(locDoc.getElementById("LocationNumber").outerText), locCheck) = 0 Then
        Dim tr As MSHTML.IHTMLElementCollection
        Dim td As MSHTML.IHTMLElementCollection
        Dim rowNumber As Long
        rowNumber = 1
        Set tr = locDoc.getElementById("ProgramsTable").getElementsByTagName("tr")
        locationServices.Add "MATCH"
        For Each trObj In tr
            Set td = trObj.getElementsByTagName("td")
            If rowNumber > 1 And rowNumber < tr.Length Then
                Dim tdCount As Long
                tdCount = 1
                For Each tdObj In td
                    If tdCount = 3 Then
                        locationServices.Add (CStr(tdObj.outerText))
                    End If
                    tdCount = tdCount + 1
                Next
            End If
            rowNumber = rowNumber + 1
        Next
    Else
        locationServices.Add "NOTMATCH"
    End If
    Set getLocationInformation = locationServices
End Function


Function getPreconLocationInformation(locCheck, locURL) As Collection
    Dim locationServices As New Collection
    
    If locPreIE Is Nothing Then
        Set locPreIE = CreateObject("InternetExplorer.Application")
        locPreIE.Visible = True
    End If
    
    locPreIE.navigate locURL
    Do While locPreIE.readyState = 4: DoEvents: Loop   'Do While
    Do Until locPreIE.readyState = 4: DoEvents: Loop
    Set locPreDoc = locPreIE.document
    
    locPreDoc.getElementsByTagName("Table")(0).Click
    
   While locPreIE.Busy Or locPreIE.readyState <> READYSTATE_COMPLETE
            DoEvents
            If locPreIE.Busy Then
                DoEvents
                hWND = FindWindow(vbNullString, "POPPP!!")
                If hWND <> 0 Then childHWND = FindWindowEx(hWND, ByVal 0&, "Button", "OK")
                If childHWND <> 0 Then SendMessage childHWND, BM_CLICK, 0, 0
                SendKeys "~"
            End If
    Wend
    
    If StrComp(CStr(locPreDoc.getElementById("LocationNumber").outerText), locCheck) = 0 Then
        locationServices.Add "MATCH"
    Else
        locationServices.Add "NOTMATCH"
    End If
    Set getPreconLocationInformation = locationServices
End Function


Function getPreconHistoryLocationInformation(locURL) As String
    Dim locationServices As New Collection
    Dim returnDate As String
    returnDate = ""
    If locPreHisIE Is Nothing Then
        Set locPreHisIE = CreateObject("InternetExplorer.Application")
        locPreHisIE.Visible = True
    End If
    locURL = locURL + "&Sort=WorkDate"
    locPreHisIE.navigate locURL
    Do While locPreHisIE.readyState = 4: DoEvents: Loop   'Do While
    Do Until locPreHisIE.readyState = 4: DoEvents: Loop
    Set locPreHisDoc = locPreHisIE.document
    
    Dim tr As MSHTML.IHTMLElementCollection
    Dim td As MSHTML.IHTMLElementCollection
    Dim rowNumber As Long
    rowNumber = 1
    Set tr = locPreHisDoc.getElementsByTagName("tr")
    For Each trObj In tr
        Set td = trObj.getElementsByTagName("td")
        If rowNumber = 2 And rowNumber < tr.Length Then
                Dim tdCount As Long
                tdCount = 1
                For Each tdObj In td
                    If tdCount = 5 Then
                        returnDate = CStr(tdObj.outerText)
                    End If
                    tdCount = tdCount + 1
                Next
            End If
            rowNumber = rowNumber + 1
        Next
    getPreconHistoryLocationInformation = returnDate
End Function

'421194643   2062929 Brian   Stobby  0402-024-343

Function getLocationEdit(locURL)
    'location number...billtonumber ...number in billto...fname,lastname,number,email...service setup...preconlink...
    Dim locationInformation As New Collection
    Dim locationServices As New Collection
    
    If locIE Is Nothing Then
        Set locIE = CreateObject("InternetExplorer.Application")
        locIE.Visible = True
    End If
    
    locIE.navigate locURL
    Do While locIE.readyState = 4: DoEvents: Loop   'Do While
    Do Until locIE.readyState = 4: DoEvents: Loop
    Set locDoc = locIE.document
    'Call locIE.document.parentWindow.execScript("IncludedPestSpan_OnClick()", "JavaScript")
    
    Dim fn As String
    Dim ln As String
    Dim p As String
    
    fn = Trim(locDoc.getElementsByName("FName")(0).Value)
    Cells(rowCount, 6).Value = fn
    locDoc.getElementsByName("FName")(0).Value = Cells(rowCount, 3)
                        
    ln = Trim(locDoc.getElementsByName("LName")(0).Value)
    Cells(rowCount, 7).Value = ln
    locDoc.getElementsByName("LName")(0).Value = Cells(rowCount, 4)
                        
    p = Trim(locDoc.getElementsByName("Phone")(0).Value)
    Cells(rowCount, 8).Value = p
    locDoc.getElementsByName("Phone")(0).Value = Cells(rowCount, 5)
    
    
    SendKeys "%s"
    
    
    Do While locIE.readyState = 4: DoEvents: Loop   'Do While
    Do Until locIE.readyState = 4: DoEvents: Loop
    
    While locIE.Busy Or locIE.readyState <> READYSTATE_COMPLETE
            DoEvents
            If locIE.Busy Then
                DoEvents
                hWND = FindWindow(vbNullString, "POPPP!!")
                If hWND <> 0 Then childHWND = FindWindowEx(hWND, ByVal 0&, "Button", "OK")
                If childHWND <> 0 Then SendMessage childHWND, BM_CLICK, 0, 0
                SendKeys "~"
            End If
    Wend
    
End Function




