Dim row As Long
Dim i As Long

Dim IE As Object
Dim objElement As Object
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim objCollection As Object
Dim Doc As HTMLDocument

Sub Automate_IE_Load_Page()

    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True



    row = 2
    Dim rng As Range, cell As Range
    Dim URL As String
    Set rng = Range("b2:b468")
    For Each cell In rng
    URL = "https://fol.flick.com.au/location/edit.asp?Mode=Detail&LocationID=" + CStr(cell.Value)
    
    Call grabLinks(URL)
    Next cell
    
    'Navigate to URL
  ' Set IE = Nothing
  '  Set objElement = Nothing
   ' Set objCollection = Nothing

    
End Sub

Sub grabLinks(URL)
      On Error GoTo eh
      If IE Is Nothing Then
        Set IE = CreateObject("InternetExplorer.Application")
        IE.Visible = True
        Else

        End If
               Call runRequest(URL)
eh:
    Call PressEnterFirst
    
End Sub

Sub PressEnterFirst()
    SendKeys "~", True
End Sub

Sub runRequest(URL)
If IE Is Nothing Then
        Set IE = CreateObject("InternetExplorer.Application")
        IE.Visible = True
        Else

        End If
IE.navigate URL
               ' Statusbar let's user know website is loading
              ' Application.StatusBar = URL &amp; " is loading. Please wait..."
            
               ' Wait while IE loading...
               'IE ReadyState = 4 signifies the webpage has loaded (the first loop is set to avoid inadvertently skipping over the second loop)
               Do While IE.readyState = 4: DoEvents: Loop   'Do While
               Do Until IE.readyState = 4: DoEvents: Loop   'Do Until
               
               
               Dim myPoints As String
               Dim LArray() As String
               
               Dim fn As String
               fn = ""
               
               Dim ln As String
               ln = ""
               
               Dim p As String
               p = ""
               
               
               Dim bodyHtml As String
               
               Set Doc = IE.document
               bodyHtml = Doc.body.innerHTML
               If InStr(msg, "No Matches") = 0 Then
               
               
                    'extractedLoc = IE.getElementsByTagName("tr")(1).getElementsByTagName("td")(2).innerText
                    'IE.getElementById("FlickCopiedLocationLink").Value
                    If Not Doc.getElementsByName("fn") Is Nothing Then
                        
                        fn = Trim(Doc.getElementsByName("FName")(0).Value)
                        Cells(row, 3).Value = fn
                        Doc.getElementsByName("FName")(0).Value = Cells(row, 6)
                        
                           ln = Trim(Doc.getElementsByName("LName")(0).Value)
                        Cells(row, 4).Value = ln
                        Doc.getElementsByName("LName")(0).Value = Cells(row, 7)
                        
                        p = Trim(Doc.getElementsByName("Phone")(0).Value)
                        Cells(row, 5).Value = p
                        Doc.getElementsByName("Phone")(0).Value = Cells(row, 8)
                        
                        SendKeys "%s"
                        Sleep 2000
                        SendKeys "%s"
                        Sleep 2000







                    Else
                        
                    End If
                Else
                
                
                
                
                    Cells(row, 3).Value = "nothing there"
                End If
               row = row + 1
End Sub






