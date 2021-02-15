Sub arrayBuilder()

myarray = Worksheets("email").Range("A2:B472")

'unlike most VBA Arrays, this array doesn't need to be declared and will be automatically dimensioned

For i = 1 To UBound(myarray)
    'Debug.Print "row" & i
    'Debug.Print myarray(i, 1)
    'Debug.Print myarray(i, 2)
    Call Mail_Range(myarray(i, 1), myarray(i, 2))
    'For j = 1 To UBound(myarray, 2)

    'Debug.Print (myarray(i, j))

   ' Next j

Next i

End Sub

Function Mail_Range(Company, Email)
    'Working in Excel 2000-2016
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    Dim Source As Range
    Dim Dest As Workbook
    Dim wb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim footer As String
    Dim header As String
    Dim request As String
    Dim closing As String
    footer = "<br><p>Kind regards,<br>Shannon Setter & Declan Kemp<br>Pre-con to Post-con Conversion Team <br> declan.kemp@flick-anticimex.com.au or shannon.setter@flick-anticimex.com.au <br> Office: 07 5512 0710 Direct: 07 5512 0725<br>flick anticimex<br>5/168 Siganto Drive, Helensvale, 4212<br>www.flick-anticimex.com.au</p>"
    header = "<p>Hello,</p>" & "<p>It has become a worrying trend that termite warranty forms are not being activated by home owners. This is a concern for both Flick - Anticimex and for " & Company & " as your 50 year warranty, from final installation. may be void, if we are not given the opportunity to inspect your clients homes annually.</p>" & "<p>Including home owner details on your purchase orders or work orders, will help to ensure that we are able to make contact with your clients.  As a valued industry partner, we want to ensure that our management systems are monitored for possible breaches, but we need your help to do this.</p>"
    request = "<p>Please provide the home owners  (First Name, Surname, Email Address and Phone Number) for the following property/s: </p>"
    closing = "<p>I look forward to a positive response and thank you for choosing Flick for all your termite and waterproofing requirements.</p>"
    Set Source = Nothing
    On Error Resume Next
    Set Source = Worksheets(Company).Range("A1:J1000")
    On Error GoTo 0

    If Source Is Nothing Then
        'MsgBox "The source is not a range or the sheet is protected, please correct and try again.", vbOKOnly
        Debug.Print "Failed: " & Company & " " & Email
        Exit Function
    End If

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set wb = ActiveWorkbook
    Set Dest = Workbooks.Add(xlWBATWorksheet)

    Source.Copy
    With Dest.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial Paste:=xlPasteValues
        .Cells(1).PasteSpecial Paste:=xlPasteFormats
        .Cells(1).Select
        Application.CutCopyMode = False
    End With

    TempFilePath = Environ$("temp") & "\"
    TempFileName = "Selection of " & wb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")

    If Val(Application.Version) < 12 Then
        'You use Excel 97-2003
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        'You use Excel 2007-2016
        FileExtStr = ".xlsx": FileFormatNum = 51
    End If

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Dest
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .to = Email
            .CC = "declan.kemp@flick-anticimex.com.au"
            .BCC = ""
            .Subject = Company & " Termite Warranty"
            .HTMLBody = header & request & RangetoHTML(Source) & closing & footer
            .Attachments.Add Dest.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            .Send   'or use .Display
        End With
        On Error GoTo 0
        .Close savechanges:=False
    End With

    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Function

Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function





