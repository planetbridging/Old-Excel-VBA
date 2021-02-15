Dim footer As String
Dim bodyEmail As String
Dim bodyEmail2 As String
Dim bodyEmail3 As String
Dim header As String
Dim row As Long


Sub LoopTrhgouh()
footer = "<p>Kind regards,</p><p>Shannon Setter & Declan Kemp<br>Pre-con to Post-con Conversion Team <br> Declan.Kemp@flick-anticimex.com.au or Shannon.Setter@flick-anticimex.com.au <br> Office: 07 5512 0710 Direct: 07 5512 0725<br>Flick-Anticimex<br>5/168 Siganto Drive, Helensvale, 4212<br>www.flick-anticimex.com.au</p>"
bodyEmail = "<p>Firstly, congratulations on your new home.</p><p>As part of the construction process, Flick-Anticimex have installed the termite protection to your property at "
bodyEmail2 = ". As a part of this procedure, a warranty activation form will need to be completed by you and returned to our office.</p><p>Currently your warranty is inactive. As part of your hand over package, there should have been a warranty activation document given to you at the time of hand over for your property.</p>"
bodyEmail2 = bodyEmail2 + "<p>Please find this form attached to fill out at your earliest convenience and forward through to myself, at this email address.</p><p>Once it has been received, I will activate your warranty and make sure everything is up to date on our system at "
    row = 2
    Dim rng As Range, cell As Range
    Set rng = Range("i2:i94")
    For Each cell In rng
    Call Mail_Range(cell)
    Next cell

End Sub


Function Mail_Range(Email)
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
    Dim buildAddress As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    'Cells(row, 3).Value
    
    
    buildAddress = CStr(Cells(row, 2).Value) + " " + CStr(Cells(row, 3).Value) + " " + CStr(Cells(row, 4).Value) + " " + CStr(Cells(row, 5).Value) + " " + CStr(Cells(row, 6).Value)
    bodyEmail3 = CStr(Cells(row, 1).Value) + " customer number.</p><p>Thank you for your understanding, and I look forward to receiving a positive response in the coming days.</p>"

        With OutMail
            .To = Email
            .CC = "declan.kemp@flick-anticimex.com.au"
            .BCC = ""
            .Subject = "Termite Warranty"
            .HTMLBody = "<p>Dear " + Cells(row, 7).Value + ",</p>" + bodyEmail + buildAddress + bodyEmail2 + bodyEmail3 + footer
            '.Attachments.Add Dest.FullName
            'You can add other files also like this
            .Attachments.Add ("C:\allyesno\Form.pdf")
            .Send   'or use .Display
        End With
    Cells(row, 8).Value = "TRIED TO SEND"
    Debug.Print Email + " Email sent"
    row = row + 1
End Function


