Sub LoopTrhgouh()

For Each c In Worksheets("emailedncounts").Range("A2")
 'MsgBox c.Value
 'MsgBox Worksheets(c.Value).Range("D3")
 Call Mail_Range("lol", "lol")
    
Next c

End Sub


Function Mail_Range(Email, Company)
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

    Set Source = Nothing
    On Error Resume Next
    Set Source = Range("A1:K50").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Source Is Nothing Then
        MsgBox "The source is not a range or the sheet is protected, please correct and try again.", vbOKOnly
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
            .To = "the-shadow-of-light@hotmail.com"
            .CC = ""
            .BCC = ""
            .Subject = "This is the Subject line"
            .Body = "Hi there"
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

