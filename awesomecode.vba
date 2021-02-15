Sub Copy_To_Worksheets()
'Note: This macro use the function LastRow
    Dim My_Range As Range
    Dim FieldNum As Long
    Dim CalcMode As Long
    Dim ViewMode As Long
    Dim ws2 As Worksheet
    Dim Lrow As Long
    Dim cell As Range
    Dim CCount As Long
    Dim WSNew As Worksheet
    Dim ErrNum As Long

    'Set filter range on ActiveSheet: A1 is the top left cell of your filter range
    'and the header of the first column, D is the last column in the filter range.
    'You can also add the sheet name to the code like this :
    'Worksheets("Sheet1").Range("A1:D" & LastRow(Worksheets("Sheet1")))
    'No need that the sheet is active then when you run the macro when you use this.
    Set My_Range = Range("A1:J" & LastRow(ActiveSheet))
    My_Range.Parent.Select

    If ActiveWorkbook.ProtectStructure = True Or _
       My_Range.Parent.ProtectContents = True Then
        MsgBox "Sorry, not working when the workbook or worksheet is protected", _
               vbOKOnly, "Copy to new worksheet"
        Exit Sub
    End If

    'This example filters on the first column in the range(change the field if needed)
    'In this case the range starts in A so Field:=1 is column A, 2 = column B, ......
    FieldNum = 1

    'Turn off AutoFilter
    My_Range.Parent.AutoFilterMode = False

    'Change ScreenUpdating, Calculation, EnableEvents, ....
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    ViewMode = ActiveWindow.View
    ActiveWindow.View = xlNormalView
    ActiveSheet.DisplayPageBreaks = False

    'Add a worksheet to copy the a unique list and add the CriteriaRange
    Set ws2 = Worksheets.Add

    With ws2
        'first we copy the Unique data from the filter field to ws2
        My_Range.Columns(FieldNum).AdvancedFilter _
                Action:=xlFilterCopy, _
                CopyToRange:=.Range("A1"), Unique:=True

        'loop through the unique list in ws2 and filter/copy to a new sheet
        Lrow = .Cells(Rows.Count, "A").End(xlUp).Row
        For Each cell In .Range("A2:A" & Lrow)

            'Filter the range
            My_Range.AutoFilter Field:=FieldNum, Criteria1:="=" & _
             Replace(Replace(Replace(cell.Value, "~", "~~"), "*", "~*"), "?", "~?")

            'Check if there are no more then 8192 areas(limit of areas)
            CCount = 0
            On Error Resume Next
            CCount = My_Range.Columns(1).SpecialCells(xlCellTypeVisible) _
                     .Areas(1).Cells.Count
            On Error GoTo 0
            If CCount = 0 Then
                MsgBox "There are more than 8192 areas for the value : " & cell.Value _
                     & vbNewLine & "It is not possible to copy the visible data." _
                     & vbNewLine & "Tip: Sort your data before you use this macro.", _
                       vbOKOnly, "Split in worksheets"
            Else
                'Add a new worksheet
                Set WSNew = Worksheets.Add(After:=Sheets(Sheets.Count))
                On Error Resume Next
                WSNew.Name = cell.Value
                If Err.Number > 0 Then
                    ErrNum = ErrNum + 1
                    WSNew.Name = "Error_" & Format(ErrNum, "0000")
                    Err.Clear
                End If
                On Error GoTo 0

                'Copy the visible data to the new worksheet
                My_Range.SpecialCells(xlCellTypeVisible).Copy
                With WSNew.Range("A1")
                    ' Paste:=8 will copy the columnwidth in Excel 2000 and higher
                    ' Remove this line if you use Excel 97
                    .PasteSpecial Paste:=8
                    .PasteSpecial xlPasteValues
                    .PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                    .Select
                End With
            End If

            'Show all data in the range
            My_Range.AutoFilter Field:=FieldNum

        Next cell

        'Delete the ws2 sheet
        On Error Resume Next
        Application.DisplayAlerts = False
        .Delete
        Application.DisplayAlerts = True
        On Error GoTo 0

    End With

    'Turn off AutoFilter
    My_Range.Parent.AutoFilterMode = False

    If ErrNum > 0 Then
        MsgBox "Rename every WorkSheet name that start with ""Error_"" manually" _
             & vbNewLine & "There are characters in the name that are not allowed" _
             & vbNewLine & "in a sheet name or the worksheet already exist."
    End If

    'Restore ScreenUpdating, Calculation, EnableEvents, ....
    My_Range.Parent.Select
    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With

End Sub


Function LastRow(sh As Worksheet)
    On Error Resume Next
    LastRow = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlValues, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function

