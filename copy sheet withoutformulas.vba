Sub CopyWorksheetValues()
    ActiveSheet.Copy
    Cells.Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub
