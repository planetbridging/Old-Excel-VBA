Sub LoopTrhgouh()
Dim sh As Worksheet
Dim rw As Range
Dim RowCount As Integer

RowCount = 0

Set sh = ActiveSheet
For Each rw In sh.Rows

  If sh.Cells(rw.Row, 1).Value = "" Then
    Exit For
  End If

  RowCount = RowCount + 1

Next rw

'MsgBox (RowCount)

End Sub
