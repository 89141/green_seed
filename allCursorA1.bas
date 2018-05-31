
Sub allCursorA1()

Dim ws As Worksheet

  For Each ws In Worksheets
    ws.Select
    Range("A1").Select
  Next ws

  Sheets(1).Select

End Sub
