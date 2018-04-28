
Sub 全てのカーソルをA1へ()
Dim ws As Worksheet

For Each ws In Worksheets
ws.Select
Range("A1").Select
Next ws
MsgBox "全てのシートのカーソルをA1に移動しました"
End Sub