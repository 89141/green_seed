Attribute VB_Name = "Module1"
Sub ChangeSheetName()
Attribute ChangeSheetName.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' ChangeSheetName Macro
'
If SearchSheet("シート編集用") Then
    Application.DisplayAlerts = False
    Worksheets("シート編集用").Delete
    Application.DisplayAlerts = True
End If


Worksheets().Add After:=Worksheets(Worksheets.Count)
ActiveSheet.name = "シート編集用"
Cells(1, 1) = "現在のシート"
Cells(1, 2) = "変更後のシート"
Cells(1, 3) = "シート名"
    
Dim i, u As Long
u = 2
For i = 1 To Worksheets.Count
    If ActiveSheet.name <> Worksheets(i).name Then
        Cells(u, 1) = Worksheets(i).name
        Cells(u, 2) = Worksheets(i).name
        Cells(u, 3) = Worksheets(i).name
        u = u + 1
    End If
Next
    
ActiveSheet.Buttons.Add(167.5, 4.5, 96, 25.5).Select
ActiveSheet.Shapes.Range(Array("Button 1")).Select
Selection.Characters.Text = "変更"
Selection.OnAction = "PERSONAL.XLSB!DoChange"

End Sub
Sub DoChange()
'
' DoChange
'
    
Dim ws As Worksheet
Set ws = Worksheets("シート編集用")

Dim y, u As Long
u = 1
For y = 2 To 50
    If IsEmpty(ws.Cells(y, 2).Value) = True Then
        GoTo LoopLast
    End If
    
    If SearchSheet(ws.Cells(y, 2).Value) Then
        Worksheets(ws.Cells(y, 2).Value).Move Before:=Worksheets(u)
    Else
        Worksheets().Add Before:=Worksheets(u)
        ActiveSheet.name = ws.Cells(y, 2).Value
    End If
    u = u + 1

LoopLast:
Next

Dim i, v As Long
v = 1
For i = 1 To Worksheets.Count
     If IsEmpty(ws.Cells(i, 2).Value) = False Then
        Worksheets(v).name = ws.Cells(i, 3)
        v = v + 1
    End If
Next
    
Application.DisplayAlerts = False
Worksheets("シート編集用").Delete
Application.DisplayAlerts = True
    
End Sub
Function SearchSheet(ByVal name As String) As Boolean

Dim i As Long
For i = 1 To Worksheets.Count
    If Worksheets(i).name = name Then
        SearchSheet = True
    End If
Next

End Function
