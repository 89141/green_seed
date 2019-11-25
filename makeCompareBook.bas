Sub makeComPareBook()
  '
  '現在のブックを比較用のブックに変える
  '基本的には新規に作成したブックで実行する

  '

  Dim i As Integer

  For i = 1 To 15

    '現在のシートを全選択
    Cells.Select
    Selection.NumberFormatLocal = "@"

    With Selection.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = -0.149998474074526
      .PatternTintAndShade = 0
    End With

    '現在のシートのA1を選択
    Range("A1").Select

    '次のシートが存在しない=追加,存在する=選択
    If ActiveSheet.Next Is Nothing Then
      Sheets.Add After:=ActiveSheet
    Else
      ActiveSheet.Next.Activate
    End If

    Cells.Select
    Selection.NumberFormatLocal = "@"

    With Selection.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .Color = 255
      .TintAndShade = 0
      .PatternTintAndShade = 0
    End With

    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
    "=A1<>" & ActiveSheet.Previous.name & "!A1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority

    With Selection.FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 15773696
      .TintAndShade = 0
    End With

    Selection.FormatConditions(1).StopIfTrue = False

    Range("A1").Select

    If i = 15 Then
      Sheets(1).Select
    Else
      If ActiveSheet.Next Is Nothing Then
        Sheets.Add After:=ActiveSheet
      Else
       ActiveSheet.Next.Activate
      End If
    End If

  Next i

  End Sub

