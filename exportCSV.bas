Attribute VB_Name = "Module2"
Sub exportCSV()
Attribute exportCSV.VB_ProcData.VB_Invoke_Func = " \n14"
'
' exportCSV Macro
'
 
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
 
Dim csvFile As String
csvFile = ActiveWorkbook.Path & "\" & ws.Name & ".csv"
 
'ADODB.Streamオブジェクトを生成
Dim adoSt As Object
Set adoSt = CreateObject("ADODB.Stream")
 
Dim strLine As String
Dim i As Long, j As Long, x As Long
i = 1

With adoSt
    .Charset = "UTF-8"
    .LineSeparator = -1
    .Open
    
    j = 1
    Do While ws.Cells(1, j + 1).Value <> ""
        x = j
        j = j + 1
    Loop
    
    Do While ws.Cells(i, 1).Value <> ""
 
        strLine = ""
 
        j = 1
        Do While j < x + 1
 
            strLine = strLine & ws.Cells(i, j).Value & ","
            j = j + 1
 
        Loop
 
        strLine = strLine & ws.Cells(i, j).Value
 
        .WriteText strLine, 1
 
        i = i + 1
 
    Loop
    
    .Position = 0 'ストリームの位置を0にする
    .Type = 1 'データの種類をバイナリデータに変更
    .Position = 3 'ストリームの位置を3にする
 
    Dim byteData() As Byte '一時格納用
    byteData = .Read 'ストリームの内容を一時格納用変数に保存
    .Close '一旦ストリームを閉じる（リセット）
 
    .Open 'ストリームを開く
    .Write byteData 'ストリームに一時格納したデータを流し込む
    .SaveToFile csvFile, 2
    .Close
 
End With
Next
 
MsgBox "CSVに書き出しました"
 
End Sub
