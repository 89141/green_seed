'https://thom.hateblo.jp/entry/2015/11/22/181902
Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long

Sub AutoCapture()
    MsgBox "AutoCaptureを開始します。" & vbNewLine & _
        "終了するには任意のシートのA1セルにExitと入力してください。", vbInformation
    Dim CB As Variant
    Do While True
        CB = Application.ClipboardFormats
        If StrConv(ActiveSheet.Cells(1, 1).Value, vbUpperCase) = "EXIT" Then GoTo Quit
        If CB(1) <> -1 Then
            For i = 1 To UBound(CB)
                If CB(i) = xlClipboardFormatBitmap Then
                    Call CutAndPaste
                    
                    'クリップボードを空にする。
                    OpenClipboard
                    EmptyClipboard
                    CloseClipboard
                End If
            Next i
        End If
        DoEvents
    Loop
    
Quit:
    MsgBox "AutoCaptureを停止しました。", vbInformation
    ActiveSheet.Cells(1, 1).ClearContents
End Sub

