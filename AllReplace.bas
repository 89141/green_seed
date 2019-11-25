Sub AllReplace()
    'RegExpオブジェクトの作成
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    '正規表現の指定
    With reg
        .Pattern = "(.*)[0-9]+$"      'パターンを指定
        .IgnoreCase = True    '大文字と小文字を区別するか(False)、しないか(True)
        .Global = True          '文字列全体を検索するか(True)、しないか(False)
    End With
  Dim myRange As Range    '検索にヒットしたセルを記録
  Dim firstCell As Range  '最初に検索にヒットしたセルを記録
 
  Set myRange = Cells.Find(what:="大安")  '最初のFindで“大安”という語を検索するように設定
  If Not myRange Is Nothing Then  '検索対象が1件以上ある場合に下記コードを実行
      Set firstCell = myRange
      Do
          Set myRange = Cells.FindNext(myRange)   '（）内のセルの次のセルから検索を続行
          myRange.Value = reg.Replace(myRange.Text, ",")
      Loop While myRange.Address <> firstCell.Address '最初のセルに戻るまでDo~Loopを継続
  End If
End Sub
