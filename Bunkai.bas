Attribute VB_Name = "Bunkai"
Option Explicit

Sub Bunkai()

    Dim r As Long
    Dim appWord As Object, objDoc As Object, objWord As Object
    Set appWord = CreateObject("Word.Application")  'Wordアプリケーションの起動
    Set objDoc = appWord.Documents.Add              '新規文書オブジェクトの作成
    
    Range(Cells(8, 2), Cells(Rows.Count, 2)).ClearContents  '単語セルのクリアー
    objDoc.Range.Text = Range("B5")   '文書オブジェクトに文章を格納
    r = 8
    For Each objWord In objDoc.Words    '順に単語を取り出し
         Cells(r, 2) = objWord         'セルに書き込み
         r = r + 1
    Next
    
    appWord.Quit SaveChanges:=0 'wdDoNotSaveChanges  'Wordの終了
    Set objDoc = Nothing    'Word文書オブジェクトのクリアー
    Set appWord = Nothing   'Wordアプリケーションのクリアー

End Sub
