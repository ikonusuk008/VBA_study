Attribute VB_Name = "L5_5_KeyWord"

Option Explicit

Sub KeyWord()
    
    Dim i As Long  'カウンター変数
    Dim r As Long  'セル行変数
        r = 5      '結果を書き込むセル開始行
    Dim colURL As New Collection    'URLコレクション
    Dim a As Variant    'コレクション要素格納用変数
    
    Dim StrURL As String    '開始URL
    Dim StrDmn As String    'ドメイン指定
    Dim StrWord(1 To 5) As String   'キーワード格納用の配列変数
    Dim StrTitle As String  'ページタイトル
    Dim StrText As String   'ページ本文
    '−−−−−−−−−−−−−−−−−−−−−−−
    '設定の取得
    '−−−−−−−−−−−−−−−−−−−−−−−
    StrURL = Cells(4, 3)    '開始URLを格納
    StrDmn = Cells(5, 3)    'ドメインを格納
    For i = 1 To 5
        StrWord(i) = Cells(5 + i, 3)  'キーワードを順に配列に格納
    Next i
    Range(Cells(5, 5), Cells(Rows.Count, 12)).ClearContents '結果欄をクリアー
    
    '−−−−−−−−−−−−−−−−−−−−−−−
    'IEオブジェクトの設定、指定ページを開く
    '−−−−−−−−−−−−−−−−−−−−−−−
    Dim objIE As Object
    Set objIE = CreateObject("InternetExplorer.application")
    objIE.Visible = True
    objIE.navigate StrURL
    
    Do While objIE.Busy Or objIE.readyState <> 4  ':READYSTATE_COMPLETE
        DoEvents
    Loop
    
    '−−−−−−−−−−−−−−−−−−−−−−−
    'リンク先URLの取得,コレクションに格納
    '−−−−−−−−−−−−−−−−−−−−−−−
    colURL.Add StrURL, StrURL   'スタートページのURLをコレクションに格納
    If objIE.document.Links.Length > 0 Then     'リンクがある場合は
        For i = 0 To objIE.document.Links.Length - 1  'リンクのあるオブジェクトを順に確認
            a = objIE.document.Links(i).href    'リンク先を取得
            If StrDmn = "" Or InStr(a, StrDmn) > 0 Then 'ドメイン指定がない、または、指定したドメインに該当するなら
                If InStr(a, "@") = 0 Then   'メールアドレスを除く
                    On Error Resume Next
                    colURL.Add a, a     'リンク先URLをコレクションに追加
                    On Error GoTo 0
                End If
            End If
        Next i
    End If

    '−−−−−−−−−−−−−−−−−−−−−−−
    'リンク先を開きキーワードの有無をチェック
    '−−−−−−−−−−−−−−−−−−−−−−−
    For Each a In colURL    'リンク先URLを順に取り出し
        objIE.navigate a   'リンク先を開く
        
        Do While objIE.Busy Or objIE.readyState <> 4 ':READYSTATE_COMPLETE
            DoEvents
        Loop
        
        StrTitle = ""
        StrText = ""
        On Error Resume Next
            StrTitle = objIE.document.Title   'タイトルを取得
            StrText = objIE.document.body.innertext     'ページ本文を取得
        On Error GoTo 0
            
        Cells(r, 5) = StrTitle              'タイトルをセルに書き込み
        Cells(r, 6) = a                     'URLをセルに書き込み
        
        For i = 1 To 5      'キーワードの数だけループ
            Cells(r, 6 + i) = "−"  '一旦、結果を"−"に
            If StrWord(i) <> "" Then    'キーワード指定があったら
                If InStr(StrText, StrWord(i)) > 0 Then   '指定キーワードがあれば
                    Cells(r, 6 + i) = "〇"  '結果をセルに書き込み
                End If
            End If
        Next i
        r = r + 1
    Next
    
    MsgBox "完了しました"

End Sub
