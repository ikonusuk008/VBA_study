Attribute VB_Name = "L4_10_GetTable1"
Option Explicit

'見出し項目を探してテーブルのテキストを取得する
Sub GetTable1()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    
    'IEを開いて操作対象画面へ遷移
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table1.html"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    For i = 0 To Doc.all.Length - 1 'ドキュメントの構成タグを一つずつ調査
        If Doc.all(i).tagName = "TH" Then   'THタグなら
            If Doc.all(i).innerText = "メモリー" Then  'テキストが「メモリー」なら
                MsgBox Doc.all(i + 1).innerText 'その次のタグのテキストを取得し表示
                Exit For
            End If
        End If
    Next i

End Sub

'見出し項目を探してテーブルのテキストを取得する（見出し項目が複数回登場する場合）
Sub GetTable2()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    Dim n As Long
    
    'IEを開いて操作対象画面へ遷移
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table2.html"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    For i = 0 To Doc.all.Length - 1 'ドキュメントの構成タグを一つずつ調査
        If Doc.all(i).tagName = "TH" Then   'THタグなら
            If Doc.all(i).innerText = "メモリー" Then  'テキストが「メモリー」なら
                n = n + 1      '登場回数をカウント
                If n = 2 Then '2回目の登場なら
                    MsgBox Doc.all(i + 1).innerText 'その次のタグのテキストを取得し表示
                    Exit For
                End If
            End If
        End If
    Next i

End Sub

Public Sub waitNavigation(ie As InternetExplorer)
    Do While ie.Busy Or ie.readyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    Do While ie.document.readyState <> "complete"
        DoEvents
    Loop
End Sub

Sub GetTable3()
    Dim ie As InternetExplorer
    
    'IEを開いて操作対象画面へ遷移
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Call waitNavigation(ie)
    
    Call MakeList(ie)

End Sub

Sub MakeList(ObjIE As InternetExplorer)
    Dim n As Long   'タグの通し番号
    Dim r As Long   'TD,THタグの通し番号
    Dim Doc As HTMLDocument
    Dim ObjTag As Object 'タグ格納用
    n = 0
    r = 0
    Sheets("Sheet1").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/標準"
    Set Doc = ObjIE.document
    
    For n = 0 To Doc.all.Length - 1
        With Doc.all(n)
            If .tagName = "TD" Or .tagName = "TH" Then
                r = r + 1
                Cells(r + 1, 1) = .tagName    'タグの名前
                Cells(r + 1, 2) = n          'タグの通し番号
                Cells(r + 1, 3) = r            'TD,THタグの通し番号
                Cells(r + 1, 4) = .innerText  'テキスト
            End If
        End With
    Next
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
End Sub

Sub GetTable4()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    Dim n As Long
    
    'IEを開いて操作対象画面へ遷移
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    Sheets("Sheet3").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/標準"
    
    For i = 537 To 855 'ドキュメントの構成タグを一つずつ調査
        If Doc.all(i).tagName = "TD" Then 'TDタグなら
            n = n + 1
            Cells(Int((n - 1) / 16) + 1, (n - 1) Mod 16 + 1) = Doc.all(i).innerText
                '行はTDタグの通し番号を16で割った商+1とすることで16タグごとに改行
                'TDタグの通し番号を16で割った余りを利用して列番号を指定
        End If
    Next i
    
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

End Sub

Sub GetTable5()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    Dim n As Long
    Dim StartTag As Long
    Dim FinishTag As Long
    
    'IEを開いて操作対象画面へ遷移
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    Sheets("Sheet3").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/標準"
    
    'テーブル開始タグの取得
    For i = 0 To Doc.all.Length - 1  'ドキュメントの構成タグを一つずつ調査
        If Doc.all(i).tagName = "TH" Then 'THタグなら
            If Doc.all(i).innerText = "液晶" Then
                StartTag = i
                Exit For
            End If
        End If
    Next i

    'テーブル修了タグの取得
    For i = StartTag To Doc.all.Length - 1  'ドキュメントの構成タグを一つずつ調査
        If Doc.all(i).tagName = "TH" Then 'THタグなら
            If Doc.all(i).innerText = "メーカー" Then
                FinishTag = i
                Exit For
            End If
        End If
    Next i
    
    For i = StartTag To FinishTag 'ドキュメントの構成タグを一つずつ調査
        If Doc.all(i).tagName = "TD" Then 'TDタグなら
            n = n + 1
            Cells(Int((n - 1) / 16) + 1, (n - 1) Mod 16 + 1) = Doc.all(i).innerText
        End If
    Next i
    
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

End Sub

Sub GetTable6()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTD As Object
    Dim ObjTag As Object
    Dim n As Long
    
    'IEを開いて操作対象画面へ遷移
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    Sheets("Sheet4").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/標準"
    
    
    'TDタグのみ取得し一覧化
    Set ObjTD = Doc.getElementsByTagName("TD")
    For Each ObjTag In ObjTD
        n = n + 1
        Cells(n, 1) = n
        Cells(n, 2) = ObjTag.tagName
        Cells(n, 3) = ObjTag.innerText
    Next ObjTag
    
End Sub

Sub GetTable7()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTables As Object
    Dim ObjTag As Object
    
    Dim r As Long
    Dim c As Long
    Dim i As Long
    
    'IEを開いて操作対象画面へ遷移
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    Sheets("Sheet5").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/標準"
    
    r = 1
    c = 1
    
    'テーブル開始タグの取得
    For i = 0 To Doc.all.Length - 1  'ドキュメントの構成タグを一つずつ調査
        If Doc.all(i).tagName = "TH" Or Doc.all(i).tagName = "TD" Then
            'TD、THタグなら列を右にテキストをセルに書き込む
            c = c + 1
            Cells(r, c) = Doc.all(i).innerText
        ElseIf Doc.all(i).tagName = "TR" Then
            'TRタグなら改行し、1列目に戻す
            r = r + 1
            c = 1
        End If
    Next i
    
End Sub
