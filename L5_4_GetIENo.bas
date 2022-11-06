Attribute VB_Name = "L5_4_GetIENo"
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'強制的に最前面にさせる
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'最小化されているか調べる
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'元の大きさに戻すAPI
Private Declare PtrSafe Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long


'起動中のShellの数を取得・表示する
    
Sub GetIENo()

    Dim colSh As Object    '起動中のShellを格納する変数
    Set colSh = CreateObject("Shell.Application")  '現在開いている IE と エクスプローラをcolShに格納
    MsgBox colSh.Windows.Count

End Sub
       
'タイトルから起動中のIEを取得する
 Sub GetGN()
    
    Dim colobj As Object
    Dim obj As Object

    Dim colSh As Object    '起動中のShellWindow一式を格納する変数
    Dim win As Object   'ShellWindowを格納する変数
    Dim strTemp As String   'IEのタイトルを格納する変数
    Dim objie As Object     '目的のIEを格納する変数
    Set colSh = CreateObject("Shell.Application")  '現在開いている IE と エクスプローラをcolShに格納
    'ColShからWindowを1つずつ取り出す
    For Each win In colSh.Windows
        strTemp = ""
        'タイトルが取得できない場合も処理を継続
        On Error Resume Next
        strTemp = win.Document.Title
        On Error GoTo 0
        'タイトルバーに”食べログ”が含まれるか判定
        If InStr(strTemp, "グルメ・レストラン予約サイト") > 0 Then
            '変数ieに取得したwinを格納
            Set objie = win
            'ループを抜ける
            Exit For
        End If
    Next
    
    If objie Is Nothing Then
        MsgBox "探しているIEはありませんでした"
        Exit Sub
    Else   'タイトルを表示する
        MsgBox objie.Document.Title & "がありました"
        Call showForeground(objie)
    End If
    
    '一覧取得
    Dim r As Long
    Sheets("MAIN").Select
    Cells.ClearContents
    
    Cells(1, 1) = "NO"
    Cells(1, 2) = "店名"
    Cells(1, 3) = "URL"
    Cells(1, 4) = "点数"
    Cells(1, 5) = "口コミ件数"
    Cells(1, 6) = "夜の予算"
    Cells(1, 7) = "昼の予算"

    Dim i As Long
    Dim i2 As Long
    r = 1
Start:
    Cells(r, 1).Select
    With objie.Document
        'ページ上部を走査対象外とすることで、お店に限定する
        For i = 700 To .all.Length - 1
        'DIV,DIV,DIV,P,Aという並びが出現する箇所を探す
            If .all(i).TagName = "LI" Then
             If .all(i + 1).TagName = "DIV" Then
              If .all(i + 2).TagName = "DIV" Then
               If .all(i + 3).TagName = "P" Then
                If .all(i + 4).TagName = "A" And .all(i + 4).innertext <> "" _
                    And .all(i + 4).innertext <> "レストランの新規登録ページ" Then
                    '店名があると判断
                    r = r + 1
                    Cells(r, 1) = r - 1
                    Cells(r, 2) = .all(i + 4).innertext '店名
                    Cells(r, 3) = .all(i + 4).href 'URL
                    '以降のタグから、目印のSPANタグを走査
                    For i2 = i To .all.Length - 1
                        If .all(i2).TagName = "SPAN" Then
                            '夜の点数があれば、そのタグを起点に値を取得
                            If InStr(.all(i2).innertext, "夜の予算") > 0 Then
                                Cells(r, 4) = .all(i2 - 7).innertext  '"点数"
                                '点数がない場合の修正
                                If InStr(Cells(r, 4), "件") > 0 Then
                                    Cells(r, 4) = "-"
                                End If
                                Cells(r, 5) = .all(i2 - 3).innertext '"口コミ件数"
                                Cells(r, 6) = .all(i2 + 1).innertext   '"夜の予算"
                                Cells(r, 7) = .all(i2 + 5).innertext  '"昼の予算"
                                Exit For
                            End If
                        End If
                    Next i2
                End If
               End If
              End If
             End If
             
            '次があるか判定
            ElseIf .all(i).TagName = "A" Then
                If .all(i).innertext = "次の20件" Then
                    '次のページへ遷移
                    .all(i).Click
                    Call waitNavigation(objie)
                    GoTo Start
                ElseIf .all(i).innertext = "レストランの新規登録ページ" Then
                    '最終ページと判断
                    Exit For
                End If
            End If
        Next i
    End With
    
    MsgBox "一覧表を作成しました"

End Sub


'ページ中の文字列から起動中のIEを取得する
Function SearchIE(strTarget As String) As Object

    Dim colSh As Object    '起動中のShellWindow一式を格納する変数
    Dim win As Object   'ShellWindowを格納する変数
    Set colSh = CreateObject("Shell.Application")  '現在開いている IE と エクスプローラをcolShに格納
    'ColShからWindowを1つずつ取り出す
    For Each win In colSh.Windows
        strTemp = ""
        'ページの文字列が取得できない場合も処理を継続
        On Error Resume Next
        strTemp = win.Document.body.innertext
        On Error GoTo 0
        If InStr(strTemp, strTarget) > 0 Then
            '変数ieに取得したwinを格納
            Set SearchIE = win
            'ループを抜ける
            Exit For
        End If
    Next
    If SearchIE Is Nothing Then
        MsgBox "探しているIEはありませんでした"
    Else   'タイトルを表示する
        MsgBox "探しているIEがありました"
    End If

End Function

'指定されたウィンドウを最前面化する
Sub showForeground(objie As Object)

    '最小化されている場合は元の大きさに戻す(9=RESTORE:最小化前の状態)
    If IsIconic(objie.hWnd) Then
        ShowWindowAsync objie.hWnd, &H9
    End If
    '最前面に表示
    SetForegroundWindow (objie.hWnd)

End Sub

Sub SearchGN()
    
    Dim IE As Object
    Dim i As Long
    
    'IEを開いて操作対象画面へ遷移
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    IE.Navigate2 "http://tabelog.com/"
    Call waitNavigation(IE)

End Sub

Sub waitNavigation(IE As Object)
    Do While IE.Busy Or (IE.readyState <> 4 And IE.readyState <> 3)
        DoEvents
    Loop

End Sub

Sub MakeIchiran()

    Dim objie As Object
    Set objie = SearchIE("食べログ")
    Call Ichiran_Make(objie)
    MsgBox "一覧を作成しました"

End Sub

Sub Ichiran_Make(objie As Object)
    Dim n As Long
    Dim objTAG As Object
    Application.ScreenUpdating = False
    On Error Resume Next
    n = 0
    
    For Each objTAG In objie.Document.all
        n = n + 1
        Cells(n + 2, 1) = objie.Name
        Cells(n + 2, 4) = "'" & TypeName(objTAG) 'TypeNameでオブジェクトのタイプを表示
        Cells(n + 2, 5) = "'" & objTAG.TagName   'タグの名前
        Cells(n + 2, 6) = n
        Cells(n + 2, 7) = objTAG.Name
        Cells(n + 2, 8) = "'" & Left(objTAG.innertext, 256)
        Cells(n + 2, 9) = "'" & Left(objTAG.InnerHTML, 256)
        Cells(n + 2, 10) = "'" & Left(objTAG.OuterHTML, 256)
    Next
            
End Sub


