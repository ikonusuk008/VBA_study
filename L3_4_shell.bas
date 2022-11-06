Attribute VB_Name = "L3_4_shell"
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
Sub SearchIE1()

    Dim colSh As Object    '起動中のShellWindow一式を格納する変数
    Dim win As Object   'ShellWindowを格納する変数
    Dim strTemp As String   'IEのタイトルを格納する変数
    Dim objIE As Object     '目的のIEを格納する変数
    Set colSh = CreateObject("Shell.Application")  '現在開いている IE と エクスプローラをcolShに格納
    'ColShからWindowを1つずつ取り出す
    For Each win In colSh.Windows
        'HTMLDocumentだったら
        If TypeName(win.document) = "HTMLDocument" Then
            'タイトルバーにPC Watchが含まれるか判定
            If InStr(win.document.Title, "PC Watch") > 0 Then
                '変数ieに取得したwinを格納
                Set objIE = win
                'ループを抜ける
                Exit For
            End If
        End If
    Next
    
    If objIE Is Nothing Then
        MsgBox "探しているIEはありませんでした"
    Else   'タイトルを表示する
        MsgBox objIE.document.Title & "がありました"
    End If

End Sub

       
'タイトルから起動中のIEを取得する
Sub SearchIE2()

    Dim colSh As Object    '起動中のShellWindow一式を格納する変数
    Dim win As Object   'ShellWindowを格納する変数
    Dim strTemp As String   'IEのタイトルを格納する変数
    Dim objIE As Object     '目的のIEを格納する変数
    Set colSh = CreateObject("Shell.Application")  '現在開いている IE と エクスプローラをcolShに格納
    'ColShからWindowを1つずつ取り出す
    For Each win In colSh.Windows
        strTemp = ""
        'タイトルが取得できない場合も処理を継続
        On Error Resume Next
        strTemp = win.document.Title
        On Error GoTo 0
        'タイトルバーにPC Watchが含まれるか判定
        If InStr(strTemp, "PC Watch") > 0 Then
            '変数ieに取得したwinを格納
            Set objIE = win
            'ループを抜ける
            Exit For
        End If
    Next
    
    If objIE Is Nothing Then
        MsgBox "探しているIEはありませんでした"
    Else   'タイトルを表示する
        MsgBox objIE.document.Title & "がありました"
    End If

End Sub


'ページ中の文字列から起動中のIEを取得する
Sub SearchIE3()

    Dim colSh As Object    '起動中のShellWindow一式を格納する変数
    Dim win As Object   'ShellWindowを格納する変数
    Dim strTemp As String   'ページの文字列を格納する変数
    Dim objIE As Object     '目的のIEを格納する変数
    Set colSh = CreateObject("Shell.Application")  '現在開いている IE と エクスプローラをcolShに格納
    'ColShからWindowを1つずつ取り出す
    For Each win In colSh.Windows
        strTemp = ""
        'ページの文字列が取得できない場合も処理を継続
        On Error Resume Next
        strTemp = win.document.body.innertext
        On Error GoTo 0
        'ページ上に文字列「アップデート情報」が存在するか判定
        If InStr(strTemp, "アップデート情報") > 0 Then
            '変数ieに取得したwinを格納
            Set objIE = win
            'ループを抜ける
            Exit For
        End If
    Next
    If objIE Is Nothing Then
        MsgBox "探しているIEはありませんでした"
    Else   'タイトルを表示する
        MsgBox "探しているIEがありました"
    End If

End Sub


'指定されたウィンドウを最前面化する
Sub showForeground(objIE As Object)

    '最小化されている場合は元の大きさに戻す(9=RESTORE:最小化前の状態)
    If IsIconic(objIE.hWnd) Then
        ShowWindowAsync objIE.hWnd, &H9
    End If
    '最前面に表示
    SetForegroundWindow (objIE.hWnd)

End Sub

'ページ中の文字列から起動中のIEを取得し、最前面化する
Sub SearchIE4()

    Dim colSh As Object    '起動中のShellWindow一式を格納する変数
    Dim win As Object   'ShellWindowを格納する変数
    Dim strTemp As String   'ページの文字列を格納する変数
    Dim objIE As Object     '目的のIEを格納する変数
    Set colSh = CreateObject("Shell.Application")  '現在開いている IE と エクスプローラをcolShに格納
    'ColShからWindowを1つずつ取り出す
    For Each win In colSh.Windows
        strTemp = ""
        'ページの文字列が取得できない場合も処理を継続
        On Error Resume Next
        strTemp = win.document.body.innertext
        On Error GoTo 0
        'ページ上に指定した文字列が存在するか判定
        If InStr(strTemp, Range("B5").Value) > 0 Then
            '変数ieに取得したwinを格納
            Set objIE = win
            'ループを抜ける
            Exit For
        End If
    Next
    If objIE Is Nothing Then
        MsgBox "探しているIEはありませんでした"
    Else   'タイトルを表示する
        MsgBox "探しているIEがありました"
        Call showForeground(objIE)
    End If

End Sub

