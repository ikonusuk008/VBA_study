Attribute VB_Name = "L3_5_6_1_SendKeys3_"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetLastActivePopup Lib "user32" (ByVal hWnd As Long) As Long

    'SendKeys "abc"              'a→b→cとキーストロークを送信
    'SendKeys "{TAB 2}{ENTER}"   'TAB→TAB→ENTERとキーストロークを送信
    'SendKeys "+{TAB}{ENTER}"    'SHIHT+TAB→ENTERとキーストロークを送信
    'SendKeys　"%Y"      　      'ALT+Yとキーストロークを送信

Sub L3_5_6_1_SendKeys3()
    Dim objIE As Object
   'IEを開いてファイルの保存URLを開く
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://book.impress.co.jp/appended/3384/IE2.html"
    
    'Busyの間、待機
    Do While objIE.Busy
        Sleep 1
    Loop

    'Busyとなるまで、待機
    Do Until objIE.Busy
        Sleep 1
    Loop

    'ファイルを開くダイアログが表示されるまでループ
    Do While objIE.hWnd = GetLastActivePopup(objIE.hWnd)
        DoEvents
    Loop
    
    SendKeys "%S", True  '保存を押すキー送信

End Sub
