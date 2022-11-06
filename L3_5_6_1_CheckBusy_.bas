Attribute VB_Name = "L3_5_6_1_CheckBusy_"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    'SendKeys "abc"              'a→b→cとキーストロークを送信
    'SendKeys "{TAB 2}{ENTER}"   'TAB→TAB→ENTERとキーストロークを送信
    'SendKeys "+{TAB}{ENTER}"    'SHIHT+TAB→ENTERとキーストロークを送信
    'SendKeys　"%Y"      　      'ALT+Yとキーストロークを送信

Sub CheckBusy()
    Dim objIE As Object

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://book.impress.co.jp/appended/3384/IE2.html"
    
    Do
        Debug.Print "IEのBusy状態：" & objIE.Busy
        DoEvents
        Sleep 250
    Loop
End Sub


