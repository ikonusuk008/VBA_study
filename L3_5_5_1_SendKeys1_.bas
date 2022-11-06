Attribute VB_Name = "L3_5_5_1_SendKeys1_"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    'SendKeys "abc"              'a→b→cとキーストロークを送信
    'SendKeys "{TAB 2}{ENTER}"   'TAB→TAB→ENTERとキーストロークを送信
    'SendKeys "+{TAB}{ENTER}"    'SHIHT+TAB→ENTERとキーストロークを送信
    'SendKeys　"%Y"      　      'ALT+Yとキーストロークを送信

Sub SendKeys1()
    Dim objIE As Object
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://book.impress.co.jp/appended/3384/IE.html"
    Do While objIE.Busy
        Sleep 1000
        
        
        SendKeys "{ENTER}", True
    Loop
    
End Sub

