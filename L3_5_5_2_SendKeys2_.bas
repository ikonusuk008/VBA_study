Attribute VB_Name = "L3_5_5_2_SendKeys2_"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    'SendKeys "abc"              'a→b→cとキーストロークを送信
    'SendKeys "{TAB 2}{ENTER}"   'TAB→TAB→ENTERとキーストロークを送信
    'SendKeys "+{TAB}{ENTER}"    'SHIHT+TAB→ENTERとキーストロークを送信
    'SendKeys　"%Y"      　      'ALT+Yとキーストロークを送信

Sub SendKeys2()
    Dim objIE As Object
   'IEを開いてファイルの保存URLを開く
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://book.impress.co.jp/appended/3384/excel.zip"
    
    '3秒休んでからALT+Sを送信
    Sleep 3000
    
    SendKeys "%S", True
    
End Sub
