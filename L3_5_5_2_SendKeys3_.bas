Attribute VB_Name = "L3_5_5_2_SendKeys3_"
Public Declare PtrSafe Function GetLastActivePopup Lib "user32" _
    (ByVal hWnd As Long) As Long

    'SendKeys "abc"              'a→b→cとキーストロークを送信
    'SendKeys "{TAB 2}{ENTER}"   'TAB→TAB→ENTERとキーストロークを送信
    'SendKeys "+{TAB}{ENTER}"    'SHIHT+TAB→ENTERとキーストロークを送信
    'SendKeys　"%Y"      　      'ALT+Yとキーストロークを送信

Sub SendKeys2_2()
    Dim objIE As Object
   'IEを開いてファイルの保存URLを開く
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://book.impress.co.jp/appended/3384/excel.zip"
    
    'ファイルを開くダイアログが表示されるまでループ
    Do While objIE.hWnd = GetLastActivePopup(objIE.hWnd)
        DoEvents
    Loop
        SendKeys "%S", True 'らALT+Sを送信
End Sub
