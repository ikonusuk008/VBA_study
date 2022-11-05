Attribute VB_Name = "WaitTest_"
Sub WaitTest()
    Dim ie As InternetExplorer

    Set ie = CreateObject("InternetExplorer.Application")

    ie.Visible = True

    ie.navigate "http://book.impress.co.jp/appended/3384/4-7.html"

    MsgBox ie.document.body.innerText
    
End Sub
