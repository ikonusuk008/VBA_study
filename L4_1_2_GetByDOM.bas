Attribute VB_Name = "L4_1_2_GetByDOM"
Sub GetByDOM()
    Dim ie As InternetExplorer

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-1.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    Dim htdoc As HTMLDocument
    Set htdoc = ie.document

    MsgBox htdoc.getElementsByTagName("LI")(1).innerText
End Sub

