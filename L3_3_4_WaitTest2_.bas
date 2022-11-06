Attribute VB_Name = "L3_3_4_WaitTest2_"
Sub WaitTest2()

    Dim ie As InternetExplorer

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-7.html"

    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        Debug.Print ie.Busy & ":" & ie.readyState
        DoEvents
    Loop

    MsgBox ie.document.body.innertext

End Sub
