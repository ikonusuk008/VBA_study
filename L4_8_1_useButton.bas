Attribute VB_Name = "L4_8_1_useButton"
Sub useButton()
    Dim ie As InternetExplorer
    Dim button As HTMLInputElement
    
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-8.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    For Each button In ie.document.getElementsByTagName("INPUT")
        If button.Type = "button" And button.Value = "ƒ{ƒ^ƒ“‚Q" Then
            button.Click
            Exit For
        End If
    Next
End Sub
