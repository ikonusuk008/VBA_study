Attribute VB_Name = "L4_7_1_useAnchor"
Sub useAnchor()
    Dim ie As InternetExplorer
    Dim anchor As HTMLAnchorElement

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-7.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    For Each anchor In ie.document.getElementsByTagName("A")
        If anchor.innerText = "‚â‚«‚»‚Îƒpƒ“‚Ì‰h—{" Then
            anchor.Click
            Exit For
        End If
    Next
End Sub
