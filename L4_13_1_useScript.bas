Attribute VB_Name = "L4_13_1_useScript"
Sub useScript()
    Dim ie As InternetExplorer

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-13.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    Dim pwin As HTMLWindow2
    Set pwin = ie.document.parentWindow
    
    pwin.alert ("VBA‚©‚çalert‚ðŽÀs")
End Sub




