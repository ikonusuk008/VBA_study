Attribute VB_Name = "L3_6_3_ExecInvisible"
Sub ExecInvisible()
    Dim ie As InternetExplorer

    Set ie = CreateObject("InternetExplorer.Application")

    ie.Visible = False

    ie.navigate "http://www.yahoo.co.jp/"

    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    MsgBox ie.document.Title

    ie.Quit
End Sub
