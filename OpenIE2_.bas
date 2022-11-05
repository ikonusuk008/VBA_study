Attribute VB_Name = "OpenIE2_"
Sub OpenIE2()
    Dim ie As InternetExplorer

    Set ie = CreateObject("InternetExplorer.Application")

    ie.Visible = True
End Sub
