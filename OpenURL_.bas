Attribute VB_Name = "OpenURL_"
Sub OpenURL()
    Dim ie As InternetExplorer

    Set ie = CreateObject("InternetExplorer.Application")

    ie.Visible = True
    
    ie.navigate "http://www.yahoo.co.jp/index.html"
End Sub
