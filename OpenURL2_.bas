Attribute VB_Name = "OpenURL2_"
Sub OpenURL2()
    Dim ie As InternetExplorer

    Set ie = CreateObject("InternetExplorer.Application")

    ie.Visible = True
    
    ie.navigate "http://search.yahoo.co.jp/search?p=" & ActiveSheet.Cells(1, 1).Value
    
End Sub

