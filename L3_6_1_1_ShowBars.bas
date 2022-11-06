Attribute VB_Name = "L3_6_1_1_ShowBars"
Sub ShowBars()
    Dim ie As InternetExplorer
    
    Set ie = CreateObject("InternetExplorer.Application")
    
    ie.Visible = True
    
    ie.Toolbar = True
    
    ie.AddressBar = True
    
    ie.MenuBar = True
    
    ie.StatusBar = True

End Sub
