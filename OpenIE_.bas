Attribute VB_Name = "OpenIE_"
Sub OpenIE()
    Dim ie As Object

    Set ie = CreateObject("InternetExplorer.Application")

    ie.Visible = True
    
End Sub
