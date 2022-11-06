Attribute VB_Name = "L3_6_2_ChangeSizeAndLocation"
Sub ChangeSizeAndLocation()
    Dim ie As InternetExplorer

    Set ie = CreateObject("InternetExplorer.Application")

    ie.Visible = True

    ie.Width = 400

    ie.Height = 300

    ie.Left = 700

    ie.Top = 100

    ie.resizable = False
    
End Sub
