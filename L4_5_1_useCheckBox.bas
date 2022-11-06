Attribute VB_Name = "L4_5_1_useCheckBox"
Sub useCheckBox()
    Dim ie As InternetExplorer
    Dim check1 As HTMLInputElement
    Dim check2 As HTMLInputElement

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-5.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    Set check1 = ie.document.getElementById("oceanview")
    Set check2 = ie.document.getElementById("nonsmoke")
    
    MsgBox "オーシャンビュー:" & check1.Checked & " / 禁煙:" & check2.Checked
    
    check1.Checked = False
    check2.Checked = True
    
'    If check1.Checked = True Then check1.Click
'    If check2.Checked = False Then check2.Click
End Sub
