Attribute VB_Name = "L4_6_1_useRadio"
Sub useRadio()
    Dim ie As InternetExplorer
    Dim radio1 As HTMLInputElement
    Dim radio2 As HTMLInputElement
    
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-6.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    For Each radio1 In ie.document.getElementsByName("Radio1")
        If radio1.Checked = True Then
            MsgBox radio1.Value
            Exit For
        End If
    Next
    
    For Each radio2 In ie.document.getElementsByName("Radio2")
        If radio2.Value = "—" Then
            radio2.Checked = True
            Exit For
        End If
    Next
End Sub
