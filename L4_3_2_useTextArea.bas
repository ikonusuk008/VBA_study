Attribute VB_Name = "L4_3_2_useTextArea"
Sub useTextArea()
    Dim ie As InternetExplorer
    Dim txtAreaInput As HTMLTextAreaElement
    Dim txtAreaOutput As HTMLTextAreaElement
    
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-3.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
    DoEvents
    Loop
    
    Set txtAreaInput = ie.document.getElementById("TextArea1")
    MsgBox txtAreaInput.Value
    
    Set txtAreaOutput = ie.document.getElementById("TextArea2")
    txtAreaOutput.Value = "VBA‚©‚ç‚Ì‘‚«ž‚Ý" & vbCrLf & "ŽŸ‚Ìs‚Ö‚Ì‘‚«ž‚Ý"
End Sub
