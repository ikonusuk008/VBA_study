Attribute VB_Name = "L4_2_1_useTextbox"
Sub useTextbox()
    Dim ie As InternetExplorer
    Dim txtInput As HTMLInputElement
    Dim txtOutput As HTMLInputElement

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-2.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    Set txtInput = ie.document.getElementById("Text1")
    MsgBox txtInput.Value

    Set txtOutput = ie.document.getElementById("Text2")
    txtOutput.Value = "VBA ‚©‚ç‚Ì‘‚«ž‚Ý"
End Sub
