Attribute VB_Name = "L4_11_2_useForm"
Sub useForm()
    Dim ie As InternetExplorer
    Dim form As HTMLFormElement

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-11.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    ie.document.getElementById("Text1").Value = "uezo"
    ie.document.getElementById("Select1").Value = "example.ne.jp"

    Set form = ie.document.forms("Form1")
    form.submit
End Sub
