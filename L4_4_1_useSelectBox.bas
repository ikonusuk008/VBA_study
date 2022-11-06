Attribute VB_Name = "L4_4_1_useSelectBox"
Sub useSelectBox()
    Dim ie As InternetExplorer
    Dim select1 As HTMLSelectElement
    Dim select2 As HTMLSelectElement

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-4.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    Set select1 = ie.document.getElementById("Select1")
    MsgBox select1.Value
    MsgBox select1(select1.selectedIndex).Text

    Set select2 = ie.document.getElementById("Select2")
    select2.Value = "03"
End Sub
