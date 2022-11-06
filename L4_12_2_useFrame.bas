Attribute VB_Name = "L4_12_2_useFrame"
Sub useFrame()
    Dim ie As InternetExplorer
    Dim htdoc As HTMLDocument
    Dim htdoc_frame As HTMLDocument
    Dim anchor As HTMLAnchorElement

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/frame2/4-12_1.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    Set htdoc = ie.document
    Set htdoc_frame = htdoc.frames("menu").document

    For Each anchor In htdoc_frame.getElementsByTagName("A")
        If anchor.innerText = "V’…î•ñ" Then
            anchor.Click
            Exit For
        End If
    Next
End Sub
