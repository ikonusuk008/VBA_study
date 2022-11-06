Attribute VB_Name = "L4_10_2_1_GetTable1"
Public Sub GetTable1()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-10_1.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    Set Doc = ie.document
    
    For i = 0 To Doc.all.Length - 1
        If Doc.all(i).tagName = "TH" Then
            If Doc.all(i).innerText = "ƒƒ‚ƒŠ[" Then
                AppActivate Application.Caption
                MsgBox Doc.all(i + 1).innerText
                Exit For
            End If
        End If
    Next i

End Sub

