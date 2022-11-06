Attribute VB_Name = "L4_10_4_GetTable4"
Sub GetTable4()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    Dim n As Long
    
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://book.impress.co.jp/appended/3384/4-10_3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    Set Doc = ie.document
    
    Sheets("Sheet3").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/•W€"
    
    For i = 537 To 855
        If Doc.all(i).tagName = "TD" Then
            n = n + 1
            Cells(Int((n - 1) / 16) + 1, (n - 1) Mod 16 + 1) = Doc.all(i).innerText
        End If
    Next i
    
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

End Sub

