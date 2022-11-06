Attribute VB_Name = "L6_1_1_2_printMain"
Option Explicit

Public Sub printMain()
    Dim ie As InternetExplorer
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate "http://book.impress.co.jp/appended/3384/6-1.html"
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    Dim HTMLString As String
    HTMLString = getHTMLString(ie)
    
    Dim FileName As String
    FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & ".txt"
    Dim FileNum As Integer
    FileNum = FreeFile()
    Open FileName For Output As #FileNum
        Print #FileNum, HTMLString
    Close #FileNum
End Sub


Private Function getHTMLString(ie As InternetExplorer) As String
    Dim htdoc As HTMLDocument
    Set htdoc = ie.document
    
    Dim ret As String
    ret = htdoc.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
    
    getHTMLString = ret
End Function

