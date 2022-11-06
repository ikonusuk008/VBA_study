Attribute VB_Name = "L6_2_1_2_printMain"
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


Private Function getHTMLString(container As Object, Optional depth As Long = 0) As String
    Dim ErrorInfo As String
    Dim htdoc As HTMLDocument
    On Error Resume Next
    Set htdoc = container.document
    If Err.Number <> 0 Then
        ErrorInfo = Trim(Str(Err.Number)) & ":" & Err.Description
    End If
    On Error GoTo 0
    
    Dim ret As String
    ret = "------------------------------------------------------" & vbCrLf
    ret = ret & "[" & Trim(Str(depth)) & "ŠK‘w]" & vbCrLf
    If Not htdoc Is Nothing Then
        ret = ret & htdoc.Title & " | " & htdoc.Location & " (" & container.Name & ")" & vbCrLf
        ret = ret & "------------------------------------------------------" & vbCrLf
        ret = ret & getElementList(htdoc) & vbCrLf
        ret = ret & "------------------------------------------------------" & vbCrLf
        ret = ret & htdoc.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
        
        Dim i As Integer
        For i = 0 To htdoc.frames.Length - 1
            ret = ret & getHTMLString(htdoc.frames(i), depth + 1)
        Next
    Else
        ret = ret & "------------------------------------------------------" & vbCrLf
        ret = ret & ErrorInfo
    End If
    getHTMLString = ret
End Function


Private Function getElementList(htdoc As HTMLDocument) As String
    Dim ret As String
    ret = "ƒ^ƒO" & vbTab & "Type" & vbTab & "ID" & vbTab & "Name" & vbTab & "Value" & vbCrLf
    Dim element As Object
    For Each element In htdoc.all
        Select Case UCase(element.tagName)
            Case "INPUT", "TEXTAREA", "SELECT"
                ret = ret & element.tagName & vbTab & element.Type & vbTab & element.ID & vbTab & element.Name & vbTab & element.Value & vbCrLf
        End Select
    Next
    getElementList = ret
End Function




