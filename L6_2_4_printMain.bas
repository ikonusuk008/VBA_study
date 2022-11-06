Attribute VB_Name = "L6_2_4_printMain"
Option Explicit

Public Sub printMain()
    Dim ie As InternetExplorer
'   Set ie = CreateObject("InternetExplorer.Application")
'   ie.Visible = True
'   ie.Navigate "http://www.impressjapan.jp/appended/3384/6-1.html"
'   Do While ie.Busy Or ie.ReadyState < READYSTATE_COMPLETE
'   DoEvents
'   Loop
    
    Dim DocumentTitle As String
    DocumentTitle = InputBox("解析対象画面のタイトルを入力してください")
    If DocumentTitle <> "" Then
        Set ie = getIE(DocumentTitle)
    End If
    
    If ie Is Nothing Then
        MsgBox "タイトル未入力または対象画面が見つかりません"
        Exit Sub
    End If
    
    Dim html As String
    html = getHTMLString(ie)
    
    Dim FileName As String
    FileName = ThisWorkbook.Path & "\HTML_" & Format(Now, "YYYYMMDDHHmmSS") & ".txt"
    Dim FileNum As Integer
    FileNum = FreeFile()
    
    Open FileName For Output As #FileNum
        Print #FileNum, html
    Close #FileNum
    
    MsgBox "Webページの解析が完了しました。解析結果は " & FileName & " に記録されています"
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
    ret = ret & "[" & Trim(Str(depth)) & "階層]" & vbCrLf
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
    ret = "#" & vbTab & "タグ" & vbTab & "Type" & vbTab & "ID" & vbTab & "Name" & vbTab & "Value" & vbCrLf
    Dim element As Object
    Dim i As Long
    For Each element In htdoc.all
        Select Case UCase(element.tagName)
            Case "INPUT", "TEXTAREA", "SELECT"
                ret = ret & CStr(i) & vbTab & element.tagName & vbTab & element.Type & vbTab & element.ID & vbTab & element.Name & vbTab & element.Value & vbCrLf
                If UCase(element.Type) <> "HIDDEN" Then
                    element.outerHTML = element.outerHTML & "&nbsp;<b style=""color:blue;"">[" & CStr(i) & "]</b>"
                End If
                i = i + 1
        End Select
    Next
    getElementList = ret
End Function


Private Function getIE(arg_title As String) As InternetExplorer
    Dim ie As InternetExplorer
    Dim sh As Object
    Dim win As Object
    Dim DocumentTitle As String
    
    Set sh = CreateObject("Shell.Application")
    
    For Each win In sh.Windows
        DocumentTitle = ""
        On Error Resume Next
        DocumentTitle = win.document.Title
        On Error GoTo 0
        If InStr(DocumentTitle, arg_title) > 0 Then
            Set ie = win
            Exit For
        End If
    Next

    Set getIE = ie
End Function

