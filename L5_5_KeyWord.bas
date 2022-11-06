Attribute VB_Name = "L5_5_KeyWord"

Option Explicit

Sub KeyWord()
    
    Dim i As Long  'ƒJƒEƒ“ƒ^[•Ï”
    Dim r As Long  'ƒZƒ‹s•Ï”
        r = 5      'Œ‹‰Ê‚ğ‘‚«‚ŞƒZƒ‹ŠJns
    Dim colURL As New Collection    'URLƒRƒŒƒNƒVƒ‡ƒ“
    Dim a As Variant    'ƒRƒŒƒNƒVƒ‡ƒ“—v‘fŠi”[—p•Ï”
    
    Dim StrURL As String    'ŠJnURL
    Dim StrDmn As String    'ƒhƒƒCƒ“w’è
    Dim StrWord(1 To 5) As String   'ƒL[ƒ[ƒhŠi”[—p‚Ì”z—ñ•Ï”
    Dim StrTitle As String  'ƒy[ƒWƒ^ƒCƒgƒ‹
    Dim StrText As String   'ƒy[ƒW–{•¶
    '|||||||||||||||||||||||
    'İ’è‚Ìæ“¾
    '|||||||||||||||||||||||
    StrURL = Cells(4, 3)    'ŠJnURL‚ğŠi”[
    StrDmn = Cells(5, 3)    'ƒhƒƒCƒ“‚ğŠi”[
    For i = 1 To 5
        StrWord(i) = Cells(5 + i, 3)  'ƒL[ƒ[ƒh‚ğ‡‚É”z—ñ‚ÉŠi”[
    Next i
    Range(Cells(5, 5), Cells(Rows.Count, 12)).ClearContents 'Œ‹‰Ê—“‚ğƒNƒŠƒA[
    
    '|||||||||||||||||||||||
    'IEƒIƒuƒWƒFƒNƒg‚Ìİ’èAw’èƒy[ƒW‚ğŠJ‚­
    '|||||||||||||||||||||||
    Dim objIE As Object
    Set objIE = CreateObject("InternetExplorer.application")
    objIE.Visible = True
    objIE.navigate StrURL
    
    Do While objIE.Busy Or objIE.readyState <> 4  ':READYSTATE_COMPLETE
        DoEvents
    Loop
    
    '|||||||||||||||||||||||
    'ƒŠƒ“ƒNæURL‚Ìæ“¾,ƒRƒŒƒNƒVƒ‡ƒ“‚ÉŠi”[
    '|||||||||||||||||||||||
    colURL.Add StrURL, StrURL   'ƒXƒ^[ƒgƒy[ƒW‚ÌURL‚ğƒRƒŒƒNƒVƒ‡ƒ“‚ÉŠi”[
    If objIE.document.Links.Length > 0 Then     'ƒŠƒ“ƒN‚ª‚ ‚éê‡‚Í
        For i = 0 To objIE.document.Links.Length - 1  'ƒŠƒ“ƒN‚Ì‚ ‚éƒIƒuƒWƒFƒNƒg‚ğ‡‚ÉŠm”F
            a = objIE.document.Links(i).href    'ƒŠƒ“ƒNæ‚ğæ“¾
            If StrDmn = "" Or InStr(a, StrDmn) > 0 Then 'ƒhƒƒCƒ“w’è‚ª‚È‚¢A‚Ü‚½‚ÍAw’è‚µ‚½ƒhƒƒCƒ“‚ÉŠY“–‚·‚é‚È‚ç
                If InStr(a, "@") = 0 Then   'ƒ[ƒ‹ƒAƒhƒŒƒX‚ğœ‚­
                    On Error Resume Next
                    colURL.Add a, a     'ƒŠƒ“ƒNæURL‚ğƒRƒŒƒNƒVƒ‡ƒ“‚É’Ç‰Á
                    On Error GoTo 0
                End If
            End If
        Next i
    End If

    '|||||||||||||||||||||||
    'ƒŠƒ“ƒNæ‚ğŠJ‚«ƒL[ƒ[ƒh‚Ì—L–³‚ğƒ`ƒFƒbƒN
    '|||||||||||||||||||||||
    For Each a In colURL    'ƒŠƒ“ƒNæURL‚ğ‡‚Éæ‚èo‚µ
        objIE.navigate a   'ƒŠƒ“ƒNæ‚ğŠJ‚­
        
        Do While objIE.Busy Or objIE.readyState <> 4 ':READYSTATE_COMPLETE
            DoEvents
        Loop
        
        StrTitle = ""
        StrText = ""
        On Error Resume Next
            StrTitle = objIE.document.Title   'ƒ^ƒCƒgƒ‹‚ğæ“¾
            StrText = objIE.document.body.innertext     'ƒy[ƒW–{•¶‚ğæ“¾
        On Error GoTo 0
            
        Cells(r, 5) = StrTitle              'ƒ^ƒCƒgƒ‹‚ğƒZƒ‹‚É‘‚«‚İ
        Cells(r, 6) = a                     'URL‚ğƒZƒ‹‚É‘‚«‚İ
        
        For i = 1 To 5      'ƒL[ƒ[ƒh‚Ì”‚¾‚¯ƒ‹[ƒv
            Cells(r, 6 + i) = "|"  'ˆê’UAŒ‹‰Ê‚ğ"|"‚É
            If StrWord(i) <> "" Then    'ƒL[ƒ[ƒhw’è‚ª‚ ‚Á‚½‚ç
                If InStr(StrText, StrWord(i)) > 0 Then   'w’èƒL[ƒ[ƒh‚ª‚ ‚ê‚Î
                    Cells(r, 6 + i) = "Z"  'Œ‹‰Ê‚ğƒZƒ‹‚É‘‚«‚İ
                End If
            End If
        Next i
        r = r + 1
    Next
    
    MsgBox "Š®—¹‚µ‚Ü‚µ‚½"

End Sub
