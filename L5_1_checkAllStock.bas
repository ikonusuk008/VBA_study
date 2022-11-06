Attribute VB_Name = "L5_1_checkAllStock"
Option Explicit

'(1)-1��ʈړ��҂��̕��ׂ��y���iwaitBrowsing�v���V�[�W�����Q�Ɓj
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'(1)-1�����܂�

Private Const ROW_START As Long = 14 '�X�܈ꗗ�̊J�n�s
Private Const COL_SHOPNAME As Long = 1 '�X�ܖ��̗�
Private Const COL_STOCK As Long = 2 '�݌ɏ󋵂̗�


Public Sub checkAllStock()
    Dim sht As Worksheet
    Set sht = ActiveSheet

    Dim ie As InternetExplorer
    Set ie = getIE("TSUTAYA")
    
    Dim i As Long
    i = ROW_START
    
    Do While Trim(sht.Cells(i, COL_SHOPNAME).Value) <> ""
        sht.Cells(i, COL_STOCK).Value = getStock(sht.Cells(i, COL_SHOPNAME).Value, ie)
        i = i + 1
    Loop
    
    MsgBox "�������������܂���"
End Sub


Private Function getStock(ShopName As String, ie As InternetExplorer) As String
'    Dim ie As InternetExplorer
'    Set ie = getIE("TSUTAYA")
    Dim htdoc As HTMLDocument
    Set htdoc = ie.document
    
    Dim img As HTMLImg
    
    For Each img In htdoc.getElementsByTagName("IMG")
        If InStr(img.alt, "�X�܂��w�肵�č݌Ɍ���") > 0 Then
            img.Click
            Exit For
        End If
    Next
    waitBrowsing ie

    htdoc.getElementsByName("SearchKey1")(0).Value = ShopName

'(2)�����{�^���������ӏ����C��
'�@IMG�^�O����INPUT�^�O�ւ̕ύX�ɑΉ�
'�A�ύX�̑����ӏ��̂��߁A�f�[�^�^��ėp�I��IHTMLElement�^�Ƃ���

'    For Each img In htdoc.getElementsByTagName("IMG")
'        If InStr(img.className, "tolCstCondSearchBtn") > 0 Then
'            img.Click
'            Exit For
'        End If
'    Next

    Dim searchBtn As IHTMLElement
    For Each searchBtn In htdoc.getElementsByTagName("INPUT")
        If InStr(searchBtn.className, "tolCstCondSearchBtn") > 0 Then
            searchBtn.Click
            Exit For
        End If
    Next
'(2)�����܂�

    waitBrowsing ie

    Dim zaiko_anchor As HTMLAnchorElement
    For Each zaiko_anchor In htdoc.getElementsByTagName("A")
        If InStr(zaiko_anchor.className, "zaiko_btn") > 0 Then
            zaiko_anchor.Click
            Exit For
        End If
    Next
    waitBrowsing ie

'(3)�݌ɏ��̎擾���@���C��
'�@SPAN�^�O����DIV�^�O�ւ̕ύX�ɑΉ�
'�A�ύX�̑����ӏ��̂��߁A�f�[�^�^��ėp�I��IHTMLElement�^�Ƃ���

'    Dim zaiko_span As HTMLSpanElement
'    For Each zaiko_span In htdoc.getElementsByTagName("SPAN")
'        If InStr(zaiko_span.className, "tolShStkInMrk") > 0 Then
'            getStock = zaiko_span.innerText
'            Exit Function
'        End If
'    Next
    
    Dim zaikoLabel As IHTMLElement
    For Each zaikoLabel In htdoc.getElementsByTagName("DIV")
        If InStr(zaikoLabel.className, "state") > 0 Then
            getStock = zaikoLabel.innerText
            Exit Function
        End If
    Next
'(3)�����܂�

End Function


Private Function getIE(arg_title As String) As InternetExplorer
    Dim ie As InternetExplorer
    Dim sh As Object
    Dim win As Object
    Dim document_title As String
    Set sh = CreateObject("Shell.Application")
    For Each win In sh.Windows
        document_title = ""
        On Error Resume Next
        document_title = win.document.Title
        On Error GoTo 0
        If InStr(document_title, arg_title) > 0 Then
            Set ie = win
            Exit For
        End If
    Next
    Set getIE = ie
End Function

Private Sub waitBrowsing(ie As InternetExplorer)
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        '(1)-2��ʈړ��҂��̕��ׂ��y��
        Sleep 1
        '(1)-2�����܂�
        DoEvents
    Loop
End Sub








