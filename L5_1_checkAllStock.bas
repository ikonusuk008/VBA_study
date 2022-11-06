Attribute VB_Name = "L5_1_checkAllStock"
Option Explicit

'(1)-1画面移動待ちの負荷を軽減（waitBrowsingプロシージャも参照）
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'(1)-1ここまで

Private Const ROW_START As Long = 14 '店舗一覧の開始行
Private Const COL_SHOPNAME As Long = 1 '店舗名の列
Private Const COL_STOCK As Long = 2 '在庫状況の列


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
    
    MsgBox "検索が完了しました"
End Sub


Private Function getStock(ShopName As String, ie As InternetExplorer) As String
'    Dim ie As InternetExplorer
'    Set ie = getIE("TSUTAYA")
    Dim htdoc As HTMLDocument
    Set htdoc = ie.document
    
    Dim img As HTMLImg
    
    For Each img In htdoc.getElementsByTagName("IMG")
        If InStr(img.alt, "店舗を指定して在庫検索") > 0 Then
            img.Click
            Exit For
        End If
    Next
    waitBrowsing ie

    htdoc.getElementsByName("SearchKey1")(0).Value = ShopName

'(2)検索ボタンを押す箇所を修正
'①IMGタグからINPUTタグへの変更に対応
'②変更の多い箇所のため、データ型を汎用的なIHTMLElement型とした

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
'(2)ここまで

    waitBrowsing ie

    Dim zaiko_anchor As HTMLAnchorElement
    For Each zaiko_anchor In htdoc.getElementsByTagName("A")
        If InStr(zaiko_anchor.className, "zaiko_btn") > 0 Then
            zaiko_anchor.Click
            Exit For
        End If
    Next
    waitBrowsing ie

'(3)在庫情報の取得方法を修正
'①SPANタグからDIVタグへの変更に対応
'②変更の多い箇所のため、データ型を汎用的なIHTMLElement型とした

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
'(3)ここまで

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
        '(1)-2画面移動待ちの負荷を軽減
        Sleep 1
        '(1)-2ここまで
        DoEvents
    Loop
End Sub








