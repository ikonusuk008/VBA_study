Attribute VB_Name = "L3_5_6_1_CheckBusy_"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    'SendKeys "abc"              'a��b��c�ƃL�[�X�g���[�N�𑗐M
    'SendKeys "{TAB 2}{ENTER}"   'TAB��TAB��ENTER�ƃL�[�X�g���[�N�𑗐M
    'SendKeys "+{TAB}{ENTER}"    'SHIHT+TAB��ENTER�ƃL�[�X�g���[�N�𑗐M
    'SendKeys�@"%Y"      �@      'ALT+Y�ƃL�[�X�g���[�N�𑗐M

Sub CheckBusy()
    Dim objIE As Object

    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://book.impress.co.jp/appended/3384/IE2.html"
    
    Do
        Debug.Print "IE��Busy��ԁF" & objIE.Busy
        DoEvents
        Sleep 250
    Loop
End Sub


