Attribute VB_Name = "L3_5_5_1_SendKeys1_"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    'SendKeys "abc"              'a��b��c�ƃL�[�X�g���[�N�𑗐M
    'SendKeys "{TAB 2}{ENTER}"   'TAB��TAB��ENTER�ƃL�[�X�g���[�N�𑗐M
    'SendKeys "+{TAB}{ENTER}"    'SHIHT+TAB��ENTER�ƃL�[�X�g���[�N�𑗐M
    'SendKeys�@"%Y"      �@      'ALT+Y�ƃL�[�X�g���[�N�𑗐M

Sub SendKeys1()
    Dim objIE As Object
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://book.impress.co.jp/appended/3384/IE.html"
    Do While objIE.Busy
        Sleep 1000
        
        
        SendKeys "{ENTER}", True
    Loop
    
End Sub

