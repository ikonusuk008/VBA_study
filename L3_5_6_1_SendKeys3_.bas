Attribute VB_Name = "L3_5_6_1_SendKeys3_"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetLastActivePopup Lib "user32" (ByVal hWnd As Long) As Long

    'SendKeys "abc"              'a��b��c�ƃL�[�X�g���[�N�𑗐M
    'SendKeys "{TAB 2}{ENTER}"   'TAB��TAB��ENTER�ƃL�[�X�g���[�N�𑗐M
    'SendKeys "+{TAB}{ENTER}"    'SHIHT+TAB��ENTER�ƃL�[�X�g���[�N�𑗐M
    'SendKeys�@"%Y"      �@      'ALT+Y�ƃL�[�X�g���[�N�𑗐M

Sub L3_5_6_1_SendKeys3()
    Dim objIE As Object
   'IE���J���ăt�@�C���̕ۑ�URL���J��
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://book.impress.co.jp/appended/3384/IE2.html"
    
    'Busy�̊ԁA�ҋ@
    Do While objIE.Busy
        Sleep 1
    Loop

    'Busy�ƂȂ�܂ŁA�ҋ@
    Do Until objIE.Busy
        Sleep 1
    Loop

    '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
    Do While objIE.hWnd = GetLastActivePopup(objIE.hWnd)
        DoEvents
    Loop
    
    SendKeys "%S", True  '�ۑ��������L�[���M

End Sub
