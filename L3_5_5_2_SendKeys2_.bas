Attribute VB_Name = "L3_5_5_2_SendKeys2_"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    'SendKeys "abc"              'a��b��c�ƃL�[�X�g���[�N�𑗐M
    'SendKeys "{TAB 2}{ENTER}"   'TAB��TAB��ENTER�ƃL�[�X�g���[�N�𑗐M
    'SendKeys "+{TAB}{ENTER}"    'SHIHT+TAB��ENTER�ƃL�[�X�g���[�N�𑗐M
    'SendKeys�@"%Y"      �@      'ALT+Y�ƃL�[�X�g���[�N�𑗐M

Sub SendKeys2()
    Dim objIE As Object
   'IE���J���ăt�@�C���̕ۑ�URL���J��
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://book.impress.co.jp/appended/3384/excel.zip"
    
    '3�b�x��ł���ALT+S�𑗐M
    Sleep 3000
    
    SendKeys "%S", True
    
End Sub
