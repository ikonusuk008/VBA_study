Attribute VB_Name = "L3_5_5_2_SendKeys3_"
Public Declare PtrSafe Function GetLastActivePopup Lib "user32" _
    (ByVal hWnd As Long) As Long

    'SendKeys "abc"              'a��b��c�ƃL�[�X�g���[�N�𑗐M
    'SendKeys "{TAB 2}{ENTER}"   'TAB��TAB��ENTER�ƃL�[�X�g���[�N�𑗐M
    'SendKeys "+{TAB}{ENTER}"    'SHIHT+TAB��ENTER�ƃL�[�X�g���[�N�𑗐M
    'SendKeys�@"%Y"      �@      'ALT+Y�ƃL�[�X�g���[�N�𑗐M

Sub SendKeys2_2()
    Dim objIE As Object
   'IE���J���ăt�@�C���̕ۑ�URL���J��
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Visible = True
    objIE.navigate "http://book.impress.co.jp/appended/3384/excel.zip"
    
    '�t�@�C�����J���_�C�A���O���\�������܂Ń��[�v
    Do While objIE.hWnd = GetLastActivePopup(objIE.hWnd)
        DoEvents
    Loop
        SendKeys "%S", True '��ALT+S�𑗐M
End Sub
