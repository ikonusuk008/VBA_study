Attribute VB_Name = "L3_4_shell"
'�����I�ɍőO�ʂɂ�����
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'�ŏ�������Ă��邩���ׂ�
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'���̑傫���ɖ߂�API
Private Declare PtrSafe Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long


'�N������Shell�̐����擾�E�\������
    
Sub GetIENo()

    Dim colSh As Object    '�N������Shell���i�[����ϐ�
    Set colSh = CreateObject("Shell.Application")  '���݊J���Ă��� IE �� �G�N�X�v���[����colSh�Ɋi�[
    MsgBox colSh.Windows.Count

End Sub

       
'�^�C�g������N������IE���擾����
Sub SearchIE1()

    Dim colSh As Object    '�N������ShellWindow�ꎮ���i�[����ϐ�
    Dim win As Object   'ShellWindow���i�[����ϐ�
    Dim strTemp As String   'IE�̃^�C�g�����i�[����ϐ�
    Dim objIE As Object     '�ړI��IE���i�[����ϐ�
    Set colSh = CreateObject("Shell.Application")  '���݊J���Ă��� IE �� �G�N�X�v���[����colSh�Ɋi�[
    'ColSh����Window��1�����o��
    For Each win In colSh.Windows
        'HTMLDocument��������
        If TypeName(win.document) = "HTMLDocument" Then
            '�^�C�g���o�[��PC Watch���܂܂�邩����
            If InStr(win.document.Title, "PC Watch") > 0 Then
                '�ϐ�ie�Ɏ擾����win���i�[
                Set objIE = win
                '���[�v�𔲂���
                Exit For
            End If
        End If
    Next
    
    If objIE Is Nothing Then
        MsgBox "�T���Ă���IE�͂���܂���ł���"
    Else   '�^�C�g����\������
        MsgBox objIE.document.Title & "������܂���"
    End If

End Sub

       
'�^�C�g������N������IE���擾����
Sub SearchIE2()

    Dim colSh As Object    '�N������ShellWindow�ꎮ���i�[����ϐ�
    Dim win As Object   'ShellWindow���i�[����ϐ�
    Dim strTemp As String   'IE�̃^�C�g�����i�[����ϐ�
    Dim objIE As Object     '�ړI��IE���i�[����ϐ�
    Set colSh = CreateObject("Shell.Application")  '���݊J���Ă��� IE �� �G�N�X�v���[����colSh�Ɋi�[
    'ColSh����Window��1�����o��
    For Each win In colSh.Windows
        strTemp = ""
        '�^�C�g�����擾�ł��Ȃ��ꍇ���������p��
        On Error Resume Next
        strTemp = win.document.Title
        On Error GoTo 0
        '�^�C�g���o�[��PC Watch���܂܂�邩����
        If InStr(strTemp, "PC Watch") > 0 Then
            '�ϐ�ie�Ɏ擾����win���i�[
            Set objIE = win
            '���[�v�𔲂���
            Exit For
        End If
    Next
    
    If objIE Is Nothing Then
        MsgBox "�T���Ă���IE�͂���܂���ł���"
    Else   '�^�C�g����\������
        MsgBox objIE.document.Title & "������܂���"
    End If

End Sub


'�y�[�W���̕����񂩂�N������IE���擾����
Sub SearchIE3()

    Dim colSh As Object    '�N������ShellWindow�ꎮ���i�[����ϐ�
    Dim win As Object   'ShellWindow���i�[����ϐ�
    Dim strTemp As String   '�y�[�W�̕�������i�[����ϐ�
    Dim objIE As Object     '�ړI��IE���i�[����ϐ�
    Set colSh = CreateObject("Shell.Application")  '���݊J���Ă��� IE �� �G�N�X�v���[����colSh�Ɋi�[
    'ColSh����Window��1�����o��
    For Each win In colSh.Windows
        strTemp = ""
        '�y�[�W�̕����񂪎擾�ł��Ȃ��ꍇ���������p��
        On Error Resume Next
        strTemp = win.document.body.innertext
        On Error GoTo 0
        '�y�[�W��ɕ�����u�A�b�v�f�[�g���v�����݂��邩����
        If InStr(strTemp, "�A�b�v�f�[�g���") > 0 Then
            '�ϐ�ie�Ɏ擾����win���i�[
            Set objIE = win
            '���[�v�𔲂���
            Exit For
        End If
    Next
    If objIE Is Nothing Then
        MsgBox "�T���Ă���IE�͂���܂���ł���"
    Else   '�^�C�g����\������
        MsgBox "�T���Ă���IE������܂���"
    End If

End Sub


'�w�肳�ꂽ�E�B���h�E���őO�ʉ�����
Sub showForeground(objIE As Object)

    '�ŏ�������Ă���ꍇ�͌��̑傫���ɖ߂�(9=RESTORE:�ŏ����O�̏��)
    If IsIconic(objIE.hWnd) Then
        ShowWindowAsync objIE.hWnd, &H9
    End If
    '�őO�ʂɕ\��
    SetForegroundWindow (objIE.hWnd)

End Sub

'�y�[�W���̕����񂩂�N������IE���擾���A�őO�ʉ�����
Sub SearchIE4()

    Dim colSh As Object    '�N������ShellWindow�ꎮ���i�[����ϐ�
    Dim win As Object   'ShellWindow���i�[����ϐ�
    Dim strTemp As String   '�y�[�W�̕�������i�[����ϐ�
    Dim objIE As Object     '�ړI��IE���i�[����ϐ�
    Set colSh = CreateObject("Shell.Application")  '���݊J���Ă��� IE �� �G�N�X�v���[����colSh�Ɋi�[
    'ColSh����Window��1�����o��
    For Each win In colSh.Windows
        strTemp = ""
        '�y�[�W�̕����񂪎擾�ł��Ȃ��ꍇ���������p��
        On Error Resume Next
        strTemp = win.document.body.innertext
        On Error GoTo 0
        '�y�[�W��Ɏw�肵�������񂪑��݂��邩����
        If InStr(strTemp, Range("B5").Value) > 0 Then
            '�ϐ�ie�Ɏ擾����win���i�[
            Set objIE = win
            '���[�v�𔲂���
            Exit For
        End If
    Next
    If objIE Is Nothing Then
        MsgBox "�T���Ă���IE�͂���܂���ł���"
    Else   '�^�C�g����\������
        MsgBox "�T���Ă���IE������܂���"
        Call showForeground(objIE)
    End If

End Sub

