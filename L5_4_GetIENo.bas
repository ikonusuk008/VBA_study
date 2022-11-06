Attribute VB_Name = "L5_4_GetIENo"
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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
 Sub GetGN()
    
    Dim colobj As Object
    Dim obj As Object

    Dim colSh As Object    '�N������ShellWindow�ꎮ���i�[����ϐ�
    Dim win As Object   'ShellWindow���i�[����ϐ�
    Dim strTemp As String   'IE�̃^�C�g�����i�[����ϐ�
    Dim objie As Object     '�ړI��IE���i�[����ϐ�
    Set colSh = CreateObject("Shell.Application")  '���݊J���Ă��� IE �� �G�N�X�v���[����colSh�Ɋi�[
    'ColSh����Window��1�����o��
    For Each win In colSh.Windows
        strTemp = ""
        '�^�C�g�����擾�ł��Ȃ��ꍇ���������p��
        On Error Resume Next
        strTemp = win.Document.Title
        On Error GoTo 0
        '�^�C�g���o�[�Ɂh�H�׃��O�h���܂܂�邩����
        If InStr(strTemp, "�O�����E���X�g�����\��T�C�g") > 0 Then
            '�ϐ�ie�Ɏ擾����win���i�[
            Set objie = win
            '���[�v�𔲂���
            Exit For
        End If
    Next
    
    If objie Is Nothing Then
        MsgBox "�T���Ă���IE�͂���܂���ł���"
        Exit Sub
    Else   '�^�C�g����\������
        MsgBox objie.Document.Title & "������܂���"
        Call showForeground(objie)
    End If
    
    '�ꗗ�擾
    Dim r As Long
    Sheets("MAIN").Select
    Cells.ClearContents
    
    Cells(1, 1) = "NO"
    Cells(1, 2) = "�X��"
    Cells(1, 3) = "URL"
    Cells(1, 4) = "�_��"
    Cells(1, 5) = "���R�~����"
    Cells(1, 6) = "��̗\�Z"
    Cells(1, 7) = "���̗\�Z"

    Dim i As Long
    Dim i2 As Long
    r = 1
Start:
    Cells(r, 1).Select
    With objie.Document
        '�y�[�W�㕔�𑖍��ΏۊO�Ƃ��邱�ƂŁA���X�Ɍ��肷��
        For i = 700 To .all.Length - 1
        'DIV,DIV,DIV,P,A�Ƃ������т��o������ӏ���T��
            If .all(i).TagName = "LI" Then
             If .all(i + 1).TagName = "DIV" Then
              If .all(i + 2).TagName = "DIV" Then
               If .all(i + 3).TagName = "P" Then
                If .all(i + 4).TagName = "A" And .all(i + 4).innertext <> "" _
                    And .all(i + 4).innertext <> "���X�g�����̐V�K�o�^�y�[�W" Then
                    '�X��������Ɣ��f
                    r = r + 1
                    Cells(r, 1) = r - 1
                    Cells(r, 2) = .all(i + 4).innertext '�X��
                    Cells(r, 3) = .all(i + 4).href 'URL
                    '�ȍ~�̃^�O����A�ڈ��SPAN�^�O�𑖍�
                    For i2 = i To .all.Length - 1
                        If .all(i2).TagName = "SPAN" Then
                            '��̓_��������΁A���̃^�O���N�_�ɒl���擾
                            If InStr(.all(i2).innertext, "��̗\�Z") > 0 Then
                                Cells(r, 4) = .all(i2 - 7).innertext  '"�_��"
                                '�_�����Ȃ��ꍇ�̏C��
                                If InStr(Cells(r, 4), "��") > 0 Then
                                    Cells(r, 4) = "-"
                                End If
                                Cells(r, 5) = .all(i2 - 3).innertext '"���R�~����"
                                Cells(r, 6) = .all(i2 + 1).innertext   '"��̗\�Z"
                                Cells(r, 7) = .all(i2 + 5).innertext  '"���̗\�Z"
                                Exit For
                            End If
                        End If
                    Next i2
                End If
               End If
              End If
             End If
             
            '�������邩����
            ElseIf .all(i).TagName = "A" Then
                If .all(i).innertext = "����20��" Then
                    '���̃y�[�W�֑J��
                    .all(i).Click
                    Call waitNavigation(objie)
                    GoTo Start
                ElseIf .all(i).innertext = "���X�g�����̐V�K�o�^�y�[�W" Then
                    '�ŏI�y�[�W�Ɣ��f
                    Exit For
                End If
            End If
        Next i
    End With
    
    MsgBox "�ꗗ�\���쐬���܂���"

End Sub


'�y�[�W���̕����񂩂�N������IE���擾����
Function SearchIE(strTarget As String) As Object

    Dim colSh As Object    '�N������ShellWindow�ꎮ���i�[����ϐ�
    Dim win As Object   'ShellWindow���i�[����ϐ�
    Set colSh = CreateObject("Shell.Application")  '���݊J���Ă��� IE �� �G�N�X�v���[����colSh�Ɋi�[
    'ColSh����Window��1�����o��
    For Each win In colSh.Windows
        strTemp = ""
        '�y�[�W�̕����񂪎擾�ł��Ȃ��ꍇ���������p��
        On Error Resume Next
        strTemp = win.Document.body.innertext
        On Error GoTo 0
        If InStr(strTemp, strTarget) > 0 Then
            '�ϐ�ie�Ɏ擾����win���i�[
            Set SearchIE = win
            '���[�v�𔲂���
            Exit For
        End If
    Next
    If SearchIE Is Nothing Then
        MsgBox "�T���Ă���IE�͂���܂���ł���"
    Else   '�^�C�g����\������
        MsgBox "�T���Ă���IE������܂���"
    End If

End Function

'�w�肳�ꂽ�E�B���h�E���őO�ʉ�����
Sub showForeground(objie As Object)

    '�ŏ�������Ă���ꍇ�͌��̑傫���ɖ߂�(9=RESTORE:�ŏ����O�̏��)
    If IsIconic(objie.hWnd) Then
        ShowWindowAsync objie.hWnd, &H9
    End If
    '�őO�ʂɕ\��
    SetForegroundWindow (objie.hWnd)

End Sub

Sub SearchGN()
    
    Dim IE As Object
    Dim i As Long
    
    'IE���J���đ���Ώۉ�ʂ֑J��
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    IE.Navigate2 "http://tabelog.com/"
    Call waitNavigation(IE)

End Sub

Sub waitNavigation(IE As Object)
    Do While IE.Busy Or (IE.readyState <> 4 And IE.readyState <> 3)
        DoEvents
    Loop

End Sub

Sub MakeIchiran()

    Dim objie As Object
    Set objie = SearchIE("�H�׃��O")
    Call Ichiran_Make(objie)
    MsgBox "�ꗗ���쐬���܂���"

End Sub

Sub Ichiran_Make(objie As Object)
    Dim n As Long
    Dim objTAG As Object
    Application.ScreenUpdating = False
    On Error Resume Next
    n = 0
    
    For Each objTAG In objie.Document.all
        n = n + 1
        Cells(n + 2, 1) = objie.Name
        Cells(n + 2, 4) = "'" & TypeName(objTAG) 'TypeName�ŃI�u�W�F�N�g�̃^�C�v��\��
        Cells(n + 2, 5) = "'" & objTAG.TagName   '�^�O�̖��O
        Cells(n + 2, 6) = n
        Cells(n + 2, 7) = objTAG.Name
        Cells(n + 2, 8) = "'" & Left(objTAG.innertext, 256)
        Cells(n + 2, 9) = "'" & Left(objTAG.InnerHTML, 256)
        Cells(n + 2, 10) = "'" & Left(objTAG.OuterHTML, 256)
    Next
            
End Sub


