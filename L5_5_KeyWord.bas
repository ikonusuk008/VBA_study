Attribute VB_Name = "L5_5_KeyWord"

Option Explicit

Sub KeyWord()
    
    Dim i As Long  '�J�E���^�[�ϐ�
    Dim r As Long  '�Z���s�ϐ�
        r = 5      '���ʂ��������ރZ���J�n�s
    Dim colURL As New Collection    'URL�R���N�V����
    Dim a As Variant    '�R���N�V�����v�f�i�[�p�ϐ�
    
    Dim StrURL As String    '�J�nURL
    Dim StrDmn As String    '�h���C���w��
    Dim StrWord(1 To 5) As String   '�L�[���[�h�i�[�p�̔z��ϐ�
    Dim StrTitle As String  '�y�[�W�^�C�g��
    Dim StrText As String   '�y�[�W�{��
    '�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
    '�ݒ�̎擾
    '�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
    StrURL = Cells(4, 3)    '�J�nURL���i�[
    StrDmn = Cells(5, 3)    '�h���C�����i�[
    For i = 1 To 5
        StrWord(i) = Cells(5 + i, 3)  '�L�[���[�h�����ɔz��Ɋi�[
    Next i
    Range(Cells(5, 5), Cells(Rows.Count, 12)).ClearContents '���ʗ����N���A�[
    
    '�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
    'IE�I�u�W�F�N�g�̐ݒ�A�w��y�[�W���J��
    '�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
    Dim objIE As Object
    Set objIE = CreateObject("InternetExplorer.application")
    objIE.Visible = True
    objIE.navigate StrURL
    
    Do While objIE.Busy Or objIE.readyState <> 4  ':READYSTATE_COMPLETE
        DoEvents
    Loop
    
    '�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
    '�����N��URL�̎擾,�R���N�V�����Ɋi�[
    '�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
    colURL.Add StrURL, StrURL   '�X�^�[�g�y�[�W��URL���R���N�V�����Ɋi�[
    If objIE.document.Links.Length > 0 Then     '�����N������ꍇ��
        For i = 0 To objIE.document.Links.Length - 1  '�����N�̂���I�u�W�F�N�g�����Ɋm�F
            a = objIE.document.Links(i).href    '�����N����擾
            If StrDmn = "" Or InStr(a, StrDmn) > 0 Then '�h���C���w�肪�Ȃ��A�܂��́A�w�肵���h���C���ɊY������Ȃ�
                If InStr(a, "@") = 0 Then   '���[���A�h���X������
                    On Error Resume Next
                    colURL.Add a, a     '�����N��URL���R���N�V�����ɒǉ�
                    On Error GoTo 0
                End If
            End If
        Next i
    End If

    '�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
    '�����N����J���L�[���[�h�̗L�����`�F�b�N
    '�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|�|
    For Each a In colURL    '�����N��URL�����Ɏ��o��
        objIE.navigate a   '�����N����J��
        
        Do While objIE.Busy Or objIE.readyState <> 4 ':READYSTATE_COMPLETE
            DoEvents
        Loop
        
        StrTitle = ""
        StrText = ""
        On Error Resume Next
            StrTitle = objIE.document.Title   '�^�C�g�����擾
            StrText = objIE.document.body.innertext     '�y�[�W�{�����擾
        On Error GoTo 0
            
        Cells(r, 5) = StrTitle              '�^�C�g�����Z���ɏ�������
        Cells(r, 6) = a                     'URL���Z���ɏ�������
        
        For i = 1 To 5      '�L�[���[�h�̐��������[�v
            Cells(r, 6 + i) = "�|"  '��U�A���ʂ�"�|"��
            If StrWord(i) <> "" Then    '�L�[���[�h�w�肪��������
                If InStr(StrText, StrWord(i)) > 0 Then   '�w��L�[���[�h�������
                    Cells(r, 6 + i) = "�Z"  '���ʂ��Z���ɏ�������
                End If
            End If
        Next i
        r = r + 1
    Next
    
    MsgBox "�������܂���"

End Sub
