Attribute VB_Name = "L4_10_GetTable1"
Option Explicit

'���o�����ڂ�T���ăe�[�u���̃e�L�X�g���擾����
Sub GetTable1()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    
    'IE���J���đ���Ώۉ�ʂ֑J��
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table1.html"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    For i = 0 To Doc.all.Length - 1 '�h�L�������g�̍\���^�O���������
        If Doc.all(i).tagName = "TH" Then   'TH�^�O�Ȃ�
            If Doc.all(i).innerText = "�������[" Then  '�e�L�X�g���u�������[�v�Ȃ�
                MsgBox Doc.all(i + 1).innerText '���̎��̃^�O�̃e�L�X�g���擾���\��
                Exit For
            End If
        End If
    Next i

End Sub

'���o�����ڂ�T���ăe�[�u���̃e�L�X�g���擾����i���o�����ڂ�������o�ꂷ��ꍇ�j
Sub GetTable2()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    Dim n As Long
    
    'IE���J���đ���Ώۉ�ʂ֑J��
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table2.html"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    For i = 0 To Doc.all.Length - 1 '�h�L�������g�̍\���^�O���������
        If Doc.all(i).tagName = "TH" Then   'TH�^�O�Ȃ�
            If Doc.all(i).innerText = "�������[" Then  '�e�L�X�g���u�������[�v�Ȃ�
                n = n + 1      '�o��񐔂��J�E���g
                If n = 2 Then '2��ڂ̓o��Ȃ�
                    MsgBox Doc.all(i + 1).innerText '���̎��̃^�O�̃e�L�X�g���擾���\��
                    Exit For
                End If
            End If
        End If
    Next i

End Sub

Public Sub waitNavigation(ie As InternetExplorer)
    Do While ie.Busy Or ie.readyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    Do While ie.document.readyState <> "complete"
        DoEvents
    Loop
End Sub

Sub GetTable3()
    Dim ie As InternetExplorer
    
    'IE���J���đ���Ώۉ�ʂ֑J��
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Call waitNavigation(ie)
    
    Call MakeList(ie)

End Sub

Sub MakeList(ObjIE As InternetExplorer)
    Dim n As Long   '�^�O�̒ʂ��ԍ�
    Dim r As Long   'TD,TH�^�O�̒ʂ��ԍ�
    Dim Doc As HTMLDocument
    Dim ObjTag As Object '�^�O�i�[�p
    n = 0
    r = 0
    Sheets("Sheet1").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/�W��"
    Set Doc = ObjIE.document
    
    For n = 0 To Doc.all.Length - 1
        With Doc.all(n)
            If .tagName = "TD" Or .tagName = "TH" Then
                r = r + 1
                Cells(r + 1, 1) = .tagName    '�^�O�̖��O
                Cells(r + 1, 2) = n          '�^�O�̒ʂ��ԍ�
                Cells(r + 1, 3) = r            'TD,TH�^�O�̒ʂ��ԍ�
                Cells(r + 1, 4) = .innerText  '�e�L�X�g
            End If
        End With
    Next
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
End Sub

Sub GetTable4()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    Dim n As Long
    
    'IE���J���đ���Ώۉ�ʂ֑J��
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    Sheets("Sheet3").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/�W��"
    
    For i = 537 To 855 '�h�L�������g�̍\���^�O���������
        If Doc.all(i).tagName = "TD" Then 'TD�^�O�Ȃ�
            n = n + 1
            Cells(Int((n - 1) / 16) + 1, (n - 1) Mod 16 + 1) = Doc.all(i).innerText
                '�s��TD�^�O�̒ʂ��ԍ���16�Ŋ�������+1�Ƃ��邱�Ƃ�16�^�O���Ƃɉ��s
                'TD�^�O�̒ʂ��ԍ���16�Ŋ������]��𗘗p���ė�ԍ����w��
        End If
    Next i
    
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

End Sub

Sub GetTable5()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    Dim n As Long
    Dim StartTag As Long
    Dim FinishTag As Long
    
    'IE���J���đ���Ώۉ�ʂ֑J��
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    Sheets("Sheet3").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/�W��"
    
    '�e�[�u���J�n�^�O�̎擾
    For i = 0 To Doc.all.Length - 1  '�h�L�������g�̍\���^�O���������
        If Doc.all(i).tagName = "TH" Then 'TH�^�O�Ȃ�
            If Doc.all(i).innerText = "�t��" Then
                StartTag = i
                Exit For
            End If
        End If
    Next i

    '�e�[�u���C���^�O�̎擾
    For i = StartTag To Doc.all.Length - 1  '�h�L�������g�̍\���^�O���������
        If Doc.all(i).tagName = "TH" Then 'TH�^�O�Ȃ�
            If Doc.all(i).innerText = "���[�J�[" Then
                FinishTag = i
                Exit For
            End If
        End If
    Next i
    
    For i = StartTag To FinishTag '�h�L�������g�̍\���^�O���������
        If Doc.all(i).tagName = "TD" Then 'TD�^�O�Ȃ�
            n = n + 1
            Cells(Int((n - 1) / 16) + 1, (n - 1) Mod 16 + 1) = Doc.all(i).innerText
        End If
    Next i
    
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

End Sub

Sub GetTable6()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTD As Object
    Dim ObjTag As Object
    Dim n As Long
    
    'IE���J���đ���Ώۉ�ʂ֑J��
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    Sheets("Sheet4").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/�W��"
    
    
    'TD�^�O�̂ݎ擾���ꗗ��
    Set ObjTD = Doc.getElementsByTagName("TD")
    For Each ObjTag In ObjTD
        n = n + 1
        Cells(n, 1) = n
        Cells(n, 2) = ObjTag.tagName
        Cells(n, 3) = ObjTag.innerText
    Next ObjTag
    
End Sub

Sub GetTable7()
    Dim ie As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTables As Object
    Dim ObjTag As Object
    
    Dim r As Long
    Dim c As Long
    Dim i As Long
    
    'IE���J���đ���Ώۉ�ʂ֑J��
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.navigate "http://xlsg.net/vbaie/table/table3.html" '"http://kakaku.com/pc/note-pc/se_15/"
    Call waitNavigation(ie)
    Set Doc = ie.document
    
    Sheets("Sheet5").Select
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/�W��"
    
    r = 1
    c = 1
    
    '�e�[�u���J�n�^�O�̎擾
    For i = 0 To Doc.all.Length - 1  '�h�L�������g�̍\���^�O���������
        If Doc.all(i).tagName = "TH" Or Doc.all(i).tagName = "TD" Then
            'TD�ATH�^�O�Ȃ����E�Ƀe�L�X�g���Z���ɏ�������
            c = c + 1
            Cells(r, c) = Doc.all(i).innerText
        ElseIf Doc.all(i).tagName = "TR" Then
            'TR�^�O�Ȃ���s���A1��ڂɖ߂�
            r = r + 1
            c = 1
        End If
    Next i
    
End Sub
