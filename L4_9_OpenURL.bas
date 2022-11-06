Attribute VB_Name = "L4_9_OpenURL"
Option Explicit

Declare PtrSafe Sub Sleep Lib "KERNEL32.dll" (ByVal dwMilliseconds As Long)

Sub OpenURL()
    
    Dim IE As InternetExplorer
    Dim Doc As HTMLDocument
    Dim ObjTag As Object
    Dim i As Long
    
    Const Src1 As String = "button_01.png"
    Const Src2 As String = "button_02.png"
    
    
    'IE���J���đ���Ώۉ�ʂ֑J��
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    IE.navigate "http://book.impress.co.jp/appended/3384/4-9.html"
    Call waitNavigation(IE)
    Set Doc = IE.document
    
    '�_�ł�10��J��Ԃ�
    For i = 1 To 10
        For Each ObjTag In Doc.getElementsByTagName("INPUT")
            With ObjTag
                'src�������Ȃ��ꍇ�̃G���[�������������͌p��
                On Error Resume Next
                '�摜����v������摜��ύX
                If InStr(.src, Src1) > 0 Then
                    .src = Src2
                    '0.2�b��~��A�摜�����ɖ߂��A�ēx0.2�b��~
                    Sleep 200
                    .src = Src1
                    Sleep 200
                    Exit For
                End If
                On Error GoTo 0
            End With
        Next
    Next i

End Sub

Sub waitNavigation(IE As Object)
    Do While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    Do While IE.document.readyState <> "complete"
        DoEvents
    Loop
End Sub


