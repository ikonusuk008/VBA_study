Attribute VB_Name = "Bunkai"
Option Explicit

Sub Bunkai()

    Dim r As Long
    Dim appWord As Object, objDoc As Object, objWord As Object
    Set appWord = CreateObject("Word.Application")  'Word�A�v���P�[�V�����̋N��
    Set objDoc = appWord.Documents.Add              '�V�K�����I�u�W�F�N�g�̍쐬
    
    Range(Cells(8, 2), Cells(Rows.Count, 2)).ClearContents  '�P��Z���̃N���A�[
    objDoc.Range.Text = Range("B5")   '�����I�u�W�F�N�g�ɕ��͂��i�[
    r = 8
    For Each objWord In objDoc.Words    '���ɒP������o��
         Cells(r, 2) = objWord         '�Z���ɏ�������
         r = r + 1
    Next
    
    appWord.Quit SaveChanges:=0 'wdDoNotSaveChanges  'Word�̏I��
    Set objDoc = Nothing    'Word�����I�u�W�F�N�g�̃N���A�[
    Set appWord = Nothing   'Word�A�v���P�[�V�����̃N���A�[

End Sub
