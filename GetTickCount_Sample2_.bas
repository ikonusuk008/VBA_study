Attribute VB_Name = "GetTickCount_Sample2_"
Sub GetTickCount_Sample2()
    Dim StartTime As Long
    StartTime = GetTickCount
    Do While GetTickCount - StartTime < 5000
        DoEvents
    Loop
    
    MsgBox "5秒経過しました"
    
End Sub
