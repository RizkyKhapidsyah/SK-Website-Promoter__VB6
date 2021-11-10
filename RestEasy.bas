Attribute VB_Name = "RestEasy"
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Function WaitASec(ByVal TimeToWait As Long) 'Time In seconds
Dim EndTime As Long
        
        EndTime = GetTickCount + TimeToWait * 1000 '* 1000 because you give seconds and GetTickCount uses Milliseconds


    Do Until GetTickCount > EndTime


        DoEvents

      If frmMain.Prg.Value >= 99 Then
        frmMain.Prg.Value = 0
        Else
        frmMain.Prg.Value = frmMain.Prg.Value + 0.05
        End If
    Loop
frmMain.Prg.Value = 0
End Function

'Prg
