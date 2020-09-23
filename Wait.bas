Attribute VB_Name = "modWait"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Wait(Miliseconds As Single)
    Dim TempTime As Long
    Dim CurTime As Long
    TempTime = GetTickCount()
    CurTime = TempTime
    Do While Miliseconds > CurTime - TempTime
        DoEvents
        CurTime = GetTickCount()
    Loop
End Sub


