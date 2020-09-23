Attribute VB_Name = "modWav"
Declare Function sndPlaysound Lib "winmm" Alias "sndPlaySoundA" _
  (ByVal soundfilename As String, ByVal flags As Long) As Long
  Const SND_ASYNC = &H1
  
Private Sub PlayWav(sFile As String)
    If Dir(sFile$) <> "" Then Call sndPlaysound(sFile, SND_ASYNC)
End Sub

Public Sub Play(ArrowSound As String)
    PlayWav App.Path & "\" & ArrowSound & ".wav"
End Sub
