Attribute VB_Name = "modMP3"
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


Public Sub PlayMP3(filename As String)
    mciSendString "Open " & filename & " Alias MM", 0, 0, 0
    mciSendString "Play MM", 0, 0, 0
End Sub


Public Sub PauseMP3()
    mciSendString "Stop MM", 0, 0, 0
End Sub


Public Sub StopMP3()
    mciSendString "Stop MM", 0, 0, 0
    mciSendString "Close MM", 0, 0, 0
End Sub
