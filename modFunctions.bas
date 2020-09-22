Attribute VB_Name = "modFunctions"
Function MessageBox(varText As String, varTitle As String)
On Error GoTo ERR

frmMessageBox.MP.FileName = App.Path & "\sounds\beep.wav"

frmMessageBox.lblText.Caption = varText
frmMessageBox.Caption = varTitle
frmMessageBox.Show

frmMessageBox.MP.Play

MsgBoxDone = False
Do Until MsgBoxDone = True
DoEvents
Loop

Exit Function
ERR:
MsgBox ("MessageBox Error. TEXT: " & varText & " TITLE: " & varTitle)

End Function

Function UserBox(varText As String, varTitle As String)
On Error GoTo ERR

frmInputBox.lblText.Caption = varText
frmInputBox.Caption = varTitle
frmInputBox.txtInput.Text = ""
Beep
frmInputBox.Show

MsgBoxDone = False
Do Until MsgBoxDone = True
DoEvents
Loop

varResults = frmInputBox.txtInput.Text

Exit Function
ERR:
MsgBox ("UserBox Error. TEXT: " & varText & " TITLE: " & varTitle)

End Function

Function LoadRemovableMedia()

On Error GoTo ERR

Open "A:\start.ini" For Input As #1

Line Input #1, varMediaType
Debug.Print "Media type is: " & varMediaType

If varMediaType = "EXECUTABLECD" Then
Line Input #1, varName
MkDir ("C:\Programs\" & varName)

Do While Not EOF(1)

Line Input #1, varSourceName

varNull = CopyFile("D:\" & varSourceName, "C:\Programs\" & varName & "\" & varSourceName, 1)

Loop

Open "C:\Programs\Index.dat" For Append As #1

Print #1, varName

Close #1


ElseIf varMediaType = "EXECUTABLEDISK" Then
Line Input #1, varName
MkDir ("C:\Programs\" & varName)

Do While Not EOF(1)

Line Input #1, varSourceName

varNull = CopyFile("A:\" & varSourceName, "C:\Programs\" & varName & "\" & varSourceName, 1)

Loop

Open "C:\Programs\Index.dat" For Append As #1

Print #1, varName

Close #1


ElseIf varMediaType = "DOCUMENTATION" Then
Shell ("C:\Windows\Notepad.exe D:\DOCS\index.txt")


ElseIf varMediaType = "MUSIC" Then


ElseIf varMediaType = "PICTURES" Then
Shell ("C:\Program Files\Accessories\MSPAINT.EXE")

Else
Beep
varNull = MessageBox("Could not load removable media. Reason: Unknown Media Type", "ERROR")

End If


Close #1

Exit Function
ERR:
Beep
varNull = MessageBox("Removable media index file could not be found. Make sure the floppy is in the drive.", "ERROR")

End Function

Function LoadRemovableMediaInfo()
On Error GoTo ERR

Open "A:\info.ini" For Input As #1

Line Input #1, varTitle
Line Input #1, varPublisher

Do While Not EOF(1)

Line Input #1, varBuffer
varDesc = varDesc & varBuffer & vbCrLf

Loop

varNull = MessageBox("This is the information about the removeable media." & vbCrLf & "Title: " & varTitle & vbCrLf & "Publisher: " & varPublisher & vbCrLf & "Description: " & varDesc, "Information about removable media.")

Close #1

Exit Function
ERR:
Beep
varNull = MessageBox("Could not load removeable media info. This could be because the info.ini file was not found or the disk was not in the drive.", "ERROR")

End Function

Function LoadApp(strProg As String)
On Error GoTo ERR

Select Case LCase(strProg)

Case "internet"
Shell ("C:\Programs\Internet\program.exe")

Case "email"
Shell ("C:\Programs\EMail\program.exe")

Case "fileviewer"
frmFileViewer.Show

Case "help"
Shell (App.Path & "\help\help.exe")

Case Else
varNull = MessageBox("Sorry. An error has occured. This is not your fault it's the programmers. This is probably because of a shourtcut to a program that was never finished in time for this release. You might want to check for an update to see if the problem is fixed.", "Uh-Oh!")
End Select

Exit Function
ERR:
varNull = MessageBox("Sorry. An error has occured. This is not your fault it's the programmers. This is probably because of a shourtcut to a program that was never finished in time for this release. You might want to check for an update to see if the problem is fixed.", "Uh-Oh!")

End Function

Function UserLogin()

frmLogin.Show
MsgBoxDone = False
Do Until MsgBoxDone = True
DoEvents
Loop

End Function

Function UpdateNexcon()
'On Error Resume Next
'varNull = CopyFile("A:\nexcon.exe", App.Path & "\nexcon.exe", 0)
'varNull = CopyFile("A:\pics", App.Path & "\pics", 0)
'varNull = CopyFile("A:\restricted.html", App.Path & "\restricted.html", 0)
'varNull = CopyFile("A:\programs.zip", "C:\Programs\programs.zip", 0)
'Shell ("C:\Program Files\WinRAR\WinRAR.exe C:\Programs\programs.zip")

Unload frmAbout
Unload BackgroundOptions
Unload frmConsole
Unload frmFileViewer
Unload frmLoad
Unload frmLogin
Unload frmMain
Unload frmMessageBox
Unload frmStart
Shell (App.Path & "\nexconupdater.exe")
End

End Function

Sub KeepOffTop(F As Form)
    'sets the given form Off TopMost
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    SetWindowPos F.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


Sub KeepOnTop(F As Form)
    'sets the given form On TopMost
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub Sleep(Seconds As Single, EventEnable As Boolean)
    On Error GoTo ErrHndl
    Dim OldTimer As Single
    
    OldTimer = Timer
    Do While (Timer - OldTimer) < Seconds
        If EventEnable Then DoEvents
    Loop

    Exit Sub
ErrHndl:
    ERR.Clear
End Sub


Function SystemFunction(FunctionName As String)

Select Case FunctionName

Case "#LOGVIEW"
frmMain.RTB.LoadFile (App.Path & "\log.log")
frmConsole.txtCommand.Text = frmMain.RTB.Text
frmConsole.cmdExecute.Visible = False

Case "#LOGPRINT"
frmMain.RTB.LoadFile (App.Path & "\log.log")
Printer.Print frmMain.RTB.Text
Printer.EndDoc

Case "#LOGCLEAR"
frmMain.RTB.Text = ""
frmMain.RTB.SaveFile (App.Path & "\log.log")

End Select

End Function
