VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrBattStatus 
      Interval        =   60000
      Left            =   11400
      Top             =   6120
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   11280
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrEmailFade 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   11400
      Top             =   6000
   End
   Begin VB.Timer tmrStartFade 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   11400
      Top             =   5880
   End
   Begin VB.Timer tmrInetFade 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   11400
      Top             =   5760
   End
   Begin VB.Timer tmrFileFade 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   11400
      Top             =   5640
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   10920
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6534
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   255
      Left            =   11040
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":D39CC
   End
   Begin VB.Timer tmrMessage 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   11400
      Top             =   5520
   End
   Begin VB.TextBox txtMessage 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2107
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "This program made in whole by Adam Pippin. All rights reserved. Click here for the about box."
      Top             =   3960
      Width           =   8010
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   10800
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrStart 
      Interval        =   1
      Left            =   11400
      Top             =   5400
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   11400
      Top             =   5280
   End
   Begin VB.Image imgPower 
      Height          =   375
      Left            =   9705
      Picture         =   "frmMain.frx":D3A4E
      Stretch         =   -1  'True
      Top             =   8505
      Width           =   360
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   641
      X2              =   641
      Y1              =   568
      Y2              =   592
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C00000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   643
      X2              =   643
      Y1              =   569
      Y2              =   593
   End
   Begin VB.Image imgSpeaker 
      Height          =   375
      Left            =   8040
      Picture         =   "frmMain.frx":D4718
      Stretch         =   -1  'True
      ToolTipText     =   "Media Player Bar: No Media Loaded"
      Top             =   8520
      Width           =   375
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   528
      X2              =   528
      Y1              =   568
      Y2              =   592
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C00000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   530
      X2              =   530
      Y1              =   569
      Y2              =   593
   End
   Begin VB.Image imgPause 
      Height          =   255
      Left            =   9000
      Picture         =   "frmMain.frx":D53E2
      Stretch         =   -1  'True
      Top             =   8580
      Width           =   255
   End
   Begin VB.Image imgStop 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":D5544
      Stretch         =   -1  'True
      Top             =   8580
      Width           =   255
   End
   Begin VB.Image imgPlay 
      Height          =   255
      Left            =   8760
      Picture         =   "frmMain.frx":D56A6
      Stretch         =   -1  'True
      Top             =   8580
      Width           =   255
   End
   Begin VB.Image imgLoad 
      Height          =   240
      Left            =   8520
      Picture         =   "frmMain.frx":D5808
      Stretch         =   -1  'True
      Top             =   8580
      Width           =   240
   End
   Begin MediaPlayerCtl.MediaPlayer WM 
      Height          =   375
      Left            =   11400
      TabIndex        =   8
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   0   'False
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblS2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5925
      TabIndex        =   7
      Top             =   8580
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblS1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   8565
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Image imgS2 
      Height          =   360
      Left            =   5520
      Picture         =   "frmMain.frx":D5B12
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgS1 
      Height          =   360
      Left            =   3000
      Picture         =   "frmMain.frx":D67DC
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image qLock 
      Height          =   360
      Left            =   2310
      Picture         =   "frmMain.frx":D74A6
      Stretch         =   -1  'True
      ToolTipText     =   "Lock Nexcon"
      Top             =   8520
      Width           =   360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C00000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   674
      X2              =   674
      Y1              =   569
      Y2              =   593
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   672
      X2              =   672
      Y1              =   568
      Y2              =   592
   End
   Begin VB.Image qEMail 
      Height          =   360
      Left            =   1920
      Picture         =   "frmMain.frx":D8170
      Stretch         =   -1  'True
      ToolTipText     =   "E-Mail"
      Top             =   8520
      Width           =   360
   End
   Begin VB.Image qInet 
      Height          =   360
      Left            =   1440
      Picture         =   "frmMain.frx":D8E3A
      Stretch         =   -1  'True
      ToolTipText     =   "Internet"
      Top             =   8520
      Width           =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C00000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   185
      X2              =   185
      Y1              =   570
      Y2              =   594
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   183
      X2              =   183
      Y1              =   568
      Y2              =   592
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   90
      X2              =   90
      Y1              =   569
      Y2              =   593
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   88
      X2              =   88
      Y1              =   568
      Y2              =   592
   End
   Begin VB.Image PicTime 
      Height          =   480
      Left            =   10185
      Picture         =   "frmMain.frx":D9144
      Top             =   8460
      Width           =   480
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10740
      TabIndex        =   4
      Top             =   8610
      Width           =   975
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start Program"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image PicStart 
      Height          =   480
      Left            =   240
      Picture         =   "frmMain.frx":D944E
      Top             =   3120
      Width           =   480
   End
   Begin VB.Label blEMail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   735
   End
   Begin VB.Image PicEMail 
      Height          =   480
      Left            =   240
      Picture         =   "frmMain.frx":D9758
      Top             =   2160
      Width           =   480
   End
   Begin VB.Label lblInet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Internet"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.Image imgInet 
      Height          =   480
      Left            =   240
      Picture         =   "frmMain.frx":DA422
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lblFiles 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Files"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Image ImgFiles 
      Height          =   480
      Left            =   240
      Picture         =   "frmMain.frx":DA72C
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "frmMain.frx":DB3F6
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   345
   End
   Begin VB.Image ImgStart 
      Height          =   555
      Left            =   -450
      Picture         =   "frmMain.frx":DB838
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   12450
   End
   Begin VB.Image PicBack 
      Height          =   8475
      Left            =   0
      Picture         =   "frmMain.frx":1AEEFA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11985
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub blEMail_DblClick()
LoadApp ("email")

lblEMail.BackStyle = 1
lblEMail.BackColor = &HFF0000
tmrEmailFade.Enabled = True

End Sub

Private Sub Form_Load()

'Open App.Path & "\background.info" For Input As #1
'Line Input #1, varBackground
'Line Input #1, varToolbar
'Close #1
'PicBack.Picture = LoadPicture(varBackground)
'BackPic = varBackground
'ImgStart.Picture = LoadPicture(varToolbar)
'ToolPic = varToolbar

'Me.Width = User.ScreenW
'Me.Height = User.ScreenH
'PicBack.Width = User.ScreenW
'PicBack.Height = User.ScreenH


Dim DevM As DEVMODE
erg& = EnumDisplaySettings(0&, 0&, DevM)
DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
DevM.dmPelsWidth = 800 'ScreenWidth
DevM.dmPelsHeight = 600 'ScreenHeight
'DevM.dmBitsPerPel = 32 (could be 8, 16, 32 or even 4)
erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)

'KeepOnTop ()

Shell "c:\windows\system32\systray.exe"

'MsgBox (Dir("C:\Programs\"))
frmStart.cmbProg.AddItem ("Shutdown Nexcon")
'frmStart.cmbProg.AddItem ("Log Off Nexcon")
frmStart.cmbProg.AddItem ("Program Off of Removeable Media...")
'frmStart.cmbProg.AddItem ("Notepad")

If varSafeMode = True Then
Else
'Populate the start menu
Open "C:\Programs\index.dat" For Input As #1

Do While Not EOF(1)

Line Input #1, varBuffer
'varBuffer = Left(varBuffer, Len(varBuffer) - 1)
'varBuffer = Right(varBuffer, Len(varBuffer) - 1)
frmStart.cmbProg.AddItem (varBuffer)

Loop

Close #1

'Network Settings
'varComputerName = RTB.LoadFile(App.Path & "\SETTINGS\computername.set")
'varConnectTo = RTB.LoadFile(App.Path & "\SETTINGS\connectto.set")

'WS.RemoteHost = "127.0.0.1"
'WS.RemotePort = 6534
'WS.Connect

Select Case SysInfo1.BatteryStatus

Case 1
'  List1.AddItem "BatteryStatus = HIGH"
imgPower.ToolTipText = "Battery Status: High"
Case 2
'  List1.AddItem "BatteryStatus = LOW"
imgPower.ToolTipText = "Battery Status: Low"
Case 4
'  List1.AddItem "BatteryStatus = CRITICAL"
imgPower.ToolTipText = "Battery Status: Critical!"
varNull = MessageBox("Battery Power is CRITICAL. It is recommended that you shut down your computer and recharge the batteries.", "EXTREMELY Low Battery")

Case 128
'  List1.AddItem "BatteryStatus = NO BATTERY"
imgPower.ToolTipText = "No Battrery Found!"
imgPower.Visible = False
Line9.Visible = False
Line10.Visible = False

Case 255
'  List1.AddItem "BatteryStatus = UNKNOWN"
imgPower.ToolTipText = "Battery Status: Unknown"

End Select


End If

Me.Show
tmrMessage.Enabled = True

RTB.LoadFile (App.Path & "\log.log")
RTB.Text = RTB.Text & vbCrLf & "EVENT: SYSTEM START " & Time & " " & Date
RTB.SaveFile (App.Path & "\log.log")

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Open App.Path & "\background.info" For Output As #1
'Write #1, BackPic
'Write #1, ToolPic
'Close #1

Dim DevM As DEVMODE
erg& = EnumDisplaySettings(0&, 0&, DevM)
DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
DevM.dmPelsWidth = 1152 'ScreenWidth
DevM.dmPelsHeight = 864 'ScreenHeight
'DevM.dmBitsPerPel = 32 (could be 8, 16, 32 or even 4)
erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)

End Sub

Private Sub Image1_Click()
frmStartMenu.Show
End Sub

Private Sub imgFiles_Click()
varNull = LoadApp("fileviewer")
End Sub

Private Sub imgInet_Click()

varNull = LoadApp("internet")

End Sub

Private Sub imgMP3Open_Click()
CM.InitDir = "C:\"
CM.DialogTitle = "Select a MP3 to play..."
CM.Filter = "MP3's|*.mp3|All Files|*.*"
CM.ShowOpen

If CM.FileName = "" Then
Else
cmbMP3.Text = CM.FileName
End If

End Sub

Private Sub imgLoad_Click()
On Error Resume Next

CD.Filter = "Supported Files|*.mp3;*.wav"
CD.DialogTitle = "Select a MP3 to play..."
CD.ShowOpen

If CD.FileName = "" Then
Else
WM.FileName = CD.FileName
imgSpeaker.ToolTipText = "Media Player Bar: " & CD.FileName
End If

End Sub

Private Sub imgPause_Click()
On Error Resume Next
WM.Pause
End Sub

Private Sub imgPlay_Click()
On Error Resume Next

If WM.FileName = "" Then
varNull = MessageBox("No media loaded!", "Media Missing!")
Else
WM.Play
End If

End Sub

Private Sub imgS1_Click()

Select Case SM_1P

Case "fileviewer"
frmFileViewer.Show

Case Else
varNull = MessageBox("This program is not recognized!", "Weird!")
NotFound = True
End Select

If Not NotFound = True Then
lblS1.Visible = False
imgS1.Visible = False
SM_1 = False
SM_1P = ""
Else
End If

End Sub

Private Sub imgS2_Click()

Select Case SM_2P

Case "fileviewer"
frmFileViewer.Show

Case Else
varNull = MessageBox("This program is not recognized!", "Weird!")
NotFound = True
End Select

If Not NotFound = True Then
lblS2.Visible = False
imgS2.Visible = False
SM_2 = False
SM_2P = ""
Else
End If

End Sub

Private Sub imgStop_Click()
On Error Resume Next
WM.Stop
End Sub

Private Sub lblFiles_DblClick()
varNull = LoadApp("fileviewer")
lblFiles.BackStyle = 1
lblFiles.BackColor = &HFF0000
tmrFileFade.Enabled = True
End Sub

Private Sub lblInet_DblClick()

varNull = LoadApp("internet")
lblInet.BackStyle = 1
lblInet.BackColor = &HFF0000
tmrInetFade.Enabled = True

End Sub

Private Sub lblS1_Click()

Select Case SM_1P

Case "fileviewer"
frmFileViewer.Show

Case Else
varNull = MessageBox("This program is not recognized!", "Weird!")
NotFound = True
End Select

If Not NotFound = True Then
lblS1.Visible = False
imgS1.Visible = False
SM_1 = False
SM_1P = ""
Else
End If

End Sub

Private Sub lblS2_Click()

Select Case SM_2P

Case "fileviewer"
frmFileViewer.Show

Case Else
varNull = MessageBox("This program is not recognized!", "Weird!")
NotFound = True
End Select

If Not NotFound = True Then
lblS2.Visible = False
imgS2.Visible = False
SM_1 = False
SM_1P = ""
Else
End If

End Sub

Private Sub lblStart_DblClick()
frmStart.Show

lblStart.BackStyle = 1
lblStart.BackColor = &HFF0000
tmrStartFade.Enabled = True

End Sub

Private Sub PicBack_DblClick()

frmBackgroundOptions.Show

End Sub

Private Sub PicEMail_Click()
varNull = LoadApp("email")
End Sub

Private Sub PicStart_Click()
frmStart.Show
End Sub

Private Sub qEMail_Click()
varNull = LoadApp("email")
End Sub

Private Sub qInet_Click()
varNull = LoadApp("internet")
End Sub

Private Sub qLock_Click()
On Error Resume Next

varNull = UserBox("Please enter the password you want to use to unlock Nexcon:", "Password")
If varResults = "CANCEL" Then
ElseIf varResults = "" Then
varNull = MessageBox("You need to enter a password!", "Oops!")
Else
strPassword = varResults
WM.Pause
Me.Hide
frmLock.Show
End If

End Sub

Private Sub tmrBattStatus_Timer()

Select Case SysInfo1.BatteryStatus

Case 1
'  List1.AddItem "BatteryStatus = HIGH"
imgPower.ToolTipText = "Battery Status: High"
Case 2
'  List1.AddItem "BatteryStatus = LOW"
imgPower.ToolTipText = "Battery Status: Low"
Case 4
'  List1.AddItem "BatteryStatus = CRITICAL"
imgPower.ToolTipText = "Battery Status: Critical!"
varNull = MessageBox("Battery Power is CRITICAL. It is recommended that you shut down your computer and recharge the batteries.", "EXTREMELY Low Battery")

Case 128
'  List1.AddItem "BatteryStatus = NO BATTERY"
imgPower.ToolTipText = "No Battrery Found!"
imgPower.Visible = False
Line9.Visible = False
Line10.Visible = False

Case 255
'  List1.AddItem "BatteryStatus = UNKNOWN"
imgPower.ToolTipText = "Battery Status: Unknown"

End Select

End Sub

Private Sub tmrEmailFade_Timer()
Select Case lblEMail.BackColor

Case &HFF0000
lblEMail.BackColor = &HC00000

Case &HC00000
lblEMail.BackColor = &H800000

Case &H800000
lblEMail.BackStyle = 0

End Select

End Sub

Private Sub tmrFileFade_Timer()

Select Case lblFiles.BackColor

Case &HFF0000
lblFiles.BackColor = &HC00000

Case &HC00000
lblFiles.BackColor = &H800000

Case &H800000
lblFiles.BackStyle = 0

End Select

End Sub

Private Sub tmrInetFade_Timer()
Select Case lblInet.BackColor

Case &HFF0000
lblInet.BackColor = &HC00000

Case &HC00000
lblInet.BackColor = &H800000

Case &H800000
lblInet.BackStyle = 0

End Select

End Sub

Private Sub tmrMessage_Timer()
txtMessage.Visible = False
tmrMessage.Enabled = False
End Sub

Private Sub tmrStart_Timer()
lblTime.Caption = Left(Time, 5) & " " & Right(Time, 2)
lblTime.ToolTipText = Date
tmrStart.Enabled = False
End Sub

Private Sub tmrStartFade_Timer()
Select Case lblStart.BackColor

Case &HFF0000
lblStart.BackColor = &HC00000

Case &HC00000
lblStart.BackColor = &H800000

Case &H800000
lblStart.BackStyle = 0

End Select

End Sub

Private Sub tmrTime_Timer()

lblTime.Caption = Left(Time, 5) & " " & Right(Time, 2)
lblTime.ToolTipText = Date

End Sub

Private Sub txtMessage_Click()

frmAbout.Show

End Sub

Private Sub WS_Connect()

Me.Hide
varNull = UserLogin()

End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)

'Select Case UCase(WS.GetData)

'Case "USERNAME"
If varUsername = "" Then
Else
WS.SendData varUsername
End If

'Case "PASSWORD"
If strPassword = "" Then
Else
WS.SendData strPassword
End If

'Case "OK"
frmMain.Show
varNull = MessageBox("You have been successfully logged in!", "Success!")

'Case "NO"
varNull = MessageBox("Your username and password were invalid. Exiting now.", "Good-Bye!")
End

'Case Else
frmMain.Show
varNull = MessageBox("An unknown command was received from the server.", "ERROR")

'End Select

End Sub

Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
varNull = MessageBox("There has been an error in the networking.", "Networking Error")
End Sub
