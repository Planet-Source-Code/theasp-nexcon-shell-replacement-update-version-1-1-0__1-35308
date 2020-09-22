VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLoad 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   0  'None
   ClientHeight    =   7380
   ClientLeft      =   180
   ClientTop       =   -180
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   492
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   629
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox RTB 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmLoad.frx":0000
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   360
      Top             =   6840
   End
   Begin VB.ListBox lstLoad 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   4905
      ItemData        =   "frmLoad.frx":0082
      Left            =   120
      List            =   "frmLoad.frx":0084
      TabIndex        =   2
      Top             =   1800
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8880
      Picture         =   "frmLoad.frx":0086
      Top             =   6840
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Nexcon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Nexcon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   855
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   9255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varSafeMode As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 126 Then

varSafeMode = True

Else
End If

End Sub

Private Sub Form_Load()

Dim DevM As DEVMODE
erg& = EnumDisplaySettings(0&, 0&, DevM)
DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
DevM.dmPelsWidth = 640 'ScreenWidth
DevM.dmPelsHeight = 480 'ScreenHeight
'DevM.dmBitsPerPel = 32 (could be 8, 16, 32 or even 4)
erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)

lstLoad.AddItem ("Press ~ now to enter safe mode.")
tmrWait.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Dim DevM As DEVMODE
'erg& = EnumDisplaySettings(0&, 0&, DevM)
'DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
'DevM.dmPelsWidth = 1152 'ScreenWidth
'DevM.dmPelsHeight = 864 'ScreenHeight
'DevM.dmBitsPerPel = 32 (could be 8, 16, 32 or even 4)
'erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)


End Sub

Private Sub lstLoad_KeyPress(KeyAscii As Integer)

If KeyAscii = 126 Then

varSafeMode = True

Else
End If

End Sub

Private Sub tmrWait_Timer()

tmrWait.Enabled = False

If varSafeMode = False Then
lstLoad.AddItem ("Loading Core Settings...")
'Not in safe mode, load up normally.
RTB.LoadFile (App.Path & "\settings\adminpass.set")
strAdminPass = RTB.Text
RTB.Text = "***PROTECTED***"
lstLoad.AddItem ("-Admin Password")


'Check for corrupt files
lstLoad.AddItem ("Checking for missing system files...")
If Not InStr(Dir(App.Path & "\"), "nexcon.exe" And "restricted.html") Then
'Files missing, notify user.
lstLoad.AddItem ("Your system is missing one of the vital Nexcon files. Please reinstall nexcon and try again.")
Exit Sub
Else
End If

Else

End If

frmMain.Show

End Sub
