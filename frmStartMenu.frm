VERSION 5.00
Begin VB.Form frmStartMenu 
   BorderStyle     =   0  'None
   Caption         =   "Start Menu"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   975
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line10 
      X1              =   840
      X2              =   4920
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image imgClose 
      Height          =   480
      Left            =   960
      Picture         =   "frmStartMenu.frx":0000
      Top             =   315
      Width           =   480
   End
   Begin VB.Label lblClose 
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   360
      Width           =   3255
   End
   Begin VB.Line Line9 
      X1              =   840
      X2              =   4920
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image imgMyFiles 
      Height          =   600
      Left            =   840
      Picture         =   "frmStartMenu.frx":0882
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   600
   End
   Begin VB.Label lblMyFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "My Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Line Line8 
      X1              =   840
      X2              =   4920
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Image imgInternet 
      Height          =   480
      Left            =   960
      Picture         =   "frmStartMenu.frx":154C
      Top             =   4440
      Width           =   480
   End
   Begin VB.Label lblInternet 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Line Line7 
      X1              =   840
      X2              =   4920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Image imgFiles 
      Height          =   480
      Left            =   960
      Picture         =   "frmStartMenu.frx":1856
      Top             =   5160
      Width           =   480
   End
   Begin VB.Label lblFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Line Line6 
      X1              =   840
      X2              =   4920
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Image imgHelp 
      Height          =   480
      Left            =   960
      Picture         =   "frmStartMenu.frx":2520
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Line Line5 
      X1              =   840
      X2              =   4920
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Image imgRun 
      Height          =   480
      Left            =   960
      Picture         =   "frmStartMenu.frx":31EA
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label lblRun 
      BackStyle       =   0  'Transparent
      Caption         =   "Run..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Line Line4 
      X1              =   840
      X2              =   4920
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Image imgShutdown 
      Height          =   480
      Left            =   960
      Picture         =   "frmStartMenu.frx":3EB4
      Top             =   7320
      Width           =   480
   End
   Begin VB.Label lblShutdown 
      BackStyle       =   0  'Transparent
      Caption         =   "Shutdown Your Computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   7320
      Width           =   3375
   End
   Begin VB.Line Line3 
      X1              =   840
      X2              =   4920
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   720
      X2              =   720
      Y1              =   7920
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   600
      X2              =   600
      Y1              =   120
      Y2              =   7680
   End
   Begin VB.Label lblSide 
      BackStyle       =   0  'Transparent
      Caption         =   "N E X C O N   S H E L L   R E P L A C E M E N T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   10815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "N E X C O N   S H E L L   R E P L A C E M E N T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   10815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   8085
      Left            =   0
      Picture         =   "frmStartMenu.frx":4B7E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5280
   End
End
Attribute VB_Name = "frmStartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_LostFocus()
Me.Hide
End Sub

Private Sub imgClose_Click()
Me.Hide
End Sub

Private Sub imgFiles_Click()
varNull = LoadApp("fileviewer")
End Sub

Private Sub imgHelp_Click()
varNull = LoadApp("help")
End Sub

Private Sub imgInternet_Click()
varNull = LoadApp("internet")
End Sub

Private Sub imgMyFiles_Click()
varNull = LoadApp("fileviewer")
End Sub

Private Sub imgRun_Click()
frmStart.Show
End Sub

Private Sub imgShutdown_Click()
Unload frmAbout
Unload BackgroundOptions
Unload frmConsole
Unload frmFileViewer
Unload frmLoad
Unload frmLogin
Unload frmMain
Unload frmMessageBox
Unload frmStart
End
Unload Me

End Sub

Private Sub lblClose_Click()

Me.Hide

End Sub

Private Sub lblFiles_Click()
varNull = LoadApp("fileviewer")
End Sub

Private Sub lblHelp_Click()
varNull = LoadApp("help")
End Sub

Private Sub lblInternet_Click()
varNull = LoadApp("Internet")
End Sub

Private Sub lblMyFiles_Click()
varNull = LoadApp("fileviewer")
End Sub

Private Sub lblRun_Click()
frmStart.Show
End Sub

Private Sub lblShutdown_Click()
Unload frmAbout
Unload frmBackgroundOptions
Unload frmConsole
Unload frmFileViewer
Unload frmLoad
Unload frmLogin
Unload frmMain
Unload frmMessageBox
Unload frmStart
End
Unload Me

End Sub
