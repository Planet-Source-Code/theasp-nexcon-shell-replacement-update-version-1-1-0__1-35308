VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBackgroundOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options..."
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "frmBackgroundOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowseT 
      Caption         =   "Browser for Toolbar"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "Browse For Background"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   0
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgStartPreview 
      Height          =   165
      Left            =   1680
      Picture         =   "frmBackgroundOptions.frx":0CCA
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2280
   End
   Begin VB.Image imgPreview 
      Height          =   1695
      Left            =   1680
      Picture         =   "frmBackgroundOptions.frx":D438C
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Image imgMonitor 
      Height          =   1935
      Left            =   1560
      Picture         =   "frmBackgroundOptions.frx":1A7A4E
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Image imgBack 
      Height          =   6615
      Left            =   0
      Picture         =   "frmBackgroundOptions.frx":1B4CD0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmBackgroundOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BackPic As String
Public ToolPic As String

Private Sub cmdApply_Click()

frmMain.PicBack.Picture = LoadPicture(BackPic)
frmMain.ImgStart.Picture = LoadPicture(ToolPic)

End Sub

Private Sub cmdBrowser_Click()

CM.filename = ""

CM.DialogTitle = "Select a background..."
CM.Filter = "All Suported Image Files|*.bmp;*.jpg;*.jpeg|All Files|*.*"
CM.ShowOpen

If CM.filename = "" Then
Else
imgPreview.Picture = LoadPicture(CM.filename)
BackPic = CM.filename
End If

End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()

frmMain.PicBack.Picture = LoadPicture(BackPic)
frmMain.ImgStart.Picture = LoadPicture(ToolPic)
Me.Hide

End Sub

Private Sub cmdBrowseT_Click()

CM.filename = ""

CM.DialogTitle = "Select a toolbar..."
CM.Filter = "All Suported Image Files|*.bmp;*.jpg;*.jpeg|All Files|*.*"
CM.ShowOpen

If CM.filename = "" Then
Else
imgStartPreview.Picture = LoadPicture(CM.filename)
ToolPic = CM.filename
End If

End Sub

Private Sub Form_Load()

'imgStartPreview.Picture = LoadPicture(ToolPic)
'imgBack.Picture = LoadPicture(BackPic)
ToolPic = "C:\Nexus\bin\Media\Wallpaper\BLUENESS.BMP"
BackPic = "C:\Nexus\bin\Media\Wallpaper\BLUENESS.BMP"

End Sub

Private Sub imgPreview_Click()

CM.filename = ""

CM.DialogTitle = "Select a background..."
CM.Filter = "All Suported Image Files|*.bmp;*.jpg;*.jpeg|All Files|*.*"
CM.ShowOpen

If CM.filename = "" Then
Else
imgPreview.Picture = LoadPicture(CM.filename)
BackPic = CM.filename
End If

End Sub

Private Sub imgStartPreview_Click()

CM.filename = ""

CM.DialogTitle = "Select a toolbar..."
CM.Filter = "All Suported Image Files|*.bmp;*.jpg;*.jpeg|All Files|*.*"
CM.ShowOpen

If CM.filename = "" Then
Else
imgStartPreview.Picture = LoadPicture(CM.filename)
ToolPic = CM.filename
End If

End Sub
