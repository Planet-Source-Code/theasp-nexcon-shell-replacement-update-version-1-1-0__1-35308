VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmFileViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nexcon File Viewer"
   ClientHeight    =   8160
   ClientLeft      =   405
   ClientTop       =   435
   ClientWidth     =   11145
   Icon            =   "frmFileViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11145
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.ComboBox cmbLoc 
      Height          =   315
      ItemData        =   "frmFileViewer.frx":0CCA
      Left            =   2520
      List            =   "frmFileViewer.frx":0CCC
      TabIndex        =   3
      Text            =   "C:\My Files\"
      Top             =   600
      Width           =   8535
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward >"
      Height          =   855
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   10935
      ExtentX         =   19288
      ExtentY         =   12303
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image Image1 
      Height          =   8175
      Left            =   0
      Picture         =   "frmFileViewer.frx":0CCE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "frmFileViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbLoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
If Left$(cmbLoc.Text, 7) = "http://" Then
cmbLoc.Text = "This is not a web browser, start up browser if you would like to access the internet!"
Else
varBuffer = Replace(UCase(cmbLoc.Text), "ROOT\", "C:\")
varBuffer = Replace(UCase(varBuffer), "FLOPPY\", "A:\")
varBuffer = Replace(UCase(varBuffer), "CD\", "D:\")

WB.Navigate varBuffer
cmbLoc.AddItem cmbLoc.Text
End If

Else
End If
End Sub

Private Sub cmdBack_Click()
On Error Resume Next
WB.GoBack
End Sub

Private Sub cmdForward_Click()
On Error Resume Next
WB.GoForward
End Sub

Private Sub cmdHide_Click()

If SM_1P = "fileviewer" Then
Me.Hide
frmMain.lblS1.Visible = True
frmMain.imgS1.Visible = True
ElseIf SM_2P = "fileviewer" Then
Me.Hide
frmMain.lblS2.Visible = True
frmMain.imgS2.Visible = True
Else
If SM_1 = False Then
frmMain.imgS1.Picture = LoadPicture(App.Path & "\pics\fileviewer.ico")
frmMain.lblS1.Caption = "File Viewer"
frmMain.imgS1.Visible = True
frmMain.lblS1.Visible = True
SM_1 = True
SM_1P = "fileviewer"
Me.Hide
ElseIf SM_2 = False Then
frmMain.imgS2.Picture = LoadPicture(App.Path & "\pics\fileviewer.ico")
frmMain.lblS2.Caption = "File Viewer"
frmMain.imgS2.Visible = True
frmMain.lblS2.Visible = True
SM_2 = True
SM_2P = "fileviewer"
Me.Hide
Else
varNull = MessageBox("Sorry, there is no room left on the task bar! Try closing any programs you aren't using.", "Out of room!")
End If
End If

End Sub

Private Sub Form_Load()
On Error Resume Next
WB.Navigate "C:\My Files\"
End Sub

Private Sub WB_DocumentComplete(ByVal pDisp As Object, URL As Variant)

varBuffer = Replace(UCase(App.Path), "C:\", "ROOT\")
If UCase(URL) = UCase(App.Path) & "\" Then
cmbLoc.Text = "RESTRICTED"
ElseIf UCase(URL) = UCase(App.Path) Then
cmbLoc.Text = "RESTRICTED"
End If

If URL = App.Path & "\" Then
WB.Navigate App.Path & "\restricted.html"

frmMain.RTB.LoadFile (App.Path & "\log.log")
frmMain.RTB.Text = frmMain.RTB.Text & vbCrLf & "SECURITY: USER TRIED TO ACCESS NEXCON ROOT FOLDER"
frmMain.RTB.SaveFile (App.Path & "\log.log")

ElseIf URL = App.Path Then

frmMain.RTB.LoadFile (App.Path & "\log.log")
frmMain.RTB.Text = frmMain.RTB.Text & vbCrLf & "SECURITY: USER TRIED TO ACCESS NEXCON ROOT FOLDER"
frmMain.RTB.SaveFile (App.Path & "\log.log")

WB.Navigate App.Path & "\restricted.html"
Else
End If

URL = Replace(UCase(URL), "C:\", "ROOT\")
URL = Replace(UCase(URL), "A:\", "FLOPPY\")
URL = Replace(UCase(URL), "D:\", "CD\")
cmbLoc.Text = URL

End Sub
