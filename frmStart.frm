VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open a Program"
   ClientHeight    =   2295
   ClientLeft      =   2850
   ClientTop       =   3465
   ClientWidth     =   6150
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6150
   Begin RichTextLib.RichTextBox RTB 
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393217
      TextRTF         =   $"frmStart.frx":030A
   End
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   120
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox cmbProg 
      Height          =   315
      ItemData        =   "frmStart.frx":038C
      Left            =   1440
      List            =   "frmStart.frx":038E
      TabIndex        =   0
      Text            =   "Select a program..."
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select a program from the list below to continue:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   360
      Picture         =   "frmStart.frx":0390
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   6765
      Left            =   0
      Picture         =   "frmStart.frx":07D2
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbProg_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
Call cmdRun_Click
Else
End If

'If KeyAscii = 126 Then
'frmConsole.Show
'Else
'End If

End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdRun_Click()
On Error Resume Next

If Left(cmbProg.Text, 9) = "NexconApp" Then
AppletName = Right(cmbProg.Text, Len(cmbProg.Text) - 10)

RTB.LoadFile (App.Path & "\apps\" & AppletName)
Script.ExecuteStatement (RTB.Text)
Me.Hide
Exit Sub

Else
End If

Select Case cmbProg.Text

Case "Select a program..."

Case "Shutdown Nexcon"
Unload frmAbout
Unload frmBackgroundOptions
Unload frmConsole
Unload frmFileViewer
Unload frmLoad
Unload frmLogin
Unload frmMain
Unload frmMessageBox
End
Unload Me


Case "Log Off Nexcon"
Me.Hide
frmMain.Hide
frmLogin.Show

Case "Notepad"
Shell ("C:\Windows\Notepad.exe")

Case "File Viewer"
frmFileViewer.Show
Me.Hide

Case "Program Off of Removeable Media..."
If InStr(Dir("A:\"), "info.ini") Then
If InStr(Dir("A:\"), "start.ini") Then
varNull = LoadRemovableMedia()
Else
End If
Else
Beep
varNull = MessageBox("Sorry, no removable media was found. If you were trying to load a CD make sure to insert the diskette into the floppy drive.", "No media found!")
End If

Case "Update Nexcon"
varNull = UpdateNexcon()
Me.Hide

Case Else
If Left$(cmbProg.Text, 1) = "#" Then
varNull = SystemFunction(cmbProg.Text)
Else
varNull = Replace(cmbProg.Text, " ", "")
Shell ("C:\Programs\" & varNull & "\program.exe")
End If

Me.Hide

End Select

End Sub


