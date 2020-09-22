VERSION 5.00
Begin VB.Form frmLock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computer Locked."
   ClientHeight    =   2025
   ClientLeft      =   2355
   ClientTop       =   3825
   ClientWidth     =   6960
   Icon            =   "frmLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6960
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Unlock"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This computer has been locked. Please type in the password to unlock it."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   6765
      Left            =   0
      Picture         =   "frmLock.frx":0CCA
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
On Error Resume Next

If txtPass.Text = "" Then
varNull = MessageBox("You forgot to enter a password!", "Oops!")
Else
If txtPass.Text = strPassword Then
frmMain.Show
frmMain.WM.Play
Me.Hide
Else
varNull = MessageBox("Invalid Password.", "Sorry.")
End If
End If

End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then

If txtPass.Text = varPassword Then
frmMain.Show
frmMain.WM.Play
Me.Hide
Else
varNull = MessageBox("Invalid Password.", "Sorry.")
End If

Else
End If

End Sub
