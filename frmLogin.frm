VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   7200
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0CCA
   ScaleHeight     =   2130
   ScaleWidth      =   7200
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.TextBox txtPassWord 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   0
      Picture         =   "frmLogin.frx":1994
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
varUsername = txtUserName.Text
strPassword = txtPassWord.Text
MsgBoxDone = True
End Sub
