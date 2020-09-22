VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3240
   ClientLeft      =   1500
   ClientTop       =   2850
   ClientWidth     =   8535
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmInputBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   8535
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   8295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8055
   End
   Begin VB.Image PicBack 
      Height          =   6765
      Left            =   0
      Picture         =   "frmInputBox.frx":0CCA
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
MsgBoxDone = True
txtInput.Text = "CANCEL"
Me.Hide
End Sub

Private Sub cmdOK_Click()
MsgBoxDone = True
Me.Hide
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
MsgBoxDone = True
Me.Hide
Else
End If

End Sub
