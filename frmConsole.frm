VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmConsole 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nexcon Console"
   ClientHeight    =   5955
   ClientLeft      =   1035
   ClientTop       =   1425
   ClientWidth     =   10110
   Icon            =   "frmConsole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10110
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   9480
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.TextBox txtCommand 
      Height          =   5205
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   9855
   End
   Begin VB.Image imgBack 
      Height          =   6135
      Left            =   0
      Picture         =   "frmConsole.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExecute_Click()
Script.ExecuteStatement (txtCommand.Text)
End Sub

