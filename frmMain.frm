VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Easy DUN Connection & Disconnection"
   ClientHeight    =   600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisConnect 
      Caption         =   "&Disconnect"
      Height          =   465
      Left            =   2505
      TabIndex        =   1
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   465
      Left            =   1230
      TabIndex        =   0
      Top             =   75
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  'Please see comments in the accompanying module for explanations
  
Private Sub cmdConnect_Click()
  Connect_DUN

End Sub

Private Sub cmdDisConnect_Click()
  Disconnect_DUN
End Sub
