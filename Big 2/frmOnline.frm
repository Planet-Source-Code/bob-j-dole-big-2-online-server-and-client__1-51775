VERSION 5.00
Begin VB.Form frmOnline 
   Caption         =   "CONNECTION FORM"
   ClientHeight    =   3840
   ClientLeft      =   705
   ClientTop       =   540
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtPort3 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Text            =   "4570"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtPort2 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Text            =   "4569"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtPort1 
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Text            =   "4568"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblP3 
      Caption         =   "Player 3 - not connected"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblP2 
      Caption         =   "Player 2 - not connected"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblP1 
      Caption         =   "Player 1 - not connected"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'HOST STUFF---------------------------------------------
Private Sub cmdConnect_Click()
frmMain.Winsock1.Close
frmMain.Winsock2.Close
frmMain.Winsock3.Close

frmMain.Winsock1.LocalPort = txtPort1.text
frmMain.Winsock2.LocalPort = txtPort2.text
frmMain.Winsock3.LocalPort = txtPort3.text

frmMain.Winsock1.Listen
frmMain.Winsock2.Listen
frmMain.Winsock3.Listen

cmdConnect.Enabled = False 'so you cant keep conecting
End Sub

'--------------------END HOST STUFF-----------------------

