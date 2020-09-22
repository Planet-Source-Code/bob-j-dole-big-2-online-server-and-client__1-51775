VERSION 5.00
Begin VB.Form frmOnline 
   Caption         =   "CONNECTION FORM"
   ClientHeight    =   3885
   ClientLeft      =   705
   ClientTop       =   540
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "Secure Connection"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdConnectHost 
      Caption         =   "Connect to Host"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtPortHost 
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Text            =   "45"
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Caption         =   "Not Connected"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Connect to Host"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------CLIENT STARTS-------------------

Private Sub cmdConnectHost_Click()
frmMain.WinsockH.Close
frmMain.WinsockH.Connect "127.00.01", txtPortHost.Text
'Now it will check the connection

End Sub

Private Sub cmdLock_Click()
frmMain.WinsockH.SendData "connected"
End Sub



