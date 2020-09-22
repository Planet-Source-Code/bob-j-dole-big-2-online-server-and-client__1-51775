VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Big 2 CLIENT"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11880
   Begin MSWinsockLib.Winsock WinsockH 
      Left            =   10320
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txtChatLog 
      Height          =   1455
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "Chatstuff"
      Top             =   6120
      Width           =   11535
   End
   Begin VB.TextBox txtSend 
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Text            =   "send message"
      Top             =   7800
      Width           =   9855
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3960
      Top             =   120
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit the friggin game"
      Height          =   735
      Left            =   9600
      TabIndex        =   10
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdLose 
      Caption         =   "Forfiet Match"
      Height          =   735
      Left            =   9600
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdPass 
      Caption         =   "Pass"
      Enabled         =   0   'False
      Height          =   735
      Left            =   7320
      TabIndex        =   8
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Cards"
      Enabled         =   0   'False
      Height          =   735
      Left            =   7320
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opponent Number Of Cards"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label lblP1 
         Caption         =   "Label7"
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbl2 
         Caption         =   "Player 2"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblP2 
         Caption         =   "Label5"
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lbl3 
         Caption         =   "Player 3"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblP3 
         Caption         =   "Label3"
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl1 
         Caption         =   "Player 1"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   12
      Left            =   6600
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   1
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   11
      Left            =   5520
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   10
      Left            =   4440
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Label lblWait 
      Caption         =   "Player 1's Turn"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   9
      Left            =   3360
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   8
      Left            =   2280
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   7
      Left            =   1200
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   6
      Left            =   120
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   5
      Left            =   5520
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   4
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   3
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   2
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Image ImgPC 
      Height          =   1440
      Index           =   0
      Left            =   120
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Image ImgC 
      Height          =   1440
      Index           =   3
      Left            =   7800
      Top             =   480
      Width           =   1065
   End
   Begin VB.Image ImgC 
      Height          =   1440
      Index           =   2
      Left            =   6720
      Top             =   480
      Width           =   1065
   End
   Begin VB.Image ImgC 
      Height          =   1440
      Index           =   1
      Left            =   5640
      Top             =   480
      Width           =   1065
   End
   Begin VB.Image ImgC 
      Height          =   1440
      Index           =   0
      Left            =   4560
      Top             =   480
      Width           =   1065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'THIS IS THE CLIENT VERSION
'Big2 game, created by BoB Dole
'created by BoB Dole
'Features:
'Multiplayer FUnctions, 2-5 player game,chat
'Card Set, passing, ez to create another mod of game
'organization of cards
'V0.0 create form
'V0.1 created deck, player,card objects
'V0.2 initialize cards with suits
'V0.3 created pictures
'V0.4 initialized pictures, updated to screen,
      'organizing and 4 card done
'ENGINE IS COMPLETED!!!!!!!!!!!!!!!!!!!!!!!!!
'V0.5 created play, first turn  construct, legalize
'missing cards
'V0.5b fixed extremly weird error, rewrote deal function
'V0.6 created pass, turn order, internet connection
'V0.6b fixed out deal function, player 0 has a weird
'tendency to have bad cards, made it a bit more fair
'V0.7 Created send commands
'V0.8 Did some recieve commands and seperated data and command
'V0.9 make main game loop, and entire contruct
'V0.9b client/host seperation

'To COME:

'V1.0 polished form, made official, massive debugging
'V1.1 added chat, easier iping, aS, Joker
'V1.2 more user friendliness
'V1.3 update package: uses winsck api

Dim cardshigh As Integer 'so you dont go selecting
'all the cards
'For updates



Option Explicit 'catch undeclared variables

'then exit, but use the unload
Private Sub cmdExit_Click()
Form_Unload (0)
End Sub
'resign? Then make sure they lose
Private Sub cmdLose_Click()
MsgBox "YOU LOST!"
Form_Unload (1)
End Sub
'pass: cant play for rest of the game
Private Sub cmdPass_Click()
send 0, "pass"
frmMain.cmdPlay.Enabled = False
frmMain.cmdPass.Enabled = False

End Sub
'Play cards
Private Sub cmdPlay_Click()
Dim valid As Boolean
'Msgbox "cards selected: " & GetStringForm(ChallengeHand(0)) & " " & GetStringForm(ChallengeHand(1)) & " " & GetStringForm(ChallengeHand(2)) & " " & GetStringForm(ChallengeHand(3))
valid = Legalize(ChallengeHand(0), ChallengeHand(1), ChallengeHand(2), ChallengeHand(3), CurrentHand(0), CurrentHand(1), CurrentHand(2), CurrentHand(3))
If valid = True Then
  'TRUE
  'Msgbox "valid cards"
  Dim i As Integer
  'MAKE CURRENT = Challenge, DUMP CHALLENGE
  For i = 0 To 3
    CurrentHand(i) = ChallengeHand(i)
    ChallengeHand(i) = EmptySlot
  Next i
  'MAKE CARDS PLAYED EMPTYSLOTS
  Dim x As Integer
  For i = 0 To 3
    For x = 0 To 12
      If Player(0).Hand(x) = CurrentHand(i) Then
        Player(0).Hand(x) = EmptySlot
        Exit For
      End If
    Next x
  Next i
  OrganizeCards
  Update
  frmMain.cmdPass.Enabled = False
  frmMain.cmdPlay.Enabled = False
  send 0, "playcard/" & CurrentHand(0) & "/" & CurrentHand(1) & "/" & CurrentHand(2) & "/" & CurrentHand(3)
  
  'resets everything
  For i = 0 To 12
    ImgPC(i).BorderStyle = 0
    ImgPC(i).Appearance = 1
    cardshigh = 0
  Next i
  For i = 0 To 3
    ChallengeHand(i) = EmptySlot
  Next i
  
Else
'Resets everything anyways
  'Msgbox "invalid cards"
  For i = 0 To 12
    ImgPC(i).BorderStyle = 0
    ImgPC(i).Appearance = 1
    cardshigh = 0
  Next i
  For i = 0 To 3
    ChallengeHand(i) = EmptySlot
  Next i
End If

End Sub

'clears up everything updates screen
Private Sub Form_Load()

Dim i As Integer 'clears everything
For i = 0 To 3
  cardsselected(i) = EmptySlot
  CurrentHand(i) = EmptySlot
  ChallengeHand(i) = EmptySlot
Next i
ImgPC(3).Picture = CardPic(3)
Update
cardshigh = 0

End Sub

Private Sub Form_Unload(Cancel As Integer) 'frees up memory
Dim i As Integer
For i = 0 To 51
  Set CardPic(i) = Nothing 'clear up the card pics
Next i

End
End Sub
'this makes sure you only have 4 selected at a time
' and modifys selected and unselected cards
Private Sub ImgPC_Click(Index As Integer)
'First, check if it is a real "card" or not
If Player(0).Hand(Index) = EmptySlot Then
  'Msgbox "DEAD CARD!: IGNORING"
  Exit Sub
End If

  Dim i As Integer
If cardshigh = -1 Then
  MsgBox "CRITICAL ERROR:cardshigh is -1"
End If

If cardshigh = 4 And ImgPC(Index).BorderStyle = 0 Then
 'RESETS EVERYTHING
 'Msgbox "Imgpc_Click warning - You selected to many cards.  Reseting...", vbInformation, "Big 2- frmMain error"

  For i = 0 To 12
    ImgPC(i).BorderStyle = 0
    ImgPC(i).Appearance = 1
    cardshigh = 0
  Next i
  For i = 0 To 3
    ChallengeHand(i) = EmptySlot
  Next i
Else
  'SELECTS THE CARD
  If ImgPC(Index).BorderStyle = 0 Then
    ImgPC(Index).BorderStyle = 1
    ImgPC(Index).Appearance = 0
    cardshigh = cardshigh + 1
    i = 0
    For i = 0 To 3
      If ChallengeHand(i) = EmptySlot Then
        ChallengeHand(i) = Player(0).Hand(Index)
        Exit For
      End If
    Next i
  ElseIf ImgPC(Index).BorderStyle = 1 Then
    'DESELECTS THE CARD
    ImgPC(Index).BorderStyle = 0
    ImgPC(Index).Appearance = 1
    cardshigh = cardshigh - 1
    i = 0
    For i = 0 To 3
      If ChallengeHand(i) = Player(0).Hand(Index) Then
        Exit For
      End If
    Next i
    ChallengeHand(i) = EmptySlot
    
  End If
End If
End Sub



Private Sub WinsockH_Connect()
frmOnline.lblStatus.Caption = "CONNECTED"
frmOnline.cmdLock.Enabled = True
frmOnline.cmdConnectHost.Enabled = False
End Sub

Private Sub WinsockH_DataArrival(ByVal bytesTotal As Long)

Dim tdata As String 'Declare a variable
WinsockH.GetData tdata 'Get arrived data and store it in the variable
If ClientNumber = 0 Then
  ClientNumber = tdata
Else
  Recieve 0, tdata
End If

End Sub
