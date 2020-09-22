Attribute VB_Name = "mdlCentral"
'this module contains the main functions for the game and cards
'used to lag on non paid programs
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Option Explicit
Public Card(0 To 51) As Cards 'Creates 52 cards
Public Player(0 To 3) As Players 'Creates 4 players
Public CardPerPlayer As Integer '# cards per player
Public Const EmptySlot = 314 'Tell you if slot in array is empty
Public cardsselected(0 To 3) As Integer 'used in form to
Public CurrentHand(0 To 3) As Integer 'current played hand
Public ChallengeHand(0 To 3) As Integer 'passed to legalized
'see which cards in hand are selected
'cardssel, curhand, chahand init in form load
Public GameList(0 To 3) As Integer 'holds the game list
Public RoundList(0 To 3) As Integer  ' holds round list
Public Roundmarker As Integer

'user friendly variables
Public NoPlayerCards(1 To 3) As Integer
Public WaitForPlayer As Integer
Public ClientNumber As Integer

'player object
Type Players
  name As String
  Index As Integer
  Hand(0 To 12) As Integer 'each hand holds an "'i' # that
  'represents card(i)
  NoCards As Integer '# cards
End Type




Public Sub Main()
Attribute Main.VB_Description = "The first prodcedure that runs."

MsgBox "PLZ look at THE READ ME!!!.txt file Before you do anything"


InitCommonControlsXP 'Makes XP controls


'MsgBox "In Order to use this, it require the latest Windows Updates.  Especially required is mswinsck.ocx."
Randomize Time 'randomizes stuff

InitCards 'init cards and pics

'--------comment for client version---------
Deal 'deals to 4 players
OrganizeCards 'organizes the cards
InitTurnList 'make turnlist
'------------------------------------------

Load frmMain
frmOnline.Show vbModal

frmMain.Show
'Thanks to Lee Hughes lphughes@btopenworld.com for Xp controls

'send cards to everyone else
send 1, "dealcard/" & Player(1).Hand(0) & "/" & Player(1).Hand(1) & "/" & Player(1).Hand(2) & "/" & Player(1).Hand(3) & "/" & Player(1).Hand(4) & "/" & Player(1).Hand(5) & "/" & Player(1).Hand(6) & "/" & Player(1).Hand(7) & "/" & Player(1).Hand(8) & "/" & Player(1).Hand(9) & "/" & Player(1).Hand(10) & "/" & Player(1).Hand(11) & "/" & Player(1).Hand(12)
send 2, "dealcard/" & Player(2).Hand(0) & "/" & Player(2).Hand(1) & "/" & Player(2).Hand(2) & "/" & Player(2).Hand(3) & "/" & Player(2).Hand(4) & "/" & Player(2).Hand(5) & "/" & Player(2).Hand(6) & "/" & Player(2).Hand(7) & "/" & Player(2).Hand(8) & "/" & Player(2).Hand(9) & "/" & Player(2).Hand(10) & "/" & Player(2).Hand(11) & "/" & Player(2).Hand(12)
send 3, "dealcard/" & Player(3).Hand(0) & "/" & Player(3).Hand(1) & "/" & Player(3).Hand(2) & "/" & Player(3).Hand(3) & "/" & Player(3).Hand(4) & "/" & Player(3).Hand(5) & "/" & Player(3).Hand(6) & "/" & Player(3).Hand(7) & "/" & Player(3).Hand(8) & "/" & Player(3).Hand(9) & "/" & Player(3).Hand(10) & "/" & Player(3).Hand(11) & "/" & Player(3).Hand(12)
GetNoOfClientCards

'client stuff than comment
GameLoop
End Sub
Public Sub GetNoOfClientCards()
Dim i, x As Integer
For i = 0 To 1000
  x = x + 34
  x = x - 34
Next i
send SendToAll, "givenoofcards"
End Sub

Public Sub GameLoop()
Dim tnextplayer As Integer


tnextplayer = GetTurn
If tnextplayer = 0 Then 'if it is host than your turn
  frmMain.cmdPass.Enabled = True
  frmMain.cmdPlay.Enabled = True
  WaitForPlayer = 0
Else
  send tnextplayer, "turn"
  WaitForPlayer = tnextplayer
End If



'Tell all "waiting for that player"
'do that in next version

'once recieved, send to all: update

'get noofcards :newer version
'send ToAllPlayer, "givenoofcards" : newer versions
'update noofcards : newer version

'if one person has 0 then, he is winner

End Sub



'updates current cards, and card pics
Public Sub Update()
Dim i As Integer


'updates hand
For i = 0 To 12
  If Player(0).Hand(i) <> EmptySlot Then
    frmMain.ImgPC(i).Picture = CardPic(Player(0).Hand(i))
  ElseIf Player(0).Hand(i) = EmptySlot Then
    frmMain.ImgPC(i).Picture = Nothing
    frmMain.ImgPC(i).Refresh
     
  End If
Next i

'update current cards
For i = 0 To 3
  If CurrentHand(i) <> EmptySlot Then
    frmMain.ImgC(i).Picture = CardPic(CurrentHand(i))
  ElseIf CurrentHand(i) = EmptySlot Then
    frmMain.ImgC(i).Picture = Nothing
    frmMain.ImgC(i).Refresh
  End If
Next i
    
'update noofcards
frmMain.lblP1 = NoPlayerCards(1)
frmMain.lblP2 = NoPlayerCards(2)
frmMain.lblP3 = NoPlayerCards(3)
    
'update who is playing
frmMain.lblWait.Caption = "Waiting for Player " & WaitForPlayer

End Sub

Public Sub InitTurnList()
Dim p, c, i, x As Integer
Dim firstplayer As Integer
For p = 0 To 3 'find who has 3 of diamonds
  If Player(p).Hand(0) = 0 Then
    firstplayer = p
    Debug.Print "Player: " & firstplayer & " is the first player"
    Exit For
  End If
Next p
Roundmarker = firstplayer

For i = 0 To 3 'sets up the game list
  GameList(i) = firstplayer
  firstplayer = firstplayer + 1
  If firstplayer = 4 Then
    firstplayer = 0
  End If
Next i
For i = 0 To 3 'makes game list = round list
  RoundList(i) = GameList(i)
Next i

End Sub
Public Function GetTurn() As Integer
GetTurn = Roundmarker 'give to function
'Now calculate for next round
Roundmarker = Roundmarker + 1 'adds one : logically works
'if no one passed
If Roundmarker = 4 Then 'make it 0 if it is 4
  Roundmarker = 0
End If

'now skip to the next if someone passed
Dim i As Integer ' for the loop
If RoundList(Roundmarker) = EmptySlot Then 'if that one is
  'empty move on
  Roundmarker = Roundmarker + 1
  If Roundmarker = 4 Then Roundmarker = 0
  
  If RoundList(Roundmarker) = EmptySlot Then 'check next
    Roundmarker = Roundmarker + 1
    If Roundmarker = 4 Then Roundmarker = 0
    
  End If
End If

End Function
Public Sub Pass(Passer As Integer)
RoundList(Passer) = EmptySlot 'make it empty
Dim npr As Integer 'number of players in the round
Dim i As Integer 'loop counter

For i = 0 To 3 'now find number of people still playing
  If RoundList(i) <> EmptySlot Then
    npr = npr + 1
  End If
Next i

If npr = 1 Then 'if one person left then...
  For i = 0 To 3
    If RoundList(i) <> EmptySlot Then
      'we got the find the winner, so this is the way
      EndRound (i)
      Exit For
    End If
  Next i
    
    
ElseIf npr = 0 Then 'weird, last person passed on free move
  Roundmarker = Passer + 1
  If Roundmarker = 4 Then Roundmarker = 0
  EndRound (Roundmarker)
  
End If

End Sub

Public Sub EndRound(RoundWinner As Integer)
Dim i As Integer
For i = 0 To 3
  RoundList(i) = GameList(i) 'reset roundlist
Next i
send SendToAll, "currenthand/314/314/314/314" 'reset the current cards
'reset the current cards
For i = 0 To 3
  CurrentHand(i) = EmptySlot
Next i
Roundmarker = RoundWinner
Update
End Sub

Public Sub AddChat(text As String)
frmMain.txtChatLog.text = frmMain.txtChatLog.text & vbCrLf & text
End Sub


Public Sub Win(Winner As Integer)
'this will make the winner win, remove from gamelist
GameList(Winner) = EmptySlot
Dim tnextguy As Integer 'find nextplayer
tnextguy = GetTurn
EndRound tnextguy 'end the round with free move
End Sub

