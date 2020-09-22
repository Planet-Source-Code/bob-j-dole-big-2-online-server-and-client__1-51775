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
InitCommonControlsXP 'Makes XP controls

'MsgBox "In Order to use this, it require the latest Windows Updates.  Especially required is mswinsck.ocx."
Randomize 'randomizes stuff

InitCards 'init cards and pics

'--------comment for client version---------
'Deal 'deals to 4 players
'OrganizeCards 'organizes the cards
'InitTurnList 'make turnlist
'------------------------------------------

Load frmMain
frmOnline.Show vbModal
frmMain.Show
'Thanks to Lee Hughes lphughes@btopenworld.com for Xp controls

'send cards to everyone else
'send 1, "dealcard/" & Player(1).Hand(0) & "/" & Player(1).Hand(1) & "/" & Player(1).Hand(2) & "/" & Player(1).Hand(3) & "/" & Player(1).Hand(4) & "/" & Player(1).Hand(5) & "/" & Player(1).Hand(6) & "/" & Player(1).Hand(7) & "/" & Player(1).Hand(8) & "/" & Player(1).Hand(9) & "/" & Player(1).Hand(10) & "/" & Player(11).Hand(11) & "/" & Player(1).Hand(12)
'send 2, "dealcard/" & Player(2).Hand(0) & "/" & Player(2).Hand(1) & "/" & Player(2).Hand(2) & "/" & Player(2).Hand(3) & "/" & Player(2).Hand(4) & "/" & Player(2).Hand(5) & "/" & Player(2).Hand(6) & "/" & Player(2).Hand(7) & "/" & Player(2).Hand(8) & "/" & Player(2).Hand(9) & "/" & Player(2).Hand(10) & "/" & Player(11).Hand(11) & "/" & Player(2).Hand(12)
'send 3, "dealcard/" & Player(3).Hand(0) & "/" & Player(3).Hand(1) & "/" & Player(3).Hand(2) & "/" & Player(3).Hand(3) & "/" & Player(3).Hand(4) & "/" & Player(3).Hand(5) & "/" & Player(3).Hand(6) & "/" & Player(3).Hand(7) & "/" & Player(3).Hand(8) & "/" & Player(3).Hand(9) & "/" & Player(3).Hand(10) & "/" & Player(11).Hand(11) & "/" & Player(3).Hand(12)


'client stuff than comment
'GameLoop
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
frmMain.Caption = "Big 2  Client Version | Player: " & ClientNumber

End Sub


