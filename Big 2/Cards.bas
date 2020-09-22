Attribute VB_Name = "mdlCards"
Option Explicit
'A card object
Type Cards
  Number As Integer 'number defines from 3 to 2
  Suit As Suits '4 suits
  Index As Integer 'index for organizing cards
End Type
'the suits for easier if's
Enum Suits
  Diamond
  Club
  Heart
  Spade
End Enum

'the numbers
Enum Numbers
  Three
  Four
  Five
  Six
  Seven
  Eight
  Nine
  ten
  Jack
  Queen
  King
  Ace
  Two
End Enum
'initializes the cards numbers and pictures
Public Sub InitCards()
Dim s, n, i As Integer

i = 0


For n = 0 To 12
  For s = 0 To 3
    Card(i).Number = n
    Card(i).Suit = s
    Card(i).Index = i
    DoEvents
    i = i + 1
  Next s
Next n
InitPics
End Sub


'Deals the players the cards
Public Sub Deal()
Dim pcards(0 To 3) As Integer
Dim c, r As Integer
CardPerPlayer = 13


For c = 0 To 51 'Go through each card giving them out to a player
  
  r = Int(Rnd * 3) 'makes a rnd player number
  If Player(r).NoCards < CardPerPlayer Then 'if the player still needs cards
    Player(r).Hand(pcards(r)) = c 'make that player card(pcard) = the c
    pcards(r) = pcards(r) + 1 'add the pcard by 1 so it will increase
    Player(r).NoCards = Player(r).NoCards + 1 'add the players number of cards by 1
  Else
    r = r + 1 'try next player
    If r = 4 Then r = 0 'if it overuns, put it to 0
    If Player(r).NoCards < CardPerPlayer Then 'if the player still needs cards
      Player(r).Hand(pcards(r)) = c 'make that player card(pcard) = the c
      pcards(r) = pcards(r) + 1 'add the pcard by 1 so it will increase
      Player(r).NoCards = Player(r).NoCards + 1 'add the players number of cards by 1
    Else
      r = r + 1
      If r = 4 Then r = 0 'if it overuns, put it to 0
      If Player(r).NoCards < CardPerPlayer Then 'if the player still needs cards
        Player(r).Hand(pcards(r)) = c 'make that player card(pcard) = the c
        pcards(r) = pcards(r) + 1 'add the pcard by 1 so it will increase
        Player(r).NoCards = Player(r).NoCards + 1 'add the players number of cards by 1
      Else
        r = r + 1
        If r = 4 Then r = 0 'if it overuns, put it to 0
        If Player(r).NoCards < CardPerPlayer Then 'if the player still needs cards
          Player(r).Hand(pcards(r)) = c 'make that player card(pcard) = the c
          pcards(r) = pcards(r) + 1 'add the pcard by 1 so it will increase
          Player(r).NoCards = Player(r).NoCards + 1 'add the players number of cards by 1
        Else
        
          MsgBox "weird error in deal, all players full"
        End If
      End If
    End If
  End If
Next c
End Sub

'used to sort the cards by lowest to highest,
'and get rid of spaces
Public Sub OrganizeCards()
Dim temphand(0 To 12) As Integer
Dim x, i, y As Integer
For i = 0 To 51 'look at one card, and searches entire
'array for it
  For x = 0 To 12
    If Player(0).Hand(x) = i Then ' if it is contained
      temphand(y) = Player(0).Hand(x) 'store into hand
      y = y + 1 'increment counter
    End If
  Next x
Next i

For y = y To 12 'fill in the rest if played cards
  temphand(y) = EmptySlot
Next y

i = 0
For i = 0 To 12
  Player(0).Hand(i) = temphand(i)
Next i
Update

End Sub

'makes sure the move is legal
Public Function Legalize(chgcard1 As Integer, chgcard2 As Integer, chgcard3 As Integer, chgcard4 As Integer, curcard1 As Integer, curcard2 As Integer, curcard3 As Integer, curcard4 As Integer) As Boolean
'change it back into an array
Dim chg(0 To 3) As Integer
Dim cur(0 To 3) As Integer

chg(0) = chgcard1
chg(1) = chgcard2
chg(2) = chgcard3
chg(3) = chgcard4

cur(0) = curcard1
cur(1) = curcard2
cur(2) = curcard3
cur(3) = curcard4

'First see if they are singles triples doubles quads
Dim chgnumofcards As Integer
Dim curnumofcards As Integer
Dim i 'gets the number of cards
For i = 0 To 3
  If chg(i) <> EmptySlot Then
    chgnumofcards = chgnumofcards + 1
  End If
Next i
i = 0
For i = 0 To 3
  If cur(i) <> EmptySlot Then
    curnumofcards = curnumofcards + 1
  End If
Next i



'first make sure chg cards are the same type
Dim sametype As Boolean
i = 0
Dim x As Integer
If chgnumofcards = 1 Then
  sametype = True
ElseIf chgnumofcards = 2 Then
  If Card(chg(0)).Number = Card(chg(1)).Number Then
    sametype = True
  End If
ElseIf chgnumofcards = 3 Then
  If Card(chg(0)).Number = Card(chg(1)).Number And Card(chg(0)).Number = Card(chg(2)).Number Then
    sametype = True
  End If
ElseIf chgnumofcards = 4 Then
  If Card(chg(0)).Number = Card(chg(1)).Number And Card(chg(0)).Number = Card(chg(2)).Number And Card(chg(0)).Number = Card(chg(3)).Number Then
    sametype = True
  End If
Else
  sametype = False
End If


If sametype = True Then
  'before we get into more complicated coding, if it contains
  'the 3 dim, it wins automatically
  For i = 0 To 3
    If chg(i) = 0 Then
      Legalize = True
      Exit Function
    End If
  Next i
  
  If chgnumofcards = curnumofcards Then 'FIRST CHECK IF IT IS 1 to 1, 2 to 2
    If Card(chg(0)).Number > Card(cur(0)).Number Then 'IF CARD NUMBER IS HIGHER IT WINS
      Legalize = True
    ElseIf chgnumofcards = 1 And chg(0) > cur(0) Then
      Legalize = True
    ElseIf chgnumofcards = 2 Then 'IF IT LOSES CHECK FOR DBL TRICK
      If Card(chg(0)).Suit = Heart Or Card(chg(0)).Suit = Spade Then 'IF ANY CARD IS HEART OR SPADE IT WON
        Legalize = True
      Else
        Legalize = False
      End If
    Else
      Legalize = False
    End If
  ElseIf curnumofcards = 0 Then 'if it is the guys free move legalize
    Legalize = True
  Else
    Legalize = False
  End If
Else
  Legalize = False
End If

End Function
