Attribute VB_Name = "mdlInternet"
'All internet functions and connecting, and sending and reciving
'go here
Option Explicit
Public Const seperator = "/" 'seperates everything



'Functions you can send/recieve
'#Host commands(used by host)
'#Comands Sent:
'/turn
'/givenoofcards
'/dealcard [card 0] [card 1] [card 2] ...
'/place [player 0 to 3] [place 0 to 3]
'/chattoall [string] V1.1
'/currenthand [card 0] [card 1] [card 2] [card 3]
'/noofcards [no of cards] [Player no}
'#Comands Recived are commands sent by client


'#Client Commands
'#Commands Sent:
'/pass
'/playcard [card 0] [card 1] [card 2] [card 3]
'/noofcards [no of cards]
'/chattohost [string]
'#Comands Recived are commands sent by host

'Host knows where to send, either loop all, or individual
'When client sends, first var is player number
'"2/pass"
Public Const SendToAll = 432

Public Sub send(ToPlayer As Integer, Info As String)
Dim message As String
message = seperator & Info

frmMain.WinsockH.SendData message

DoEvents
End Sub
Public Sub Recieve(Playerno As Integer, Command As String)
Dim tcommand As String 'doesnt mess with command, only tcommand
Dim seplocation As Long 'find seperator
Dim fcommand As String 'stores real command
Dim fdata As String 'stores all the data

tcommand = Right(Command, Len(Command) - 1)
'MsgBox tcommand 'first it removes / in command

If tcommand <> "turn" And tcommand <> "givenoofcards" Then
  'now seperate command from data "player/0/1/2/314"
  'find /
  seplocation = InStr(1, tcommand, seperator)
  'now get the left of the /
  fcommand = Left(tcommand, seplocation - 1)
  '-1 because we dont want the slash
  fdata = Right(tcommand, Len(tcommand) - seplocation)
Else
  fcommand = tcommand
  fdata = ""
End If

'Msgbox "command : " & fcommand
'Msgbox "data: " & fdata

CommandConstruct Playerno, fcommand, fdata

End Sub

Public Sub CommandConstruct(Playerno As Integer, Command As String, data As String)
Dim i As Integer
Dim seplocation As Long
Dim tdata As String

Select Case Command
  Case "currenthand" 'Updates the current hand
    i = 0
    seplocation = 0
    tdata = ""
    tdata = data
    
    For i = 0 To 2
      seplocation = InStr(1, tdata, seperator)
      CurrentHand(i) = Left(tdata, seplocation - 1) 'get card
      tdata = Right(tdata, Len(tdata) - seplocation)
    Next i
    CurrentHand(3) = tdata
    'Msgbox "new cards are : " & CurrentHand(0) & " " & CurrentHand(1) & " " & CurrentHand(2) & " " & CurrentHand(3)
    Update
    
  Case "turn" 'Yo its your turn, let him play
    frmMain.cmdPass.Enabled = True
    frmMain.cmdPlay.Enabled = True
    WaitForPlayer = ClientNumber
    Beep
    MsgBox "Wake up! Its your turn!"
  Case "givenoofcards" 'if host wants card, than GIVE HIM CARDS
    'first calculate number of cards
    Dim tnoofcards As Integer
    For i = 0 To 12
      If Player(0).Hand(i) <> EmptySlot Then
        tnoofcards = tnoofcards + 1
      End If
    Next i
    send 0, "noofcards/" & tnoofcards
    
    
        
  Case "dealcard" 'recieves random cards from host
    i = 0
    seplocation = 0
    tdata = ""
    tdata = data
    
    For i = 0 To 11
      seplocation = InStr(1, tdata, seperator)
      Player(0).Hand(i) = Left(tdata, seplocation - 1) 'get card
      tdata = Right(tdata, Len(tdata) - seplocation)
    Next i
    Player(0).Hand(12) = tdata
    frmOnline.Hide
    Update
  Case "chattohost" 'Update chat screen
    frmMain.txtChatLog.Text = frmMain.txtChatLog.Text & vbCrLf & data
  Case Else
    MsgBox "Weird client command: " & Command & "/" & data
End Select
End Sub
