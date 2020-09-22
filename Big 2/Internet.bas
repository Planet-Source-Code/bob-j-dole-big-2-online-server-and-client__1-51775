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
Public Const SendToAll = 432 'this is just a rnd number

Public Sub send(ToPlayer As Integer, Info As String)
Dim message As String
message = seperator & Info

If ToPlayer = 1 Then
  frmMain.Winsock1.SendData message
ElseIf ToPlayer = 2 Then
  frmMain.Winsock2.SendData message
ElseIf ToPlayer = 3 Then
  frmMain.Winsock3.SendData message
ElseIf ToPlayer = SendToAll Then
  frmMain.Winsock1.SendData message
  frmMain.Winsock2.SendData message
  frmMain.Winsock3.SendData message
End If
DoEvents 'If it isn't there, there would be a stack error
End Sub
Public Sub Recieve(Playerno As Integer, Command As String)
Dim tcommand As String 'doesnt mess with command, only tcommand
Dim seplocation As Long 'find seperator
Dim fcommand As String 'stores real command
Dim fdata As String 'stores all the data

tcommand = Right(Command, Len(Command) - 1)
'first it removes / in command

If tcommand <> "pass" Then 'if no /
  'now seperate command from data "player/0/1/2/314"
  'find /
  seplocation = InStr(1, tcommand, seperator)
  'now get the left of the /
  fcommand = Left(tcommand, seplocation - 1)
  '-1 because we dont want the slash
  fdata = Right(tcommand, Len(tcommand) - seplocation)
Else
  fcommand = "pass"
  fdata = ""
End If

CommandConstruct Playerno, fcommand, fdata


End Sub
Public Sub CommandConstruct(Playerno As Integer, Command As String, data As String)
Dim i As Integer
Select Case Command
  Case "pass" 'pass
    Pass (Playerno)
    GameLoop
    
  Case "playcard" 'playcard 0/1/2/314
    Dim seplocation As Long
    Dim tcard(0 To 3) As Integer
    Dim tdata As String
    tdata = data
    
    seplocation = InStr(1, tdata, seperator)
    tcard(0) = Left(data, seplocation - 1) 'get first card
    tdata = Right(tdata, Len(tdata) - seplocation)
    
    seplocation = InStr(1, tdata, seperator)
    tcard(1) = Left(tdata, seplocation - 1) 'get second card
    tdata = Right(tdata, Len(tdata) - seplocation)
    
    seplocation = InStr(1, tdata, seperator)
    tcard(2) = Left(tdata, seplocation - 1) 'get third card
    tdata = Right(tdata, Len(tdata) - seplocation)
    
    tcard(3) = tdata 'get last card
    
    
    send SendToAll, "currenthand/" & tcard(0) & "/" & tcard(1) & "/" & tcard(2) & "/" & tcard(3)
    For i = 0 To 3
      CurrentHand(i) = tcard(i)
    Next i
    send Playerno, "givenoofcards"
    Update
    GameLoop
  Case "noofcards" 'noofcard 23
    NoPlayerCards(Playerno) = data
    If NoPlayerCards(Playerno) = 0 Then
      Win Playerno
  
    End If
    Update
  Case "chattohost"
    
  Case Else
    MsgBox "Weird client command: " & Command & "/" & data
End Select
End Sub

