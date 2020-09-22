Attribute VB_Name = "mdlPicture"
'sets up the pics and tells you where the pics are
Option Explicit
Public CardPic(0 To 51) As StdPicture



Public Sub InitPics()
Dim cardno As String
Dim i As Integer



For i = 0 To 51
  cardno = GetStringForm(i)
  Set CardPic(i) = New StdPicture
  Set CardPic(i) = LoadPicture(App.Path & "\Card Pictures\" & cardno & ".bmp")
Next i


End Sub


'converts a number into the pic name
Public Function GetStringForm(i As Integer) As String
Dim x As String 'temp holder
Select Case i
  Case 0
    x = "3D" '3
  Case 1
    x = "3C"
  Case 2
    x = "3H"
  Case 3
    x = "3S"
  Case 4
    x = "4D" '4
  Case 5
    x = "4C"
  Case 6
    x = "4H"
  Case 7
    x = "4S"
  Case 8
    x = "5D"
  Case 9
    x = "5C"
  Case 10
    x = "5H"
  Case 11
    x = "5S"
  Case 12
    x = "6D"
  Case 13
    x = "6C"
  Case 14
    x = "6H"
  Case 15
    x = "6S"
  Case 16
    x = "7D"
  Case 17
    x = "7C"
  Case 18
    x = "7H"
  Case 19
    x = "7S"
  Case 20
    x = "8D"
  Case 21
    x = "8C"
  Case 22
    x = "8H"
  Case 23
    x = "8S"
  Case 24
    x = "9D"
  Case 25
    x = "9C"
  Case 26
    x = "9H"
  Case 27
    x = "9S"
  Case 28
    x = "10D"
  Case 29
    x = "10C"
  Case 30
    x = "10H"
  Case 31
    x = "10S"
  Case 32
    x = "JD"
  Case 33
    x = "JC"
  Case 34
    x = "JH"
  Case 35
    x = "JS"
  Case 36
    x = "QD"
  Case 37
    x = "QC"
  Case 38
    x = "QH"
  Case 39
    x = "QS"
  Case 40
    x = "KD"
  Case 41
    x = "KC"
  Case 42
    x = "KH"
  Case 43
    x = "KS"
  Case 44
    x = "AD"
  Case 45
    x = "AC"
  Case 46
    x = "AH"
  Case 47
    x = "AS"
  Case 48
    x = "2D"
  Case 49
    x = "2C"
  Case 50
    x = "2H"
  Case 51
    x = "2S"
  Case EmptySlot
    x = "EmptySlot"
  Case Else
    MsgBox "Error in GetStringForm function"
End Select
GetStringForm = x
End Function
