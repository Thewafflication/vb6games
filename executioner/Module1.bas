Attribute VB_Name = "Module1"
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
(ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
' Function above allows sound to be used using the Windows Multimedia DLL

Public PlayerTwoDead As Boolean ' Used to See if Player 2 Dead
Public PlayerOneDead As Boolean
Public TimeLeft As Integer 'Time left for Player one
Public TimeSave As Integer ' Save the original ( or Last value ) of the Time setting
Public HintSave As Integer '  Save the original ( or Last value ) of the AMount of Hints
Public PhraseLength As Integer 'Length of Phrase
Public PhraseUnknown As Variant ' holds
Public PhraseArray(0 To 2000) As Variant
Public RandomPhrase, Turn1player, Try1, Try2, Hints, Hints2, Players As Integer
Public TimeOff, WrongLetter, First, There, PLAYERTWO, PLAYERONE As Boolean
Public Player2name, Player1name, Letter As String
Public bSound As Boolean
Public iTimesWon As Integer ' Number of times won (used to reset x,y of picwinner after 10 times won to avoid pic box overlapping another picbox)

Public Sub Main()
bSound = True
Load frmSplash
frmSplash.Show 'Show The Splash Screen at startup
frmSplash.Refresh 'Refresh Splash form
Load frmGamescreen 'Load Gamescreen into Memory
'Unload frmSplash 'Remove Splash from Screen
'frmGamescreen.Show 'Start the Game Screen

End Sub


