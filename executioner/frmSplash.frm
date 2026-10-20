VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   9405
   ClientTop       =   5370
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Time2go 
      Interval        =   100
      Left            =   1440
      Top             =   1440
   End
   Begin VB.Timer timerbegin 
      Interval        =   20
      Left            =   600
      Top             =   1320
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2880
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   3135
      TabIndex        =   17
      Top             =   1560
      Width           =   3135
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3000
      Picture         =   "frmSplash.frx":0B62
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   16
      Top             =   360
      Width           =   3015
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   1
      Left            =   0
      Picture         =   "frmSplash.frx":18B5
      ScaleHeight     =   1335
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   0
      Left            =   0
      Picture         =   "frmSplash.frx":1E61
      ScaleHeight     =   1335
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   8040
      Picture         =   "frmSplash.frx":240D
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   0
      Left            =   8040
      Picture         =   "frmSplash.frx":2C58
      ScaleHeight     =   1215
      ScaleWidth      =   855
      TabIndex        =   12
      Top             =   480
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   6
      Left            =   4200
      Picture         =   "frmSplash.frx":34A3
      ScaleHeight     =   495
      ScaleWidth      =   1455
      TabIndex        =   11
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   7080
      Picture         =   "frmSplash.frx":3CD6
      ScaleHeight     =   495
      ScaleWidth      =   1455
      TabIndex        =   10
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1320
      Picture         =   "frmSplash.frx":4509
      ScaleHeight     =   495
      ScaleWidth      =   1455
      TabIndex        =   9
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   2760
      Picture         =   "frmSplash.frx":4D3C
      ScaleHeight     =   495
      ScaleWidth      =   1455
      TabIndex        =   8
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   5640
      Picture         =   "frmSplash.frx":556F
      ScaleHeight     =   495
      ScaleWidth      =   1455
      TabIndex        =   7
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   -120
      Picture         =   "frmSplash.frx":5DA2
      ScaleHeight     =   495
      ScaleWidth      =   1455
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   5
      Left            =   7200
      Picture         =   "frmSplash.frx":65D5
      ScaleHeight     =   615
      ScaleWidth      =   1455
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   4
      Left            =   0
      Picture         =   "frmSplash.frx":D571
      ScaleHeight     =   615
      ScaleWidth      =   1455
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   3
      Left            =   5760
      Picture         =   "frmSplash.frx":1450D
      ScaleHeight     =   615
      ScaleWidth      =   1455
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   2
      Left            =   4320
      Picture         =   "frmSplash.frx":1B4A9
      ScaleHeight     =   615
      ScaleWidth      =   1455
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   1
      Left            =   2880
      Picture         =   "frmSplash.frx":22445
      ScaleHeight     =   615
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   0
      Left            =   1440
      Picture         =   "frmSplash.frx":293E1
      ScaleHeight     =   615
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Lblonly 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "This Is Only A Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   1200
      Width           =   4095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************************************************************************
'                                               EXECUTIONER By Corey McInroy
'*********************************************************************************************************
'
'    Student # 00133
'    Instructor : Paul Barretto
'    Instituion : CDI College , Brampton Campus ( Canada )
'    Compiler Used: Visual Basic 5.0
'
'    Purpose: To create a game using the skills learned. To Load a phase from disk
'                    and search through The phrase for a letter from the users input.
'
' Date: May 29/99
' *********************************************************************************************************

' *********************************************************************************************************
' Varialbles Used
' *********************************************************************************************************
Dim Number As Integer ' Used as a counter to determine when to unload FrmSplash
Dim Number2 As Integer ' Used as a random number to decide which color text will be

Private Sub Form_Load()
If LocLanguage <> "en" Then Lblonly.Caption = Loc("splash.onlyAGame")
'*********************************************************************************************************
' FORM LOAD
' *********************************************************************************************************
' Purpose: Playsound and enable Timerbegin
'
' Variables Used:Playsound -Used to play sound
'
' Function(s) Called: TimerBegin -Timer used to change text different colors
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

'if bsound = true then PlaySound "sounds\intro.wav", 0, 1
timerbegin.Enabled = True ' start the Timer ( timerbegin )

End Sub



Private Sub Time2go_Timer()

'*********************************************************************************************************
' Time2go
' *********************************************************************************************************
' Purpose: Count to 14 then unload frmsplash
'
' Variables Used: Number2 - (Counter) When = 0 frmsplash is unloaded
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

If Number2 > 24 Then Unload frmSplash
Number2 = Number2 + 1

End Sub

Private Sub timerbegin_Timer()

'*********************************************************************************************************
' Timerbegin
' *********************************************************************************************************
' Purpose: Count to 15 then call start funtion, Use random number for text color.
'
' Variables Used: Number - (Counter) When = 0 Start Funtion is Called,
'                         a - random number for color change in text
'
' Funtion(s) Called: Start
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

Dim a As Integer ' random number

If Number > 15 Then Start ' if done 15 times call the start function

a = Int((6 * Rnd) + 1) ' get random number

If a = 1 Then Lblonly.ForeColor = vbBlack
If a = 2 Then Lblonly.ForeColor = vbBlue
If a = 3 Then Lblonly.ForeColor = vbRed
If a = 4 Then Lblonly.ForeColor = vbGreen
If a = 5 Then Lblonly.ForeColor = vbCyan
If a = 6 Then Lblonly.ForeColor = vbgray

Number = Number + 1 ' increase Number by 1

End Sub

Public Function Start()

'*********************************************************************************************************
' START
' *********************************************************************************************************
' Purpose:Disable timerbegin, Start Time2go ( timer ), Showgamescreen
'
' Variables Used: NONE
'
' Funtion(s) Called:Time2go (timer enabled)
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

timerbegin.Enabled = False ' Disable the timerbegin
Time2go.Enabled = True      ' Start Time2go timer ( SEE ABOVE )
'Me.Hide
'Load frmGamescreen
frmGamescreen.Show  'Show the gamescreen ( start the game )


End Function
