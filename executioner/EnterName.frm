VERSION 5.00
Begin VB.Form frmEnterName 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   11355
   ClientTop       =   3945
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picNameBegin 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      MousePointer    =   2  'Cross
      Picture         =   "EnterName.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   2775
      TabIndex        =   10
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtPlayerNameTwo 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   600
      MaxLength       =   7
      TabIndex        =   2
      ToolTipText     =   "Enter Name"
      Top             =   3600
      Width           =   4575
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   480
      Picture         =   "EnterName.frx":0948
      ScaleHeight     =   495
      ScaleWidth      =   2415
      TabIndex        =   9
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txtPlayerNameOne 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   600
      MaxLength       =   7
      TabIndex        =   1
      ToolTipText     =   "Enter Name"
      Top             =   1920
      Width           =   4575
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   480
      Picture         =   "EnterName.frx":0F2E
      ScaleHeight     =   615
      ScaleWidth      =   1815
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   840
      Picture         =   "EnterName.frx":14F8
      ScaleHeight     =   375
      ScaleWidth      =   4695
      TabIndex        =   7
      Top             =   960
      Width           =   4695
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1560
      Picture         =   "EnterName.frx":186A
      ScaleHeight     =   855
      ScaleWidth      =   3375
      TabIndex        =   6
      Top             =   240
      Width           =   3375
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   5760
      Picture         =   "EnterName.frx":2453
      ScaleHeight     =   4695
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      Picture         =   "EnterName.frx":2DC7
      ScaleHeight     =   4695
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      Picture         =   "EnterName.frx":373B
      ScaleHeight     =   495
      ScaleWidth      =   6135
      TabIndex        =   3
      Top             =   4920
      Width           =   6135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      Picture         =   "EnterName.frx":3D5F
      ScaleHeight     =   495
      ScaleWidth      =   6135
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmEnterName"
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

Private Sub Form_Load()
' *********************************************************************************************************
' FORM LOAD
' *********************************************************************************************************
' Purpose: To play a Sound when form is loaded.
'
' Variables Used: Playsound ( used to play sound )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

If bSound = True Then PlaySound "sounds\name.wav", 0, 1

End Sub

Private Sub picNameBegin_Click()

' *********************************************************************************************************
' picNameBegin
' *********************************************************************************************************
' Purpose:When clicked the Name is saved
'
' Global Variables Used: Player1name, Player2name, (names of the players )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

If txtPlayerNameOne = "" Then txtPlayerNameOne = "Player 1" ' set default names if empty
If txtPlayerNameTwo = "" Then txtPlayerNameTwo = "Player 2"

Player1name = txtPlayerNameOne ' text typed in by User is now Player1name
Player2name = txtPlayerNameTwo

On Error Resume Next
frmGamescreen.lblPlayerName(0) = Player1name
frmGamescreen.lblPlayerName(1) = Player2name


Unload frmEnterName ' unload form

End Sub

