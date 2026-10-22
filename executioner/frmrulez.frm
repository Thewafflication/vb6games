VERSION 5.00
Begin VB.Form frmrulez 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Rulez"
   ClientHeight    =   5265
   ClientLeft      =   1560
   ClientTop       =   1380
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicOptGo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2880
      MousePointer    =   2  'Cross
      Picture         =   "frmrulez.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   2775
      TabIndex        =   4
      Top             =   4440
      Width           =   2775
   End
   Begin VB.PictureBox PicoptTitle 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2760
      Picture         =   "frmrulez.frx":09AF
      ScaleHeight     =   855
      ScaleWidth      =   3255
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      Picture         =   "frmrulez.frx":1734
      ScaleHeight     =   495
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.PictureBox picRulez 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   720
      Picture         =   "frmrulez.frx":1CEC
      ScaleHeight     =   2775
      ScaleWidth      =   6615
      TabIndex        =   1
      Top             =   1320
      Width           =   6615
   End
   Begin VB.PictureBox PicOptBanner 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      Picture         =   "frmrulez.frx":76AA
      ScaleHeight     =   1215
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   600
      Width           =   8055
   End
End
Attribute VB_Name = "frmrulez"
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
    ThemeCenter Me, 0
If LocLanguage <> "en" Then Caption = Loc("rulez.title")
End Sub

Private Sub PicOptGo_Click()
' ******************************************************************
' When picture Clicked unload frmRulez
' ******************************************************************

Unload frmrulez

End Sub
