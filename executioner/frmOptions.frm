VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Rulez"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optTimer 
      BackColor       =   &H00000000&
      Caption         =   "Timer off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   4
      Left            =   720
      TabIndex        =   15
      Top             =   4320
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   3000
      Picture         =   "frmOptions.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   2775
      TabIndex        =   12
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   3480
      MousePointer    =   2  'Cross
      Picture         =   "frmOptions.frx":0ABB
      ScaleHeight     =   375
      ScaleWidth      =   1335
      TabIndex        =   11
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame FrameTimeSettings 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3135
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
      Begin VB.OptionButton optTimer 
         BackColor       =   &H00000000&
         Caption         =   "30 sec. (Knight)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   2895
      End
      Begin VB.OptionButton optTimer 
         BackColor       =   &H00000000&
         Caption         =   "15 sec. (King)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   390
         Index           =   3
         Left            =   360
         TabIndex        =   9
         Top             =   2040
         Width           =   2415
      End
      Begin VB.OptionButton optTimer 
         BackColor       =   &H00000000&
         Caption         =   "45 sec. (Adventurer)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton optTimer 
         BackColor       =   &H00000000&
         Caption         =   "60 sec. (Squire)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Index           =   0
         Left            =   360
         MaskColor       =   &H0000C000&
         TabIndex        =   7
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame FrameHintSettings 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   5280
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
      Begin VB.OptionButton optHints 
         BackColor       =   &H00000000&
         Caption         =   "Cheater"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   2400
         Width           =   1455
      End
      Begin VB.OptionButton optHints 
         BackColor       =   &H00000000&
         Caption         =   "5 Hints"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optHints 
         BackColor       =   &H00000000&
         Caption         =   "2 Hints"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optHints 
         BackColor       =   &H00000000&
         Caption         =   "1 Hint"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      Picture         =   "frmOptions.frx":0F7A
      ScaleHeight     =   1215
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   600
      Width           =   8055
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         Picture         =   "frmOptions.frx":1EB7
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   14
         Top             =   720
         Width           =   2535
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   4680
         Picture         =   "frmOptions.frx":28A7
         ScaleHeight     =   495
         ScaleWidth      =   2535
         TabIndex        =   13
         Top             =   720
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmOptions"
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

Private Sub ApplyLocale()
    If LocLanguage = "en" Then Exit Sub
    Caption = Loc("options.title")
    optTimer(4).Caption = Loc("options.timerOff")
    optTimer(0).Caption = Loc("options.time60")
    optTimer(1).Caption = Loc("options.time45")
    optTimer(2).Caption = Loc("options.time30")
    optTimer(3).Caption = Loc("options.time15")
    optHints(0).Caption = Loc("options.hint1")
    optHints(1).Caption = Loc("options.hint2")
    optHints(2).Caption = Loc("options.hint5")
    optHints(3).Caption = Loc("options.cheater")
End Sub

Private Sub Form_Load()
    ApplyLocale
End Sub

Public Sub Picture2_Click(Index As Integer)

' *********************************************************************************************************
' Picture2_click
' *********************************************************************************************************
' Purpose:When clicked the options chosen are saved and form is unloaded
'
' Global Variables Used: Timesave - Saves the value of the old time in case of problems
'                                    Hints -Number of Hints for 1 player game
'                                    Hints2 - Number of hints for 2 player game
'                                    TimeOff - Determines if the timer is on or off
'
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

 If TimeOff = True Then TimeOff = False 'If Time is off turn it on
 
 If optTimer(0).Value = True Then TimeSave = 60 ' Time is set to 60 seconds if chosen
 If optTimer(1).Value = True Then TimeSave = 45
  If optTimer(2).Value = True Then TimeSave = 30
 If optTimer(3).Value = True Then TimeSave = 15

If optTimer(4).Value = True Then TimeOff = True ' turn sound off if chosen

If optHints(0).Value = True Then ' One Hint set for both 1 player and two player
Hints = 1
Hints2 = 1
End If

If optHints(1).Value = True Then
Hints = 2
Hints2 = 2
End If

If optHints(2).Value = True Then
Hints = 5
Hints2 = 5
End If

If optHints(3).Value = True Then
Hints = 666
Hints2 = 666
End If

Unload frmOptions

End Sub
