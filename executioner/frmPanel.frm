VERSION 5.00
Begin VB.Form frmPanel
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Panel"
   ClientHeight    =   5600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   ScaleHeight     =   5600
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape shpEdge
      BorderColor     =   &H0054B4FF&
      BorderWidth     =   3
      Height          =   5500
      Left            =   50
      Top             =   50
      Width           =   4100
   End
   Begin VB.CommandButton cmdDone
      Style           =   1  'Graphical
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   13
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OK"
      Height          =   560
      Left            =   300
      TabIndex        =   9
      Top             =   4800
      Width           =   3600
   End
   Begin VB.CommandButton cmdItem 
      Style           =   1  'Graphical
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   13
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "item0"
      Height          =   560
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   800
      Width           =   3600
   End
   Begin VB.CommandButton cmdItem 
      Style           =   1  'Graphical
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   13
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "item1"
      Height          =   560
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Top             =   1440
      Width           =   3600
   End
   Begin VB.CommandButton cmdItem 
      Style           =   1  'Graphical
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   13
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "item2"
      Height          =   560
      Index           =   2
      Left            =   300
      TabIndex        =   3
      Top             =   2080
      Width           =   3600
   End
   Begin VB.CommandButton cmdItem 
      Style           =   1  'Graphical
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   13
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "item3"
      Height          =   560
      Index           =   3
      Left            =   300
      TabIndex        =   4
      Top             =   2720
      Width           =   3600
   End
   Begin VB.CommandButton cmdItem 
      Style           =   1  'Graphical
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   13
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "item4"
      Height          =   560
      Index           =   4
      Left            =   300
      TabIndex        =   5
      Top             =   3360
      Width           =   3600
   End
   Begin VB.CommandButton cmdItem 
      Style           =   1  'Graphical
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   13
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "item5"
      Height          =   560
      Index           =   5
      Left            =   300
      TabIndex        =   6
      Top             =   4000
      Width           =   3600
   End
   Begin VB.Label lblTitle
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   16
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "title"
      ForeColor       =   &H0054B4FF&
      Height          =   500
      Left            =   300
      TabIndex        =   0
      Top             =   200
      Width           =   3600
   End
End
Attribute VB_Name = "frmPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' A floating panel of buttons, shown over the game in place of a menu.
Dim mWhich As String
Dim mCount As Integer

Public Sub OpenPanel(ByVal which As String)
    mWhich = which
    Setup
    Show 1
End Sub

Private Sub Setup()
    Dim i As Integer
    For i = 0 To 5
        cmdItem(i).BackColor = &HC6D63D
    Next i
    Select Case mWhich
    Case "game"
        lblTitle.Caption = Replace(Loc("menu.game"), "&", "")
        mCount = 4
        cmdItem(0).Caption = Replace(Loc("menu.newGame"), "&", "")
        cmdItem(1).Caption = Loc("menu.endQuest")
        cmdItem(2).Caption = Loc("menu.onePlayer")
        cmdItem(3).Caption = Loc("menu.twoPlayers")
        cmdItem(2).BackColor = IIf(frmGamescreen.mnu1player.Checked, &H54B4FF, &HC6D63D)
        cmdItem(3).BackColor = IIf(frmGamescreen.mnu2player.Checked, &H54B4FF, &HC6D63D)
    Case "options"
        lblTitle.Caption = Replace(Loc("menu.options"), "&", "")
        mCount = 5
        cmdItem(0).Caption = Replace(Loc("menu.sound"), "&", "")
        cmdItem(1).Caption = Loc("menu.showTimer")
        cmdItem(2).Caption = Loc("menu.showHints")
        cmdItem(3).Caption = Loc("menu.setNames")
        cmdItem(4).Caption = Replace(Loc("menu.advanced"), "&", "")
        cmdItem(0).BackColor = IIf(frmGamescreen.mnuSound.Checked, &H54B4FF, &HC6D63D)
        cmdItem(1).BackColor = IIf(frmGamescreen.mnuShowTimer.Checked, &H54B4FF, &HC6D63D)
        cmdItem(2).BackColor = IIf(frmGamescreen.mnuHints.Checked, &H54B4FF, &HC6D63D)
    Case Else
        lblTitle.Caption = Replace(Loc("menu.about"), "&", "")
        mCount = 2
        cmdItem(0).Caption = Replace(Loc("menu.rules"), "&", "")
        cmdItem(1).Caption = Loc("menu.creator")
    End Select
    For i = 0 To 5
        cmdItem(i).Visible = (i < mCount)
    Next i
    cmdDone.Caption = Loc("button.done")
    cmdDone.BackColor = &H816BFF
    cmdDone.Top = 800 + mCount * 640 + 100
    Height = cmdDone.Top + 560 + 300
    shpEdge.Height = Height - 100
    ' centered in the screen (Screen is in the program's own units)
    Dim px As Single, py As Single
    px = (Screen.Width - Width) / 2
    py = (Screen.Height - Height) / 2
    Me.Move px, py
End Sub

Private Sub cmdItem_Click(Index As Integer)
    Dim g As String
    g = mWhich
    Unload Me
    Select Case g
    Case "game"
        Select Case Index
        Case 0
            frmGamescreen.mnuStartGame_Click
        Case 1
            frmGamescreen.mnuEndGame_Click
        Case 2
            frmGamescreen.mnu1player_Click
        Case 3
            frmGamescreen.mnu2player_Click
        End Select
    Case "options"
        Select Case Index
        Case 0
            frmGamescreen.mnuSound_Click
        Case 1
            frmGamescreen.mnuShowTimer_Click
        Case 2
            frmGamescreen.mnuHints_Click
        Case 3
            frmGamescreen.mnuGetNames_Click
        Case 4
            frmGamescreen.mnuAdvancedOptions_Click
        End Select
    Case Else
        Select Case Index
        Case 0
            frmGamescreen.mnuRulez_Click
        Case 1
            frmGamescreen.mnuCreator_Click
        End Select
    End Select
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub
