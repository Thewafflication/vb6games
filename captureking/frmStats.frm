VERSION 5.00
Begin VB.Form frmStats 
   BorderStyle     =   0  'None
   Caption         =   "Statistics"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1  'Graphical
      Caption         =   "OK"
      Height          =   600
      Left            =   3000
      TabIndex        =   30
      Top             =   6300
      Width           =   2400
   End
   Begin VB.TextBox txtstats 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
   Begin VB.Shape shpDlg 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpDlgS1 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpDlgS2 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wins, loses, surrenders As Integer

Private Sub Form_Load()
frmStats.Caption = LocF("stats.for", frmChat.lblUser.Caption)
ThemeApply Me
txtstats.Enabled = True
txtstats.Locked = True
txtstats.ForeColor = ThText
txtstats.Font.Size = 16
cmdClose.Caption = Loc("button.done")
cmdClose.BackColor = ThAccent2
ThemeDialog Me, shpDlg, shpDlgS1, shpDlgS2, 15, 300

Call refreshStat
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub refreshStat()
'Open App.Path & "\User\User.txt" For Input As #8
'    Input #8, C
'        userlevel = C
'        MyUserID = frmChat.lblUser
'    Input #8, C
'        wins = C
'    Input #8, C
'        loses = C
'    Input #8, C
'        surrenders = C
'Close #8
'txtstats = MyUserID & vbNewLine & _
'"User Level:" & vbTab & vbTab & userlevel & vbNewLine & _
'"Wins:" & vbTab & vbTab & vbTab & wins & vbNewLine & _
'"loses:" & vbTab & vbTab & vbTab & loses & vbNewLine & _
'"Surrenders:" & vbTab & vbTab & surrenders & vbNewLine
End Sub
