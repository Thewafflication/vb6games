VERSION 5.00
Begin VB.Form frmAchievments 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   9240
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
      Left            =   3400
      TabIndex        =   30
      Top             =   2300
      Width           =   2400
   End
   Begin VB.PictureBox picAchievment 
      BorderStyle     =   0  'None
      Height          =   1030
      Index           =   1
      Left            =   1440
      Picture         =   "frmAchievments.frx":0000
      ScaleHeight     =   51.75
      ScaleMode       =   2  'Point
      ScaleWidth      =   51.75
      TabIndex        =   2
      Top             =   840
      Width           =   1030
   End
   Begin VB.PictureBox picAchievment 
      BorderStyle     =   0  'None
      Height          =   1030
      Index           =   0
      Left            =   240
      Picture         =   "frmAchievments.frx":0D6D
      ScaleHeight     =   51.75
      ScaleMode       =   2  'Point
      ScaleWidth      =   51.75
      TabIndex        =   0
      Top             =   840
      Width           =   1030
   End
   Begin VB.Label lblCaption 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   $"frmAchievments.frx":1A63
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   8175
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
Attribute VB_Name = "frmAchievments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
ThemeApply Me
cmdClose.Caption = Loc("button.done")
cmdClose.BackColor = ThAccent2
If LocLanguage <> "en" Then
    Caption = Loc("menu.achievements")
    lblCaption.Caption = Loc("ach.help")
End If
ThemeDialog Me, shpDlg, shpDlgS1, shpDlgS2, 15, 300
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCaption.Caption = Loc("ach.help")
End Sub


Private Sub picAchievment_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then lblCaption.Caption = Loc("ach.1")
If Index = 1 Then lblCaption.Caption = Loc("ach.2")
End Sub

