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
End
Attribute VB_Name = "frmAchievments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ThemeApply Me
If LocLanguage <> "en" Then
    Caption = Loc("menu.achievements")
    lblCaption.Caption = Loc("ach.help")
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCaption.Caption = Loc("ach.help")
End Sub


Private Sub picAchievment_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then lblCaption.Caption = Loc("ach.1")
If Index = 1 Then lblCaption.Caption = Loc("ach.2")
End Sub

