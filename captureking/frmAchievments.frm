VERSION 5.00
Begin VB.Form frmAchievments 
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
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCaption.Caption = "To unlock achievments you must perform certain actions during chess matches. To see view how to unlock achievments place your cursor over the images below."
End Sub


Private Sub picAchievment_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then lblCaption.Caption = "Rewarded for winning a game of chess through the LAN network"
If Index = 1 Then lblCaption.Caption = "Rewarded for winning 50 games of chess through the LAN network"
End Sub

