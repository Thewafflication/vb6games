VERSION 5.00
Begin VB.Form frmStats 
   Caption         =   "Statistics"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtstats 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   8535
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

Call refreshStat
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
