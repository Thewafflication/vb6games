VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   LinkTopic       =   "Form2"
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   5160
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5880
      Top             =   3360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "(c) Christopher McKirgan 2007 - Capture King"
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   4920
      Width           =   3255
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmWelcome.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
frmWelcome.Show
Unload Me
End Sub
