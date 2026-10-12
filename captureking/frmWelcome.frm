VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H8000000E&
   Caption         =   "Welcome - Capture King"
   ClientHeight    =   3330
   ClientLeft      =   4740
   ClientTop       =   3030
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4515
   Begin VB.Timer tmrLocal 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2160
      Top             =   840
   End
   Begin VB.PictureBox picLocalGame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   480
      Picture         =   "frmWelcome.frx":0000
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   960
      Width           =   1530
   End
   Begin VB.PictureBox picChat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   2520
      Picture         =   "frmWelcome.frx":7572
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   960
      Width           =   1530
   End
   Begin VB.CommandButton cmdChat 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   -120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Go to pre-game Chat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblLocalGame 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create Local Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "##NAME##"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Goingout As Boolean
Dim MyUserID As String
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Private Sub Form_Load()
lblName.Top = lblName.Top + 10
Goingout = False

frmGame.Show
frmGame.Visible = False

Dim sBuffer As String
Dim lSize As Long
sBuffer = Space$(255)
lSize = Len(sBuffer)
Call GetUserName(sBuffer, lSize)
If lSize > 0 Then
    MyUserID = Mid(Left$(sBuffer, lSize - 1), 1, 9)
    If Len(MyUserID) < 9 Then
        Do Until Len(MyUserID) = 9
            MyUserID = MyUserID & "_"
        Loop
    End If
Else
    MyUserID = vbNullString
End If
lblName = MyUserID


End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Goingout = True Then

Else
Unload frmChat
Unload frmStats
Unload frmGame
End If

End Sub

Private Sub picChat_Click()
'frmGame.Visible = False
Goingout = True
frmChat.Show
Unload Me
End Sub

Private Sub picLocalGame_Click()
frmGame.Show
Goingout = True


Unload Me

End Sub
