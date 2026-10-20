VERSION 5.00
Begin VB.Form frmPawnPromotion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pawn Promotion - Capture King"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPromotepiece 
      Height          =   135
      Left            =   3960
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select this piece"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.OptionButton optSelect 
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   8
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optSelect 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optSelect 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   6
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton optSelect 
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox picboard 
      Height          =   615
      Index           =   3
      Left            =   3000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox picboard 
      Height          =   615
      Index           =   2
      Left            =   2280
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox picboard 
      Height          =   615
      Index           =   1
      Left            =   1560
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.PictureBox picboard 
      Height          =   615
      Index           =   0
      Left            =   840
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Congradulations! Your Pawn has been promoted, please  choose which piece you wihs to promote it to:"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmPawnPromotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selected As Integer
Private Sub cmdSelect_Click()
For i = 0 To 3
    If optSelect(i).Value = True Then
        selected = i
    End If
Next i

picPromotepiece.Picture = picboard(selected).Picture

frmPawnPromotion.Visible = False
Call frmGame.PawnPromotion
Unload frmPawnPromotion
End Sub

Private Sub Form_Load()
If LocLanguage <> "en" Then
    Caption = Loc("promo.title")
    cmdSelect.Caption = Loc("promo.select")
    lblTitle.Caption = Loc("promo.text")
End If
picboard(0).BackColor = &HFFFFFF
picboard(1).BackColor = &H0&
picboard(2).BackColor = &HFFFFFF
picboard(3).BackColor = &H0&

For i = 0 To 3
    optSelect(i).BackColor = frmGame.CommonDialog1.Color
Next i

optSelect(3).Value = True

If frmGame.lblColor.Caption = "white" Then
    picboard(0) = LoadPicture(App.Path & "\img\w2.ico")
    picboard(1) = LoadPicture(App.Path & "\img\w3.ico")
    picboard(2) = LoadPicture(App.Path & "\img\w4.ico")
    picboard(3) = LoadPicture(App.Path & "\img\w5.ico")
Else
    picboard(0) = LoadPicture(App.Path & "\img\b2.ico")
    picboard(1) = LoadPicture(App.Path & "\img\b3.ico")
    picboard(2) = LoadPicture(App.Path & "\img\b4.ico")
    picboard(3) = LoadPicture(App.Path & "\img\b5.ico")
End If

End Sub

