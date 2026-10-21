VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resta 1"
   ClientHeight    =   5500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBoard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5500
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione a bola e mova-a saltando sobre a bola adjacente para o local vazio. O objetivo e ficar com o menor numero de bolas."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   240
      TabIndex        =   0
      Top             =   3930
      Width           =   4815
   End
   Begin VB.Label lblCount
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   420
      Left            =   240
      TabIndex        =   40
      Top             =   4050
      Width           =   4725
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":0442
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   64
      Left            =   2820
      Picture         =   "frmBoard.frx":0884
      Top             =   3240
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":0CC6
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   63
      Left            =   2340
      Picture         =   "frmBoard.frx":1108
      Top             =   3240
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":154A
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   62
      Left            =   1860
      Picture         =   "frmBoard.frx":198C
      Top             =   3240
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":1DCE
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   54
      Left            =   2820
      Picture         =   "frmBoard.frx":2210
      Top             =   2760
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":2652
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   53
      Left            =   2340
      Picture         =   "frmBoard.frx":2A94
      Top             =   2760
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":2ED6
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   52
      Left            =   1860
      Picture         =   "frmBoard.frx":3318
      Top             =   2760
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":375A
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   46
      Left            =   3780
      Picture         =   "frmBoard.frx":3B9C
      Top             =   2280
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":3FDE
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   45
      Left            =   3300
      Picture         =   "frmBoard.frx":4420
      Top             =   2280
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":4862
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   44
      Left            =   2820
      Picture         =   "frmBoard.frx":4CA4
      Top             =   2280
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":50E6
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   43
      Left            =   2340
      Picture         =   "frmBoard.frx":5528
      Top             =   2280
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":596A
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   42
      Left            =   1860
      Picture         =   "frmBoard.frx":5DAC
      Top             =   2280
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":61EE
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   41
      Left            =   1380
      Picture         =   "frmBoard.frx":6630
      Top             =   2280
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":6A72
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   40
      Left            =   900
      Picture         =   "frmBoard.frx":6EB4
      Top             =   2280
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":72F6
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   36
      Left            =   3780
      Picture         =   "frmBoard.frx":7738
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":7B7A
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   35
      Left            =   3300
      Picture         =   "frmBoard.frx":7FBC
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":83FE
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   34
      Left            =   2820
      Picture         =   "frmBoard.frx":8840
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":8C82
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   33
      Left            =   2340
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":90C4
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   32
      Left            =   1860
      Picture         =   "frmBoard.frx":9506
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":9948
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   31
      Left            =   1380
      Picture         =   "frmBoard.frx":9D8A
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":A1CC
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   30
      Left            =   900
      Picture         =   "frmBoard.frx":A60E
      Top             =   1800
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":AA50
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   26
      Left            =   3780
      Picture         =   "frmBoard.frx":AE92
      Top             =   1320
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":B2D4
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   25
      Left            =   3300
      Picture         =   "frmBoard.frx":B716
      Top             =   1320
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":BB58
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   24
      Left            =   2820
      Picture         =   "frmBoard.frx":BF9A
      Top             =   1320
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":C3DC
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   23
      Left            =   2340
      Picture         =   "frmBoard.frx":C81E
      Top             =   1320
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":CC60
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   22
      Left            =   1860
      Picture         =   "frmBoard.frx":D0A2
      Top             =   1320
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":D4E4
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   21
      Left            =   1380
      Picture         =   "frmBoard.frx":D926
      Top             =   1320
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":DD68
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   20
      Left            =   900
      Picture         =   "frmBoard.frx":E1AA
      Top             =   1320
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":E5EC
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   14
      Left            =   2820
      Picture         =   "frmBoard.frx":EA2E
      Top             =   840
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":EE70
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   13
      Left            =   2340
      Picture         =   "frmBoard.frx":F2B2
      Top             =   840
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":F6F4
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   12
      Left            =   1860
      Picture         =   "frmBoard.frx":FB36
      Top             =   840
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":FF78
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   4
      Left            =   2820
      Picture         =   "frmBoard.frx":103BA
      Top             =   360
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":107FC
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   3
      Left            =   2340
      Picture         =   "frmBoard.frx":10C3E
      Top             =   360
      Width           =   510
   End
   Begin VB.Image imgGoti 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DragIcon        =   "frmBoard.frx":11080
      DragMode        =   1  'Automatic
      Height          =   510
      Index           =   2
      Left            =   1860
      Picture         =   "frmBoard.frx":114C2
      Top             =   360
      Width           =   510
   End
   Begin VB.Image imgGotipic 
      Height          =   480
      Left            =   60
      Picture         =   "frmBoard.frx":11904
      Top             =   5100
      Width           =   480
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   0
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   1
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   2
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   3
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   4
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   5
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   6
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   7
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   8
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   9
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   10
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   11
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   12
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   13
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   14
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   15
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   16
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   17
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   18
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   19
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   20
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   21
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   22
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   23
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   24
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   25
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   26
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   27
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   28
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   29
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   30
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   31
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpHole 
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   32
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   400
   End
   Begin VB.Shape shpBoard 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpBoardS1 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpBoardS2 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' The label's designed text is Portuguese; every language, English
    ' included, is set from the translation table.
    Caption = Loc("app.title")
    Label1.Caption = Loc("help.text")
    ThemeApply Me
    Dim c As Control, n As Long
    For Each c In Controls
        If c.Name = "imgGoti" Then
            c.BorderStyle = 0
            ' a round socket behind every place on the board
            shpHole(n).Move c.Left + 55, c.Top + 55, 400, 400
            shpHole(n).FillColor = ThShade(ThBack, 80)
            shpHole(n).BorderColor = ThLine
            shpHole(n).BorderWidth = 1
            n = n + 1
        End If
    Next c
    imgGotipic.Visible = False
    ThemeCard shpBoard, shpBoardS1, shpBoardS2, 780, 240, 3630, 3630, 15
    Label1.Font.Size = 10
    Label1.ForeColor = ThMuted
    Label1.Alignment = 2
    Label1.Move 240, 4650, 4725, 800
    lblCount.Font.Size = 16
    lblCount.Font.Bold = True
    lblCount.ForeColor = ThAccent
    UpdateCount
End Sub

Private Sub UpdateCount()
    Dim c As Control, n As Long
    For Each c In Controls
        If c.Name = "imgGoti" Then
            If c.Picture <> 0 Then n = n + 1
        End If
    Next c
    lblCount.Caption = LocF("count.text", n)
End Sub

Private Sub imgGoti_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

    If (imgGoti(Index).Picture <> 0) _
        Or Abs(Source.Index - Index) <> 2 _
       And Abs(Source.Index - Index) <> 20 Then
       Exit Sub
    End If
    
    If (imgGoti(Source.Index - ((Source.Index - Index) / 2)).Picture = 0) Then
        Exit Sub
    End If
    
    Source.Picture = LoadPicture
    imgGoti(Source.Index - ((Source.Index - Index) / 2)).Picture = LoadPicture
    imgGoti(Index).Picture = imgGotipic.Picture
    UpdateCount

End Sub

