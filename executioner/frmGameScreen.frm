VERSION 5.00
Begin VB.Form frmGamescreen 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Executioner By Corey McInroy                                               "
   ClientHeight    =   9200
   ClientLeft      =   7860
   ClientTop       =   2280
   ClientWidth     =   11880
   DrawStyle       =   2  'Dot
   FillColor       =   &H00FFFF00&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmGameScreen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MouseIcon       =   "frmGameScreen.frx":0442
   MousePointer    =   2  'Cross
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9200
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicWinnertext 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   4800
      Picture         =   "frmGameScreen.frx":0D0C
      ScaleHeight     =   855
      ScaleWidth      =   2415
      TabIndex        =   60
      Top             =   3600
      Width           =   2415
   End
   Begin VB.PictureBox picLightning 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5700
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   66
      Top             =   1100
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picClear 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   2280
      ScaleHeight     =   1815
      ScaleWidth      =   7455
      TabIndex        =   65
      Top             =   6480
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.PictureBox PicWinnertext 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   0
      Left            =   5320
      Picture         =   "frmGameScreen.frx":1643
      ScaleHeight     =   900
      ScaleWidth      =   3090
      TabIndex        =   59
      Top             =   2540
      Width           =   3090
   End
   Begin VB.TextBox txtSolve1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   480
      Left            =   3360
      MaxLength       =   50
      TabIndex        =   3
      ToolTipText     =   "Type The phrase then press ENTER"
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Timer TimerSolve 
      Interval        =   200
      Left            =   1800
      Top             =   840
   End
   Begin VB.PictureBox PicWinner 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   6960
      ScaleHeight     =   2295
      ScaleWidth      =   1935
      TabIndex        =   58
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Timer TimerWinner 
      Left            =   1800
      Top             =   840
   End
   Begin VB.Timer timerDeath 
      Interval        =   5
      Left            =   1800
      Top             =   840
   End
   Begin VB.CommandButton cmdBarNew 
      Style           =   1  'Graphical
      Caption         =   "cmdBarNew"
      Height          =   560
      Left            =   1390
      TabIndex        =   70
      Top             =   8450
      Width           =   1900
   End
   Begin VB.CommandButton cmdBarGame 
      Style           =   1  'Graphical
      Caption         =   "cmdBarGame"
      Height          =   560
      Left            =   3440
      TabIndex        =   71
      Top             =   8450
      Width           =   1900
   End
   Begin VB.CommandButton cmdBarOptions 
      Style           =   1  'Graphical
      Caption         =   "cmdBarOptions"
      Height          =   560
      Left            =   5490
      TabIndex        =   72
      Top             =   8450
      Width           =   1900
   End
   Begin VB.CommandButton cmdBarAbout 
      Style           =   1  'Graphical
      Caption         =   "cmdBarAbout"
      Height          =   560
      Left            =   7540
      TabIndex        =   73
      Top             =   8450
      Width           =   1900
   End
   Begin VB.CommandButton cmdBarExit 
      Style           =   1  'Graphical
      Caption         =   "cmdBarExit"
      Height          =   560
      Left            =   9590
      TabIndex        =   74
      Top             =   8450
      Width           =   1900
   End
   Begin VB.CommandButton cmdSolve2 
      Caption         =   "Solve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   56
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton cmdSolve1 
      Caption         =   "Solve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2280
      TabIndex        =   55
      Top             =   7080
      Width           =   855
   End
   Begin VB.PictureBox PicPlayer2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   7920
      Picture         =   "frmGameScreen.frx":A7D5
      ScaleHeight     =   3495
      ScaleWidth      =   1575
      TabIndex        =   54
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PicPlayer1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2280
      Picture         =   "frmGameScreen.frx":BA0D
      ScaleHeight     =   3495
      ScaleWidth      =   1695
      TabIndex        =   53
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox PicSkeleton 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   3615
      Left            =   4800
      ScaleHeight     =   3615
      ScaleWidth      =   2775
      TabIndex        =   52
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Timer TimerIdle 
      Interval        =   69
      Left            =   1800
      Tag             =   "1"
      Top             =   840
   End
   Begin VB.Timer TimerGameOver 
      Interval        =   20
      Left            =   1800
      Top             =   840
   End
   Begin VB.PictureBox picGameOver 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   5520
      ScaleHeight     =   4335
      ScaleWidth      =   1575
      TabIndex        =   51
      Top             =   1080
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   0
      Picture         =   "frmGameScreen.frx":CC45
      ScaleHeight     =   1215
      ScaleWidth      =   2175
      TabIndex        =   48
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmdHint 
      BackColor       =   &H00404080&
      Caption         =   "HINT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   8880
      Picture         =   "frmGameScreen.frx":D681
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmdHint 
      BackColor       =   &H00404080&
      Caption         =   "HINT"
      DisabledPicture =   "frmGameScreen.frx":E690
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2280
      Picture         =   "frmGameScreen.frx":E99B
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7680
      Width           =   855
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   9480
      Picture         =   "frmGameScreen.frx":F9AA
      ScaleHeight     =   855
      ScaleWidth      =   375
      TabIndex        =   44
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   2160
      Picture         =   "frmGameScreen.frx":1031E
      ScaleHeight     =   855
      ScaleWidth      =   375
      TabIndex        =   43
      Top             =   0
      Width           =   375
      Begin VB.Timer TimerPAUSE 
         Interval        =   300
         Left            =   240
         Top             =   240
      End
      Begin VB.Timer TimerSolve2player 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   360
         Top             =   240
      End
   End
   Begin VB.PictureBox picDude1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   5280
      Picture         =   "frmGameScreen.frx":10C92
      ScaleHeight     =   3615
      ScaleWidth      =   1575
      TabIndex        =   42
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Picture         =   "frmGameScreen.frx":11ECA
      ScaleHeight     =   735
      ScaleWidth      =   3015
      TabIndex        =   41
      Top             =   0
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Picture         =   "frmGameScreen.frx":12C1D
      ScaleHeight     =   375
      ScaleWidth      =   7815
      TabIndex        =   40
      Top             =   720
      Width           =   7815
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "Z"
      DownPicture     =   "frmGameScreen.frx":14A8A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   7560
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":14F6C
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7680
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "Y"
      DownPicture     =   "frmGameScreen.frx":1547A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   6960
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1595C
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "X"
      DownPicture     =   "frmGameScreen.frx":15E6A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   6360
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1634C
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "W"
      DownPicture     =   "frmGameScreen.frx":1685A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   5760
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":16D3C
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "V"
      DownPicture     =   "frmGameScreen.frx":1724A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   5160
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1772C
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "U"
      DownPicture     =   "frmGameScreen.frx":17C3A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   4560
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1811C
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "T"
      DownPicture     =   "frmGameScreen.frx":1862A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   3960
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":18B0C
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7680
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "S"
      DownPicture     =   "frmGameScreen.frx":1901A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   8160
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":194FC
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "R"
      DownPicture     =   "frmGameScreen.frx":19A0A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   7560
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":19EEC
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "Q"
      DownPicture     =   "frmGameScreen.frx":1A3FA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   6960
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1A8DC
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "P"
      DownPicture     =   "frmGameScreen.frx":1ADEA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   6360
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1B2CC
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "O"
      DownPicture     =   "frmGameScreen.frx":1B7DA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   5760
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1BCBC
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "N"
      DownPicture     =   "frmGameScreen.frx":1C1CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   5160
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1C6AC
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "M"
      DownPicture     =   "frmGameScreen.frx":1CBBA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   4560
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1D09C
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "L"
      DownPicture     =   "frmGameScreen.frx":1D5AA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   3960
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1DA8C
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "K"
      DownPicture     =   "frmGameScreen.frx":1DF9A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   3360
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1E47C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "J"
      DownPicture     =   "frmGameScreen.frx":1E98A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   8400
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1EE6C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "I"
      DownPicture     =   "frmGameScreen.frx":1F37A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   7800
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":1F85C
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "H"
      DownPicture     =   "frmGameScreen.frx":1FD6A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   7200
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":2024C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "G"
      DownPicture     =   "frmGameScreen.frx":2075A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   6600
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":20C3C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "F"
      DownPicture     =   "frmGameScreen.frx":2114A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6000
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":2162C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "E"
      DownPicture     =   "frmGameScreen.frx":21B3A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   5400
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":2201C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "D"
      DownPicture     =   "frmGameScreen.frx":2252A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   4800
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":22A0C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "C"
      DownPicture     =   "frmGameScreen.frx":22F1A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4200
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":233FC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "B"
      DownPicture     =   "frmGameScreen.frx":2390A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3600
      MaskColor       =   &H00C000C0&
      Picture         =   "frmGameScreen.frx":23DEC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdLetter 
      BackColor       =   &H00000000&
      Caption         =   "A"
      DownPicture     =   "frmGameScreen.frx":242FA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3000
      MaskColor       =   &H000000FF&
      Picture         =   "frmGameScreen.frx":247DC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   14
      Left            =   9240
      Picture         =   "frmGameScreen.frx":24CEA
      ScaleHeight     =   975
      ScaleWidth      =   2775
      TabIndex        =   14
      Top             =   0
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   13
      Left            =   9240
      Picture         =   "frmGameScreen.frx":257A4
      ScaleHeight     =   1335
      ScaleWidth      =   2775
      TabIndex        =   13
      Top             =   2040
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   12
      Left            =   9240
      Picture         =   "frmGameScreen.frx":2625E
      ScaleHeight     =   1335
      ScaleWidth      =   2775
      TabIndex        =   12
      Top             =   3360
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   11
      Left            =   9240
      Picture         =   "frmGameScreen.frx":26D18
      ScaleHeight     =   1095
      ScaleWidth      =   2775
      TabIndex        =   11
      Top             =   960
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   8
      Left            =   9240
      Picture         =   "frmGameScreen.frx":277D2
      ScaleHeight     =   1335
      ScaleWidth      =   2775
      TabIndex        =   10
      Top             =   4680
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   10
      Left            =   9240
      Picture         =   "frmGameScreen.frx":2828C
      ScaleHeight     =   1335
      ScaleWidth      =   2775
      TabIndex        =   9
      Top             =   6000
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   9
      Left            =   9240
      Picture         =   "frmGameScreen.frx":28D46
      ScaleHeight     =   1335
      ScaleWidth      =   2775
      TabIndex        =   8
      Top             =   7320
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   6
      Left            =   0
      Picture         =   "frmGameScreen.frx":29800
      ScaleHeight     =   1335
      ScaleWidth      =   2775
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
      Begin VB.Timer TmrLightning 
         Interval        =   200
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   3
      Left            =   0
      Picture         =   "frmGameScreen.frx":2A23C
      ScaleHeight     =   1335
      ScaleWidth      =   2775
      TabIndex        =   6
      Top             =   6000
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   4
      Left            =   0
      Picture         =   "frmGameScreen.frx":2AC78
      ScaleHeight     =   1335
      ScaleWidth      =   2775
      TabIndex        =   5
      Top             =   7320
      Width           =   2775
      Begin VB.Timer Timer 
         Interval        =   1000
         Left            =   240
         Top             =   360
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   0
      Picture         =   "frmGameScreen.frx":2B6B4
      ScaleHeight     =   1095
      ScaleWidth      =   2775
      TabIndex        =   4
      Top             =   4920
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   0
      Picture         =   "frmGameScreen.frx":2C0F0
      ScaleHeight     =   1335
      ScaleWidth      =   2775
      TabIndex        =   2
      Top             =   2520
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   0
      Picture         =   "frmGameScreen.frx":2CB2C
      ScaleHeight     =   1095
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label LblWinnerbox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   64
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label LblWinnerbox 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   735
      Index           =   0
      Left            =   3600
      TabIndex        =   63
      Top             =   6120
      Width           =   4815
   End
   Begin VB.Label lblPlayerName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Player 2 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   1
      Left            =   7440
      TabIndex        =   62
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblPlayerName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   61
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblEndPhrase 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   855
      Left            =   3120
      TabIndex        =   57
      Top             =   6840
      Width           =   5895
   End
   Begin VB.Label lblWinLose 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   4080
      TabIndex        =   50
      Top             =   5520
      Width           =   3975
   End
   Begin VB.Label lblGamePhrase 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1575
      Left            =   2880
      TabIndex        =   49
      Top             =   1440
      Width           =   6375
   End
   Begin VB.Label lblTimer1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5480
      TabIndex        =   47
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Menu munGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuStartGame 
         Caption         =   "Start &New Game"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuEndGame 
         Caption         =   "End Quest"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu1player 
         Caption         =   "One Player"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu2player 
         Caption         =   "Two Players"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSound 
         Caption         =   "&Sound"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuShowTimer 
         Caption         =   "Show Timer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuHints 
         Caption         =   "Show Hints"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGetNames 
         Caption         =   "Set Player Names"
      End
      Begin VB.Menu mnuAdvancedOptions 
         Caption         =   "A&dvanced Options"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      WindowList      =   -1  'True
      Begin VB.Menu mnuRulez 
         Caption         =   "&Rulez"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuCreator 
         Caption         =   "About the Creator"
      End
   End
End
Attribute VB_Name = "frmGamescreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************************************************************************
'                                               EXECUTIONER By Corey McInroy
'*********************************************************************************************************
'
'    Student # 00133
'    Instructor : Paul Barretto
'    Instituion : CDI College , Brampton Campus ( Canada )
'    Compiler Used: Visual Basic 5.0
'
'    Purpose: To create a game using the skills learned. To Load a phase from disk
'                    and search through The phrase for a letter from the users input.
'
' Date: May 29/99
' *********************************************************************************************************


' *********************************************************************************************************
' Variables Used in frmGamescreen
' ***************************************************************************************************
  Dim randomSound As Integer ' picks random number for random sound
  Dim GameCheck As Integer ' Checks if Phrase is the same as the old one
  Public DeltaX, DeltaY As Integer ' Used for winner grafix
  Public Turn1player As Integer  ' keeps track of player 1's turn in one player game
  Dim EndtimeWinner, PauseTime, Hangcounter, Endtime As Integer ' Counters
  Dim skel As Integer, iChange As Integer, bIdle As Boolean
  Dim bRandsound As Boolean
  Dim iLight As Integer
  Dim bothdead As Boolean ' Used to Tell if both players are Dead in 2 player Game

Private Sub cmdLetter_Click(Index As Integer)

Dim Z As Integer

' *********************************************************************************************************
' cmdLetter (on click)
' *********************************************************************************************************
' Purpose:Gets Letter From Player and Sends it to checkletter, To see if its in phrase
'
' Variables Letter - The letter which the user clicked on, Z - ( counter )
'
' Function(s) Called: CheckLetter (Funtion to check if letter in phrase)
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

If txtSolve1.Visible = True Then ' the the solve text box is visible, clear and hide.
    txtSolve1.Visible = False
     txtSolve1 = ""
    End If

For Z = 0 To 25
If cmdLetter(Z) Then Letter = Chr$(97 + Z) ' checks for letter
Next
Z = 0

For Z = 0 To 25
If cmdLetter(Z) Then cmdLetter(Z).Visible = False ' hide the button that was clicked
Next

CheckLetter ' Goto the check letter funtion

End Sub

Private Sub cmdHint_Click(Index As Integer)

' *********************************************************************************************************
' cmdHint (on click)
' *********************************************************************************************************
' Purpose: Find out if the player has any hints left and if so Call 'GetHint' to put the
'               Letter in the Phrase. If the Solve textbox is visible, set visible to false
'
' Variables: Hints -Amount of hints Left for Player 1, Playesound - Used to Playsound
'                 Hints2 - Amount of Hints for Player 2
'
' Function(s) Called: GetHint ( used to put a letter in the phrase )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
    
    If txtSolve1.Visible = True Then ' the the solve text box is visible, clear and hide.
    txtSolve1.Visible = False
    txtSolve1 = ""
    End If
    
    If cmdHint(0) Then ' If Player 1 hint  button was clicked
        If Hints > 0 Then GetHint ' if player 1 has hints remaining Call GetHint Function
           Hints = Hints - 1 ' Decrement hints for player 1 by 1
           If Hints = 0 Then cmdHint(0).Visible = False ' if no hints left hide the hints button
    End If
     
    If cmdHint(1) Then ' if Player 2 hint button was clicked
        If Hints2 > 0 Then GetHint ' if player 2 has hints remaining Call GetHint Function
           Hints2 = Hints2 - 1 ' Decrement hints for player 2 by 1
           If Hints2 = 0 Then cmdHint(1).Visible = False ' if no hints left hide the hints button
    End If
     
   If bSound = True Then PlaySound "sounds\cheat\cheater.wav", 0, 1

End Sub

Private Sub cmdSolve1_Click()
' *********************************************************************************************************
' cmdSolve1 (on click) For Player 1
' *********************************************************************************************************
' Purpose: If the Solve Button is Clicked Pop up the Textbox for user to type phrase
'
' Variables: NONE
'
' Function(s) Called: TimerSolve (timer) - Checks the Phrase
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

txtSolve1.Visible = True ' pops up text box when Solve button clicked
txtSolve1.SetFocus ' Make it active for typing
TimerSolve.Enabled = True ' Start TimerSolve to check Phrase

End Sub

Private Sub cmdSolve2_Click()

' *********************************************************************************************************
' cmdSolve1 (on click) For Player 2
' *********************************************************************************************************
' Purpose: If the Solve Button is Clicked Pop up the Textbox for user to type phrase
'
' Variables: NONE
'
' Function(s) Called: TimerSolve (timer) - Checks the Phrase
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

txtSolve1.Visible = True ' pops up text box when Solve button clicked
txtSolve1.SetFocus ' Make it active for typing
TimerSolve.Enabled = True ' Start TimerSolve to check Phrase

End Sub

Private Sub Form_Click()
Dim x As Integer
x = Int((2 * Rnd) + 1)
If bIdle = True Then
bIdle = False

ClearGrafix 'Calls Function to Clear grafics and restore settings

If x = 1 Then
    If bSound = True Then PlaySound "sounds\begin.wav", 0, 1
End If

If x = 2 Then
 If bSound = True Then PlaySound "sounds\begin2.wav", 0, 1
End If


If mnu1player.Checked = True Then Players = 1 ' if checked players = 1
If mnu1player.Checked = True Then Set1playerGame ' if checked call funtion
If mnu2player.Checked = True Then Players = 2
If mnu2player.Checked = True Then Set2PlayerGame
End If

End Sub

Private Sub Form_DblClick()

Dim x As Integer
x = Int((2 * Rnd) + 1)
If bIdle = True Then
bIdle = False

ClearGrafix 'Calls Function to Clear grafics and restore settings

If x = 1 Then
    If bSound = True Then PlaySound "sounds\begin.wav", 0, 1
End If

If x = 2 Then
 If bSound = True Then PlaySound "sounds\begin2.wav", 0, 1
End If

If mnu1player.Checked = True Then Players = 1 ' if checked players = 1
If mnu1player.Checked = True Then Set1playerGame ' if checked call funtion
If mnu2player.Checked = True Then Players = 2
If mnu2player.Checked = True Then Set2PlayerGame
End If

End Sub

Private Sub ApplyLocale()
    If LocLanguage = "en" Then Exit Sub
    Caption = Loc("app.title")
    txtSolve1.ToolTipText = Loc("tip.solve")
    cmdSolve1.Caption = Loc("button.solve")
    cmdSolve2.Caption = Loc("button.solve")
    cmdHint(0).Caption = Loc("button.hint")
    cmdHint(1).Caption = Loc("button.hint")
    lblPlayerName(0).Caption = Loc("player.1")
    lblPlayerName(1).Caption = Loc("player.2")
    munGame.Caption = Loc("menu.game")
    mnuStartGame.Caption = Loc("menu.newGame")
    mnuEndGame.Caption = Loc("menu.endQuest")
    mnu1player.Caption = Loc("menu.onePlayer")
    mnu2player.Caption = Loc("menu.twoPlayers")
    mnuExit.Caption = Loc("menu.exit")
    mnuOptions.Caption = Loc("menu.options")
    mnuSound.Caption = Loc("menu.sound")
    mnuShowTimer.Caption = Loc("menu.showTimer")
    mnuHints.Caption = Loc("menu.showHints")
    mnuGetNames.Caption = Loc("menu.setNames")
    mnuAdvancedOptions.Caption = Loc("menu.advanced")
    mnuAbout.Caption = Loc("menu.about")
    mnuRulez.Caption = Loc("menu.rules")
    mnuCreator.Caption = Loc("menu.creator")
End Sub

' Touch replacements for the menu bar: each button does what its menu does
' (the menus themselves stay, as pop-ups).
Private Sub BarButtons()
    Dim c As Control
    cmdBarNew.Caption = Replace(Loc("menu.newGame"), "&", "")
    cmdBarGame.Caption = Replace(Loc("menu.game"), "&", "") & " ..."
    cmdBarOptions.Caption = Replace(Loc("menu.options"), "&", "") & " ..."
    cmdBarAbout.Caption = Replace(Loc("menu.about"), "&", "") & " ..."
    cmdBarExit.Caption = Replace(Loc("menu.exit"), "&", "")
    For Each c In Controls
        If Left(c.Name, 6) = "cmdBar" Then
            c.BackColor = ThAccent2
            c.Font.Size = 14
            c.Font.Bold = True
            c.MousePointer = 0
        End If
    Next c
    cmdBarExit.BackColor = ThBad
    munGame.Visible = False
    mnuOptions.Visible = False
    mnuAbout.Visible = False
End Sub

Private Sub cmdBarNew_Click()
    mnuStartGame_Click
End Sub

Private Sub cmdBarGame_Click()
    frmPanel.OpenPanel "game"
End Sub

Private Sub cmdBarOptions_Click()
    frmPanel.OpenPanel "options"
End Sub

Private Sub cmdBarAbout_Click()
    frmPanel.OpenPanel "about"
End Sub

Private Sub cmdBarExit_Click()
    mnuExit_Click
End Sub

Private Sub Form_Load()
    ApplyLocale
    BarButtons

' *********************************************************************************************************
' FORM LOAD
' *********************************************************************************************************
' Purpose: Get phrase from file ( phrase.txt ) and store it in memory, Set default Values
'                Go into idle state ( Call idle Funtion )
'
' Variables Used: Counter - (counter), Hints , Hintsave, Timeleft, Timesave, Playsound
'
' Funtions Called: Idle - Sets Game into Idle state
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

'*******************************************************************************
' OPEN Phrase file on disk.
'*******************************************************************************
Open "data\phrases.txt" For Input As #1

'*******************************************************************************
' Put Phrase list in memory using an array
'*******************************************************************************

Dim counter As Integer

If Player1name = "" Then Player1name = Loc("player.1")
If Player2name = "" Then Player2name = Loc("player.2")

lblPlayerName(0).Visible = True
lblPlayerName(1).Visible = True

Do Until EOF(1)

  Input #1, PhraseArray(counter)
  
  counter = counter + 1

Loop

'*******************************************************************************
' set default values and start game at idle
'*******************************************************************************
Dim I As Integer

TimerPAUSE.Enabled = False
TimerSolve2player.Enabled = False

Hints = 1 ' Default  = 1
HintSave = 1 ' Default = 1

TimeLeft = 60 ' Default TimeLeft = 60
TimeSave = 60 ' Default Timesave = 60

If bSound = True Then PlaySound "sounds\back.wav", 0, 1
Me.Caption = Loc("caption.start1")
Me.Refresh
Idle

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmrulez
Unload FrmCreator
Unload frmOptions
Unload frmEnterName
Unload frmSplash


End Sub

Private Sub lblGamePhrase_Click()


If bIdle = True Then
bIdle = False

ClearGrafix 'Calls Function to Clear grafics and restore settings

Dim x As Integer
x = Int((2 * Rnd) + 1)
If x = 1 Then
    If bSound = True Then PlaySound "sounds\begin.wav", 0, 1
End If

If x = 2 Then
    If bSound = True Then PlaySound "sounds\begin2.wav", 0, 1
End If

If mnu1player.Checked = True Then Players = 1 ' if checked players = 1
If mnu1player.Checked = True Then Set1playerGame ' if checked call funtion
If mnu2player.Checked = True Then Players = 2
If mnu2player.Checked = True Then Set2PlayerGame
End If
End Sub

Public Sub mnu1player_Click()

' *********************************************************************************************************
' mnu1player (on click)
' *********************************************************************************************************
' Purpose: If the Player 1 is checked in the menu asign a checkmark and disable
'                player2 's checkmark
'
' Variables Used: NONE
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
     
     mnu1player.Checked = True 'Put checkmark for Player 1
     mnu2player.Checked = False ' disable Checkmark for Player 2

End Sub

Public Sub mnu2player_Click()

' *********************************************************************************************************
' mnu2player (on click)
' *********************************************************************************************************
' Purpose: If the Player 2 is checked in the menu asign a checkmark and disable
'                player1 's checkmark
'
' Variables Used: NONE
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

    mnu2player.Checked = True 'Put checkmark for Player 2
    mnu1player.Checked = False ' disable Checkmark for Player 1

End Sub

Public Sub mnuAdvancedOptions_Click()

' *********************************************************************************************************
' mnuAdvancedOptions (on click)
' *********************************************************************************************************
' Purpose:  Load the Options form to get the Game options
'
' Variables Used: NONE
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
    
    Load frmOptions ' Load the frmOptions in memory
    frmOptions.Show ' Show frmOptions

End Sub

Public Sub mnuCreator_Click()

' *********************************************************************************************************
' mnuCreator (on click)
' *********************************************************************************************************
' Purpose: Load the Form about the author ( me )
'
' Variables Used: NONE
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
    
    Load FrmCreator 'Load frmCreator into Memory
    FrmCreator.Show ' Show frmCreator

End Sub

Public Sub mnuEndGame_Click()

' *********************************************************************************************************
' mnuEndGame (on click)
' *********************************************************************************************************
' Purpose: When the end Game is Clicked, The Screen is cleared and Idle state is enabled
'
' Variables Used: NONE
'
' Function(s) Called: ClearGrafix ( Funtion to turn off all Grafixs and restore settings )
'                               Idle ( Function to put game in Idle state )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

ClearGrafix ' Turn off all grafics and restore settings
Idle ' Put game in Idle state

End Sub

Public Sub mnuExit_Click()
' *********************************************************************************************************
' mnuExit (on click)
' *********************************************************************************************************
' Purpose: END GAME, When the Game is exited, All forms are unloaded from memory
'
' Variables Used: NONE
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

Unload FrmCreator
Unload frmEnterName
Unload frmrulez
Unload frmOptions


End ' *********  THE ONLY WAY OUT!!!!!!!! ( &F4 ) ********

End Sub

Public Sub mnuGetNames_Click()

' *********************************************************************************************************
' mnuGetNames ( on click )
' *********************************************************************************************************
' Purpose: Load the Form frmEnterName to get Players Names
'
' Variables Used: NONE
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

    Load frmEnterName
    frmEnterName.Show
   
End Sub

Public Sub mnuHints_Click()
' *********************************************************************************************************
' mnuHints ( on click )
' *********************************************************************************************************
' Purpose: Turns On And off the Hint(s) Buttons
'
' Variables Used: NONE
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *******************************************************************************************************

If mnuHints.Checked = True Then ' if hints are visible disable them
    cmdHint(0).Visible = False
    cmdHint(1).Visible = False
    mnuHints.Checked = False
Else
    If Players = 2 And PLAYERTWO = True Then cmdHint(1).Visible = True
    If Players = 2 And PLAYERONE = True Then cmdHint(0).Visible = True
    If Players = 1 Then cmdHint(0).Visible = True
    mnuHints.Checked = True
End If


End Sub

Public Sub mnuRulez_Click()

' *********************************************************************************************************
' mnuRulez ( on click )
' *********************************************************************************************************
' Purpose: Form rulez is Loaded
'
' Variables Used: NONE
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
    
    frmrulez.Show
    frmrulez.Refresh
    
End Sub

Public Sub mnuShowTimer_Click()

Dim TimeBox As Boolean

' *********************************************************************************************************
' mnuShowTimer ( on click )
' *********************************************************************************************************
' Purpose:  Turns On And off the Timer display on the screen
'
' Variables Used: TimeBox
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *******************************************************************************************************

If mnuShowTimer.Checked = False Then ' If unchecked in Menu make the timer visible
    mnuShowTimer.Checked = True
    TimeBox = True
    lblTimer1.ForeColor = vbBlue
    lblTimer1.Refresh
 End If

If TimeBox <> True Then  ' If unchecked in Menu make the timer visible
    mnuShowTimer.Checked = False
    lblTimer1.ForeColor = vbBlack
    lblTimer1.Refresh
 End If

End Sub

Public Sub mnuSound_Click()
If mnuSound.Checked = True Then
    mnuSound.Checked = False
    bSound = False
Exit Sub
End If

If mnuSound.Checked = False Then
    mnuSound.Checked = True
    bSound = True
    Exit Sub
End If

End Sub

Public Sub mnuStartGame_Click()

' *********************************************************************************************************
' mnuStartGame ( on click )
' *********************************************************************************************************
' Purpose: Get phrase from file ( phrase.txt ) and store it in memory, Set default Values
'                Go into idle state ( Call idle Funtion )
'
' Variables Used: Playsound -Used to playsound, Players -Used to See how many players
'
' Funtions Called: ClearGrafix - Clears all grafics and restores all settings
'                           Set1playerGame - Sets 1 Player settings
'                           Set2playerGame - Sets 2 Player settings
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

ClearGrafix 'Calls Function to Clear grafics and restore settings

Dim x As Integer
x = Int((2 * Rnd) + 1)
If x = 1 Then
    If bSound = True Then PlaySound "sounds\begin.wav", 0, 1
End If

If x = 2 Then
    If bSound = True Then PlaySound "sounds\begin2.wav", 0, 1
End If

If mnu1player.Checked = True Then Players = 1 ' if checked players = 1
If mnu1player.Checked = True Then Set1playerGame ' if checked call funtion
If mnu2player.Checked = True Then Players = 2
If mnu2player.Checked = True Then Set2PlayerGame

End Sub

Private Sub PicSkeleton_Click()
' *********************************************************************************************************
' PicSkeleton ( on click )
' *********************************************************************************************************
' Purpose: When clicked start the game ( copy of mnuStart ) so the user can click
'                On the picture to start the game.
'
' Variables Used: Playsound -Used to playsound, Players -Used to See how many players
'
' Funtions Called: ClearGrafix - Clears all grafics and restores all settings
'                           Set1playerGame - Sets 1 Player settings
'                           Set2playerGame - Sets 2 Player settings
'
' Written By: Corey McInroy
' Date Modified: May 30/99
' *********************************************************************************************************

ClearGrafix 'Calls Function to Clear grafics and restore settings

Dim x As Integer
x = Int((2 * Rnd) + 1)
If x = 1 Then
    If bSound = True Then PlaySound "sounds\begin.wav", 0, 1
End If

If x = 2 Then
    If bSound = True Then PlaySound "sounds\begin2.wav", 0, 1
End If

If mnu1player.Checked = True Then Players = 1 ' if checked players = 1
If mnu1player.Checked = True Then Set1playerGame ' if checked call funtion
If mnu2player.Checked = True Then Players = 2
If mnu2player.Checked = True Then Set2PlayerGame
 End Sub

Private Sub Picture1_Click(Index As Integer)
If bIdle = True Then
bIdle = False

ClearGrafix 'Calls Function to Clear grafics and restore settings

Dim x As Integer
x = Int((2 * Rnd) + 1)
If x = 1 Then
    If bSound = True Then PlaySound "sounds\begin.wav", 0, 1
End If

If x = 2 Then
    If bSound = True Then PlaySound "sounds\begin2.wav", 0, 1
End If

If mnu1player.Checked = True Then Players = 1 ' if checked players = 1
If mnu1player.Checked = True Then Set1playerGame ' if checked call funtion
If mnu2player.Checked = True Then Players = 2
If mnu2player.Checked = True Then Set2PlayerGame
End If
End Sub

Private Sub Picture2_Click()

' *********************************************************************************************************
' Picture2  ( on click )
' *********************************************************************************************************
' Purpose: When the top Banner is clickd playsound
'
' Variables Used: Playsound - Used to playsound
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *******************************************************************************************************

   If bSound = True Then PlaySound "sounds\back.wav", 0, 1

End Sub

Private Sub Picture3_Click()

' *********************************************************************************************************
' Picture3( on click )
' *********************************************************************************************************
' Purpose: When the logo on the top of the screen is clicked playsound
'
' Variables Used: Playsound - Used to Playsound
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *******************************************************************************************************
    
    If bSound = True Then PlaySound "sounds\intro.wav", 0, 1

End Sub

Private Sub Timer_Timer()
' *********************************************************************************************************
' Timer
' *********************************************************************************************************
' Purpose: CountDown the time, And if time is over end the Game, as well as check
'               if the phrase has been guessed. If guessed then display Winning Sequence.
'
' Variables Used: Timeleft - How much time is left, Players - The amount of Players
'
' Funtions Called: Winner - Calls Funtion to Display winning Sequence
'                           Gameover - Calls Function to Display Gameover Sequence
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

If mnuShowTimer.Checked = True And lblTimer1.Visible = False _
Then lblTimer1.Visible = True ' If timer Enabled show timeleft

If TimeLeft = 0 And Players = 1 Then Gameover ' If time is up Call Gameover ( SEE ABOVE )
If TimeLeft = 0 And Players = 2 Then Gameover ' If time is up Call Gameover ( SEE ABOVE )

If UCase(lblGamePhrase) = UCase(PhraseArray(RandomPhrase)) Then _
Winner ' If Player guesses the phrase then Display Winning Seuence


  lblTimer1.Caption = TimeLeft ' display number on the screen
  TimeLeft = TimeLeft - 1 ' increment time by 1
      
 End Sub

Private Function GetPhrase()
Dim I As Integer

' *********************************************************************************************************
' GetPhrase
' *********************************************************************************************************
' Purpose: Get a random Phrase from array in Memory,
'
'
' Variables Used: RandomPhrase - Phrase picked randomly, PhraseUnknown holds the
'                        Phrase that is unknown ( eg E_ec_tion_r ). PhrareLength - Then length
'                        of the Phrase. I - ( counter )
'
' Funtions Called: Winner - Calls Funtion to Display winning Sequence
'                           Gameover - Calls Function to Display Gameover Sequence
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

lblGamePhrase = "" ' empty the value
PhraseUnknown = "" ' empty the value
   
RandomPhrase = Int((2000 * Rnd) + 1) ' Generate random value between 1 and 2000.
PhraseLength = Len(PhraseArray(RandomPhrase)) ' get length of the phrase
 
lblWinLose.Visible = True ' make visible
lblWinLose.Caption = "" ' empty the value

  
   For I = 1 To (PhraseLength) ' do as many times as letters in the phrase
   
If UCase(Mid(PhraseArray(RandomPhrase), I, 1)) >= UCase("a") _
And UCase(Mid(PhraseArray(RandomPhrase), I, 1)) <= UCase("z") Then
'if letter in Phrase
   
   PhraseUnknown = PhraseUnknown + "-" ' If it is a letter use the - to indicate a letter
 Else
   PhraseUnknown = PhraseUnknown + " " ' If it is a space then indicate a space
End If

    Next
 
 If lblGamePhrase.Visible = False Then lblGamePhrase.Visible = True
 ' if phrase no visible make visible
 
 lblGamePhrase = PhraseUnknown ' Gamephrase on screen = ( eg E_ec_tion_r )
 
   End Function
  
Sub Gameover()
     
' *********************************************************************************************************
' Gameover
' *********************************************************************************************************
' Purpose: To Display Gameover Sequence
'
' Variables Used: NONE
'
' Funtions Called: ClearGrafix - Used to clear grafics and restore settings
'                           SoundGameOver - Plays a losing Sound
'                           Timer - Timer to Display last try grafic. ( HEADSHOT )
'                           TimerGameOver - Timer to display Gameover at end of game
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
     
    ClearGrafix
    Timer.Enabled = False ' Turn off the countdown timer
    SoundGameOver ' Play the losing sound
    lblWinLose.Visible = True
    lblWinLose.Caption = Loc("msg.gameOver")
    TimerGameOver.Enabled = True
    picGameOver.Visible = True
        
End Sub

Private Sub timerdeath_Timer()

' *********************************************************************************************************
' timerdeath ( timer )
' *********************************************************************************************************
' Purpose: To Display Gameover Sequence
'
' Variables Used: Players - How many players, - PLAYERONE - Playerone's turn
'                          PLAYERTWO - Playertwo's turn, Endtime ( counter )
'                          Bothdead - variable to see if both players are dead
'
' Funtions Called: ClearGrafix - Used to clear grafics and restore settings
'                           Gameover - Display Game over Sequence
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
     
If Players = 1 Then ' if 1 player game display headshot then gameover sequence.
   If Endtime = 107 Then picDude1.Picture = LoadPicture("grafix\trys\try6.jpg")
   PLAYERONE = False
   PLAYERTWO = False
   If Endtime = 200 Then Gameover
   If Endtime = 250 Then ClearGrafix
   If Endtime > 320 Then Endtime = 0 And timerDeath.Enabled = False
End If

If PLAYERONE = True And bothdead = False Then ' if playerone is dead and it's his turn
   If Players <> 1 Then
    If Endtime = 1 Then picClear.Visible = True
    If Endtime = 107 Then ' display the head shot and return to game
       PicPlayer2.Picture = LoadPicture("grafix\trys\try6.jpg")
       cmdHint(1).Visible = False ' hide the hint button for player 1
       cmdSolve2Visible = False ' hide the solve button for player 1
        If Endtime = 107 Then picClear.Visible = False
       timerDeath.Enabled = False ' Turn off timer
       Endtime = 0
   End If
   End If
End If

If PLAYERTWO = True And bothdead = False Then ' if playertwo's turn and it's his turn
    If Endtime = 1 Then picClear.Visible = True
   If Endtime = 107 Then ' display headshot and return to game
   PicPlayer1.Picture = LoadPicture("grafix\trys\try6.jpg")
   cmdHint(0).Visible = False
   cmdSolve1.Visible = False
    If Endtime = 107 Then picClear.Visible = False
   timerDeath.Enabled = False
    Endtime = 0
  End If
  End If


If bothdead = True Then ' If both players are dead end the game
   
    If Endtime = 1 Then picClear.Visible = True
   If Endtime = 107 Then PicPlayer2.Picture = LoadPicture("grafix\trys\try6.jpg")
   If Endtime = 107 Then PicPlayer1.Picture = LoadPicture("grafix\trys\try6.jpg")
   If Endtime = 200 Then picClear.Visible = False
   If Endtime = 250 Then Gameover
   If Endtime = 250 Then ClearGrafix
   If Endtime > 320 Then Endtime = 0 And timerDeath.Enabled = False
   PLAYERTWO = False
   PLAYERONE = False

End If
   
      Endtime = Endtime + 1 ' incease timer by one

End Sub

Private Sub TimerGameOver_Timer()
  Dim x As Integer
  Dim a As Integer
' *********************************************************************************************************
' TimerGameover
' *********************************************************************************************************
' Purpose: To display the Player Hanging, and the gameover at the end of the game,
'               clear the grafics, and return the game to idle
'
' Variables Used: HangCounter - (counter), X - ( random number ), A - ( random Number )
'
' Funtions Called: ClearGrafix - Used to clear grafics and restore settings
'                           Idle - Set the game to idle state
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
Hangcounter = Hangcounter + 1
 x = Int((6 * Rnd) + 1)
     
If x = 1 Then
     picGameOver.Picture = LoadPicture("grafix\Gameover\1.jpg")
     picGameOver.Refresh
  End If
If x = 2 Then
    picGameOver.Picture = LoadPicture("grafix\Gameover\2.jpg")
    picGameOver.Refresh
  End If
If x = 3 Then
    picGameOver.Picture = LoadPicture("grafix\Gameover\3.jpg")
    picGameOver.Refresh
  End If
If x = 4 Then
    picGameOver.Picture = LoadPicture("grafix\Gameover\4.jpg")
    picGameOver.Refresh
  End If
If x = 5 Then
    picGameOver.Picture = LoadPicture("grafix\Gameover\5.jpg")
    picGameOver.Refresh
   End If
If x = 6 Then
    picGameOver.Picture = LoadPicture("grafix\Gameover\6.jpg")
    picGameOver.Refresh
  
    a = Int((6 * Rnd) + 1)
' Display Gameover in Different colors
If a = 1 Then lblWinLose.ForeColor = vbMagenta
If a = 2 Then lblWinLose.ForeColor = vbBlue
If a = 3 Then lblWinLose.ForeColor = vbRed
If a = 4 Then lblWinLose.ForeColor = vbGreen
If a = 5 Then lblWinLose.ForeColor = vbCyan
If a = 6 Then lblWinLose.ForeColor = vbYellow
  End If

lblEndPhrase.Visible = True
lblEndPhrase.Caption = PhraseArray(RandomPhrase) ' Display the Phrase
    
    If Hangcounter > 175 Then ' When counter reaches 40 go into idle state
        ClearGrafix
        Idle
    End If
    
End Sub

Public Function SoundGameOver()
  If bSound = False Then Exit Function
  randomSound = 0
  
' *********************************************************************************************************
' SoundGameover
' *********************************************************************************************************
' Purpose: To Play a random sound When the Game is over
'
' Variables Used: Playsound - Used to play sound, randomSound - ( random # )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

   randomSound = Int((9 * Rnd) + 1)    ' Generate random value between 1 and 9.
 
If randomSound = 1 Then
    PlaySound "sounds\lose\lose1.wav", 0, 1
End If
 
If randomSound = 2 Then
    PlaySound "sounds\lose\lose2.wav", 0, 1
 End If
 
If randomSound = 3 Then
    PlaySound "sounds\lose\lose3.wav", 0, 1
 End If
  
If randomSound = 4 Then
    PlaySound "sounds\lose\lose4.wav", 0, 1
 End If
 
If randomSound = 5 Then
    PlaySound "sounds\lose\lose5.wav", 0, 1
 End If
 
If randomSound = 6 Then
    PlaySound "sounds\lose\lose6.wav", 0, 1
 End If
 
If randomSound = 7 Then
    PlaySound "sounds\lose\lose7.wav", 0, 1
 
 End If
If randomSound = 8 Then
    PlaySound "sounds\lose\lose8.wav", 0, 1
 End If

If randomSound = 9 Then
    PlaySound "sounds\lose\lose9.wav", 0, 1
 End If

End Function

Public Function SoundWinner()
 
 ' *********************************************************************************************************
' SoundWinner
' *********************************************************************************************************
' Purpose: To Play a random sound When the player Wins
'
' Variables Used: Playsound - Used to play sound, randomSound - ( random # )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
 
If bSound = False Then Exit Function
   randomSound = Int((7 * Rnd) + 1)    ' Generate random value between 1 and 7.
   
If randomSound = 1 Then
    PlaySound "sounds\winner\win1.wav", 0, 1
 End If
  
If randomSound = 2 Then
    PlaySound "sounds\winner\win2.wav", 0, 1
 End If
 
If randomSound = 3 Then
    PlaySound "sounds\winner\win3.wav", 0, 1
 End If
  
If randomSound = 4 Then
    PlaySound "sounds\winner\win4.wav", 0, 1
 End If
 
If randomSound = 5 Then
     PlaySound "sounds\winner\win5.wav", 0, 1
End If
 
If randomSound = 6 Then
    PlaySound "sounds\winner\win6.wav", 0, 1
 End If
 
 If randomSound = 7 Then
    PlaySound "sounds\winner\win7.wav", 0, 1
 End If
 randomSound = 0
End Function

Public Function Idle()

' *********************************************************************************************************
' Idle
' *********************************************************************************************************
' Purpose: To Play a random sound When the Game is over
'
' Variables Used: Playsound - Used to play sound, randomSound - ( random # )
'
' Function(s) Called: ClearGrafix - Used to clear grafics and restore settings
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

ClearGrafix
If bSound = True Then PlaySound "sounds\back.wav", 0, 1
Timer.Enabled = False ' Turn off the timer
Me.Caption = Loc("caption.start2")
Me.Refresh
TimerIdle.Tag = 2
If iTimesWon = 10 Then
PicWinner.Left = 6960
PicWinner.Top = 4080
iTimesWon = 0
End If

TimerIdle.Enabled = True ' Idle timer for Idle Grafics
bIdle = True
PicSkeleton.Left = 4800
PicSkeleton.Top = 2400
skel = 0
iChange = 2
PicSkeleton.Tag = 1
mnu1player.Enabled = True ' Enable players to be Enabled in menu
mnu2player.Enabled = True
mnuAdvancedOptions.Enabled = True ' Enable Options in Menu
mnuStartGame.Enabled = True ' Enable Start new game in Menu
mnuGetNames.Enabled = True ' Enable the Set names in menu

PicSkeleton.Visible = True

End Function

Public Sub TimerIdle_Timer()
  Dim x As Integer
  Dim iRandSound As Integer
  
' *********************************************************************************************************
' TimerIdle ( timer )
' *********************************************************************************************************
' Purpose: To show a random Grafic
'
' Variables Used: X - ( random # )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
     
x = Int((10 * Rnd) + 1) ' Generate # from 1 to 6
If bRandsound = True Then
   iRandSound = Int((5 * Rnd) + 1)
   bRandsound = False
End If


   skel = skel + 1
 Select Case Val(TimerIdle.Tag)
 
 Case (1)
   
   If skel >= 4 Or skel <= 0 Then skel = 1

If PicSkeleton.Top >= 8000 Then
   PicSkeleton.Visible = False
   PicSkeleton.Picture = LoadPicture("grafix\idle\axe\1.gif")
   PicSkeleton.Refresh
End If

If PicSkeleton.Top >= 8200 Then
   TimerIdle.Tag = 2
   iChange = 2
   bRandsound = True
   PicSkeleton.Top = 2400
   PicSkeleton.Left = 4800

End If
If iRandSound = 5 Or iRandSound = 4 Then
   iRandSound = Int((3 * Rnd) + 1)
   If iRandSound = 1 And bSound = True Then PlaySound "sounds\idle\elaugh1.wav", 0, 1
   If iRandSound = 2 And bSound = True Then PlaySound "sounds\idle\elaugh2.wav", 0, 1
   If iRandSound = 3 And bSound = True Then PlaySound "sounds\idle\elaugh3.wav", 0, 1
End If

PicSkeleton.Top = PicSkeleton.Top + 25 + 40
PicSkeleton.Left = PicSkeleton.Left + (x * 2) - 2

If PicSkeleton.Top < 8300 Then
    PicSkeleton.Picture = LoadPicture("grafix\idle\boss\" & skel & ".gif")
    PicSkeleton.Refresh
End If
  
Case (2)

PicSkeleton.Tag = Val(PicSkeleton.Tag) + 1
If skel >= 8 Then skel = 1

If iChange = 2 Then
    
     PicSkeleton.Visible = True
     PicSkeleton.Picture = LoadPicture("grafix\idle\axe\" & skel & ".gif")
     PicSkeleton.Refresh
       
End If
If iRandSound = 5 Then
   iRandSound = Int((4 * Rnd) + 1)
   If iRandSound = 1 And bSound = True Then PlaySound "sounds\idle\neverwin.wav", 0, 1
   If iRandSound = 2 And bSound = True Then PlaySound "sounds\idle\try.wav", 0, 1
End If
If Val(PicSkeleton.Tag) = 77 Then
    iChange = Int((3 * Rnd) + 1)
    If iChange <> 2 Then
        
        If iChange = 1 Then
            PicSkeleton.Visible = False
            PicSkeleton.Picture = LoadPicture("grafix\idle\boss\1.gif")
            PicSkeleton.Refresh
        End If
        If iChange = 3 Then
            
            PicSkeleton.Picture = LoadPicture("grafix\idle\walk\1.gif")
            PicSkeleton.Visible = False
            PicSkeleton.Refresh
        End If
    End If
End If

If Val(PicSkeleton.Tag) >= 80 Then

    If iChange <> 2 Then
        PicSkeleton.Top = 2000 + Int(400 * Rnd) + 1
        PicSkeleton.Left = 4700 - Int(200 * Rnd) + 1
    PicSkeleton.Visible = True
    PicSkeleton.Refresh
    End If
    
    TimerIdle.Tag = iChange
    iChange = 2
    bRandsound = True
    PicSkeleton.Refresh
    PicSkeleton.Tag = 1
End If


Case (3)

If skel >= 21 Then skel = 1

If PicSkeleton.Top >= 7900 Then
   PicSkeleton.Visible = False
   PicSkeleton.Picture = LoadPicture("grafix\idle\axe\1.gif")
   PicSkeleton.Refresh
End If

If PicSkeleton.Top >= 8000 Then

 TimerIdle.Tag = 2
 iChange = 2
   bRandsound = True
   PicSkeleton.Top = 2400
   PicSkeleton.Left = 4800

End If

PicSkeleton.Top = PicSkeleton.Top + 25 + 65
PicSkeleton.Left = PicSkeleton.Left + (x * 2) - 9

If iRandSound = 4 Then
     If bSound = True Then PlaySound "sounds\idle\thunder.wav", 0, 1
        picLightning.Visible = True
        TmrLightning.Enabled = True
End If


If iRandSound = 5 Then
   iRandSound = Int((4 * Rnd) + 1)
   If iRandSound = 1 Or iRandSound = 2 Then
     If bSound = True Then PlaySound "sounds\idle\thunder.wav", 0, 1
        picLightning.Visible = True
        TmrLightning.Enabled = True
   End If
    
   If iRandSound = 3 And bSound = True Then PlaySound "sounds\idle\wind.wav", 0, 1
   If iRandSound = 4 And bSound = True Then PlaySound "sounds\idle\scream.wav", 0, 1

End If
If PicSkeleton.Top < 8000 Then
   PicSkeleton.Picture = LoadPicture("grafix\idle\walk\" & skel & ".gif")
   PicSkeleton.Refresh
End If


End Select


End Sub

Public Function ClearGrafix()
Dim a As Integer
' *********************************************************************************************************
' ClearGrafix
' *********************************************************************************************************
' Purpose:To Clear all the grafix from the screen and restore all initial settings
'
' Variables used:Endtime,EndtimeWinner,Turn1player, A - ( counter )
'
'
' Funtions Called: ClearGrafix - Clears all grafics and restores all settings
'                           Set1playerGame - Sets 1 Player settings
'                           Set2playerGame - Sets 2 Player settings
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

' *********************************************************************************************************
'  *NOTE : All game settings are reset to either '0' or false ( OFF )
'                Exept for any menu options ( Advanced options turned on here )
' *********************************************************************************************************

' variables are set to zero , - 1 player turns = 0 and counter set to 0
bIdle = False
Endtime = 0
EndtimeWinner = 0
Hangcounter = 0
Turn1player = 0

' All timers are turned off

TimerWinner.Enabled = False
timerDeath.Enabled = False
Timer.Enabled = False
TimerGameOver.Enabled = False
TimerIdle.Enabled = False

' all pictures in form ( used in game ) are hidden
Me.Caption = Loc("app.titleShort")
PicWinner.Visible = False
picGameOver.Visible = False
PicSkeleton.Visible = False
picDude1.Visible = False
PicPlayer1.Visible = False
PicPlayer2.Visible = False
PicWinnertext(1).Visible = False
PicWinnertext(0).Visible = False

' Labels are set to default

LblWinnerbox(0) = ""
LblWinnerbox(1) = ""
lblPlayerName(0) = Player1name
lblPlayerName(1) = Player2name
lblTimer1.Visible = False
lblWinLose.Visible = False
lblGamePhrase.Visible = False
lblEndPhrase.Visible = False

' Restore the Advanced options in the menu

mnuAdvancedOptions.Enabled = True
 
 ' Hide all the letters on the screen
 
 For a = 0 To 25
cmdLetter(a).Visible = False
Next

' Hide the Solve buttons and the Hint buttons from the screen

cmdSolve1.Visible = False
cmdSolve2.Visible = False
cmdHint(0).Visible = False
cmdHint(1).Visible = False

' hide the Solve text box

txtSolve1 = ""
txtSolve1.Visible = False

End Function

Public Function CheckLetter()

Dim R As Boolean
Dim Miss As Integer
Dim There As Boolean
There = False

' *********************************************************************************************************
' CheckLetter
' *********************************************************************************************************
' Purpose:To To Check the Letter the user Clicked
'
' Variables used: I - (counter), R - used to see if letter is there, There  used to
'                         see if letter is there,  PhraseLength - the Length of the
'                         Phrase, Miss - ( random # ), PhraseUnknown ( result ( eg E_e_ )),
'                         PlaySound - Play a sound
'
' Funtions Called: Player1Game, TwoPlayerGame
'
'Variables From/To (Function): Letter - FROM cmdletter, Wrongletter TO
'                                               PlayerOneGame, and TwoplayerGame
'
'Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

PhraseLength = Len(PhraseArray(RandomPhrase))

    For I = 1 To (PhraseLength)
   
If UCase(Mid(PhraseArray(RandomPhrase), I, 1)) = UCase(Letter) _
And UCase(Mid(PhraseArray(RandomPhrase), I, 1)) = UCase(Letter) Then
' If the letter in the Phrase
R = True
There = True
End If

' if r = True then put the letter in the phrase
If R = True Then Mid(PhraseUnknown, I, 1) = Letter
R = False
Next
R = False

' if letter not there then Wrongletter = false
If There = False Then WrongLetter = True

lblGamePhrase = PhraseUnknown

' If the Phrase is now correct then the player wins call winner for winng sequence
If UCase(lblGamePhrase) = UCase(PhraseArray(RandomPhrase)) Then
Winner
End If

' If the letter is in the phrase then play this sound
If WrongLetter = False Then
  Miss = Int((2 * Rnd) + 1)
  If Miss = 1 And bSound = True Then PlaySound "sounds\Trys\miss1.wav", 0, 1
  If Miss = 2 And bSound = True Then PlaySound "sounds\Trys\miss2.wav", 0, 1
End If

' if the letter is Wrong goto the function and carry on with the game
If Players = 1 And WrongLetter = True Then Player1Game
If Players = 2 And WrongLetter = True Then TwoPlayerGame


End Function

Public Function GetHint()

Dim Hintdone, NEWSTRING, RE, THERE2 As Boolean
Dim F, B, LoseButton As Integer

' *********************************************************************************************************
' GetHint
' *********************************************************************************************************
' Purpose:To Insert a letter in the phrase.
'
' Variables used: F - (counter), B- (counter), Losebutton - (counter)
'                          PhraseLength - the Length of the Phrase,
'                          PhraseUnknown ( result ( eg E_e_ )
'
'
' Funtions Called: Player1Game, TwoPlayerGame
'
'Variables From/To (Function): Letter - FROM cmdletter, Wrongletter TO
'                                               PlayerOneGame, and TwoplayerGame
'
'Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************


PhraseLength = Len(PhraseArray(RandomPhrase))


    For F = 1 To (PhraseLength)
   
' Find first available letter
 If UCase(Mid(PhraseUnknown, F, 1)) = ("-") Then _
  NEWSTRING = Mid(PhraseArray(RandomPhrase), F, 1)
 
 If NEWSTRING <> "" Then F = PhraseLength
 If NEWSTRING <> "" Then Hintdone = True
  
  Next

PhraseLength = Len(PhraseArray(RandomPhrase))

For B = 1 To (PhraseLength)
 ' Insert the letter from above ( first letter found )
If UCase(Mid(PhraseArray(RandomPhrase), B, 1)) = UCase(NEWSTRING) _
And UCase(Mid(PhraseArray(RandomPhrase), B, 1)) = UCase(NEWSTRING) _
Then RE = True ' if a letter went into the phrase


 ' If the phrase has changed it = Newstring
If RE = True Then Mid(PhraseUnknown, B, 1) = NEWSTRING
    
 ' Hide the letter that was chosen
   
   For LoseButton = 0 To 25
      If UCase(cmdLetter(LoseButton).Caption) = UCase(NEWSTRING) Then _
        cmdLetter(LoseButton).Visible = False
    Next

If RE = True Then RE = False ' set Re to false

Next

lblGamePhrase = PhraseUnknown

' If the phrase is correct the winning sequence is called
If UCase(lblGamePhrase) = UCase(PhraseArray(RandomPhrase)) Then
Winner
End If

Hintdone = False ' reset hintdone to false

End Function

Private Sub TimerPAUSE_Timer()
' *********************************************************************************************************
' TimerPAUSE ( timer )
' *********************************************************************************************************
' Purpose: To Pause a second before playing end sound ( BUG-FIX )
'
' Variables Used: Playsound - Used to play sound, PauseTime (counter)
'
' Functions Called: Sound Winner ( Function to play random Sound )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
PauseTime = PauseTime + 1

If PauseTime = 0 Then PauseTime = 1
If PauseTime = 4 Then SoundWinner ' pause then play the sound
If PauseTime > 4 Then
   PauseTime = 0
    TimerPAUSE.Enabled = False
 End If

End Sub



Private Sub TimerWinner_Timer()
 Dim v As Integer
 
' *********************************************************************************************************
' TimerWinner ( timer )
' *********************************************************************************************************
' Purpose: Move Grafic across the screen when player wins
'
' Variables Used: V - (random #) , EndTimeWinner - (counter), Hints - Amount of
'                          hints left, Hintsave - Saved settings for amount of hints
'                          DeltaX - left to right movements , DeltaY - Horizontal movements
'
' Functions Called:ClearGrafix ( Funtion to turn off all Grafixs and restore settings )
'                            Idle ( Function to put game in Idle state )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
 
   PicWinnertext(1).Visible = True
   PicWinnertext(0).Visible = True
   
   EndtimeWinner = EndtimeWinner + 1
   
   v = Int((4 * Rnd) + 1)
     
     If v = 1 Then
     PicWinner.Picture = LoadPicture("grafix\winner\1.jpg")
     PicWinner.Refresh
    End If
    If v = 2 Then
     PicWinner.Picture = LoadPicture("grafix\winner\2.jpg")
     PicWinner.Refresh
    End If
    If v = 3 Then
     PicWinner.Picture = LoadPicture("grafix\winner\3.jpg")
     PicWinner.Refresh
    End If
    If v = 4 Then
     PicWinner.Picture = LoadPicture("grafix\winner\4.jpg")
     PicWinner.Refresh
   End If
    ' move picture left
    PicWinner.Move PicWinner.Left + DeltaX, PicWinner.Top + DeltaY
    If PicWinner.Left > ScaleLeft Then DeltaX = -5
    If PicWinner.Left + PicWinner.Width > ScaleWidth + ScaleLeft Then
        DeltaX = 1
    End If
   ' move picture down
    
    If PicWinner.Top > ScaleTop Then DeltaY = 10
    If PicWinner.Top + PicWinner.Height > ScaleHeight + ScaleTop Then
        DeltaY = -12
    End If

If EndtimeWinner = 80 Then   'if time is up then
TimerWinner.Enabled = False 'Turn off timer
EndtimeWinner = 0 'clear value for next time
Hints = HintSave ' reset the amount of hints
ClearGrafix ' clear the grafix
Idle ' enter game into idle
End If

End Sub

Public Function Winner()
' *********************************************************************************************************
' Winner
' *********************************************************************************************************
' Purpose: Seqence for displaying final sequence for winner
'
' Variables Used:DeltaX - left to right movements , DeltaY - Horizontal movements
'
' Functions Called:ClearGrafix ( Funtion to turn off all Grafixs and restore settings )
'
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
    ClearGrafix ' clear all grafics
    PicWinner.Visible = True ' Goblin is shown on screen
    lblGamePhrase = PhraseArray(RandomPhrase)
    lblGamePhrase.Visible = True ' The correct answer is shown on screen
    TimerPAUSE.Enabled = True ' start timer to pause for a moment
    TimerWinner.Enabled = True ' start timer for grafics
    TimerWinner.Interval = 49 ' Set Interval.
    DeltaX = 200    ' Initialize variables.
    DeltaY = 20
    iTimesWon = iTimesWon + 1
End Function

Private Sub TmrLightning_Timer()
TmrLightning.Tag = Val(TmrLightning.Tag) + 1

If Val(TmrLightning.Tag) >= 6 Then
TmrLightning.Tag = 1
picLightning.Visible = False
TmrLightning.Enabled = False
End If

picLightning.Picture = LoadPicture("grafix\idle\lightning\" & TmrLightning.Tag & ".gif")
picLightning.Refresh


End Sub

Private Sub txtSolve1_KeyPress(KeyAscii As Integer)

' *********************************************************************************************************
' txtSolve1 ( checks for keypress )
' *********************************************************************************************************
' Purpose: To check for a correct phrase, and if so the player wins
'
' Variables Used:DeltaX - left to right movements , DeltaY - Horizontal movements
'
' Functions Called:Winner - Displays the final winning sequence
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
 
' IF the phrase typed is the same as the the phrase, if not nothing happens
 If UCase(txtSolve1.Text) = UCase(PhraseArray(RandomPhrase)) Then
     txtSolve1.Visible = False
     txtSolve1 = ""
     Winner
    
   
   'display winning message if Player 1 wins
     
     If PLAYERONE = True Then
     Winner
     LblWinnerbox(1).Visible = True
     LblWinnerbox(1) = LocF("msg.wins", Player1name)
     txtSolve1 = ""
     End If
  
    If PLAYERTWO = True Then
     Winner
     LblWinnerbox(1).Visible = True
     LblWinnerbox(1) = LocF("msg.wins", Player2name)
     txtSolve1 = ""
     End If
End If


End Sub

Public Function Set1playerGame()
Dim a As Integer

'*********************************************************************************************************
' Set1playerGame ( Called from cmdStart_click )
' *********************************************************************************************************
' Purpose: To set the setting for a one player game
'
' Variables Used: A - (counter), Hints - Amount of hints left
'                         Hintsave - Original amount of hints.
'                         Timeleft - ( Amount of time for timer )
'                         TimeSave - Timeleft original value
'                         TimeOff- Determines if time is disabed/enabled
'
' Functions Called:GetPhrase - Gets a random Phrase from memory
'
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

' ***********************************************************************************
' Intialize the amount of hints
' ***********************************************************************************
' If show hints is checked in the menu
If mnuHints.Checked = True Then cmdHint(0).Visible = True
If HintSave = 0 Then HintSave = 1 ' set Hintsave to default (1)
If Hints = 0 Then Hints = HintSave ' if hints at zero make them default
   HintSave = Hints

' ***********************************************************************************
' Disable Menu Settings
' ***********************************************************************************

   mnu1player.Enabled = False
   mnu2player.Enabled = False
   mnuAdvancedOptions.Enabled = False

' ***********************************************************************************
' Initialize Grafics
' ***********************************************************************************
 
' show all letters on screen
 For a = 0 To 25
   cmdLetter(a).Visible = True
 Next

   ' Show player1 ' name
   lblPlayerName(1).Visible = False
   cmdSolve1.Visible = True ' Turn Solve button on
   ' place first pic of player on screen
   picDude1.Picture = LoadPicture("grafix\trys\try0.jpg")
   picDude1.Visible = True
   TimeLeft = 0 ' initialize Timeleft to zero

' ***********************************************************************************
' Set Timer Setting
' ***********************************************************************************

' Get how much time should be left ( Timesave is the old value )
If TimeSave = 0 Then TimeLeft = 60
If TimeLeft <> TimeSave Or TimeLeft = 0 Then TimeLeft = TimeSave
   Timer.Enabled = True
If TimeOff = True Then Timer.Enabled = False

   GetPhrase ' Get the phrase and start the game

End Function

Public Function Set2PlayerGame()
'*********************************************************************************************************
' Set2playerGame ( Called from cmdStart_click )
' *********************************************************************************************************
' Purpose: To set the setting for a Two player game
'
' Variables Used: A - (counter), Hints - Amount of hints left
'                         Hintsave - Original amount of hints.
'                         Hints2 - Amount of Hints for player 2
'                         try1, # of trys for playerone. try2 # of trys for player 2
'                         G-(counter), PLAYERONE - Playerone's turn
'                         PLAYERTWO - Playertwo's turn, BOTHDEAD = Both players Dead
'                         Playeronedead - IF player one is dead
'                         Playertwodead - If player 2 is Dead
'                         Player1name - Playerone's name
'                         Player2name - Playertwo's name
'
' Functions Called:GetPhrase - Gets a random Phrase from memory
'
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

Dim a As Integer

 ' ***********************************************************************************
' Intialize the amount of hints
' ***********************************************************************************
' if hints = 0 then set to the original setting ( timesave )
 If HintSave = 0 Then HintSave = 1
 If Hints = 0 Then Hints = HintSave
    HintSave = Hints
    Hints2 = Hints

' ***********************************************************************************
' Determine Menu Settings
' ***********************************************************************************

' Turn off menu settings that need to be turned off
If mnuHints.Checked = True Then cmdHint(0).Visible = True
   mnu1player.Enabled = False
   mnu2player.Enabled = False
   mnuAdvancedOptions.Enabled = False
' Show hints if menu ' show hints ' is enabled
If mnuHints.Checked = True Then
cmdHint(0).Visible = True
cmdHint(1).Visible = True
 End If


' ***********************************************************************************
' Initialize Grafics
' ***********************************************************************************

' Place letters on the screen
 For a = 0 To 25
    cmdLetter(a).Visible = True
  Next
' Show the 2 players beginning picture
    
    PicPlayer1.Picture = LoadPicture("grafix\trys\try0.jpg")
    PicPlayer1.Visible = True
    PicPlayer2.Picture = LoadPicture("grafix\trys\try0.jpg")
    PicPlayer2.Visible = True
    
' ***********************************************************************************
' Set Timer Setting
' ***********************************************************************************
' Set the timer settings ( timesave = original values )
   
   TimeLeft = 0
 If TimeSave = 0 Then TimeLeft = 60
 If TimeLeft <> TimeSave Or TimeLeft = 0 Then TimeLeft = TimeSave
    Timer.Enabled = True
 If TimeOff = True Then Timer.Enabled = False

' ***********************************************************************************
' Set up WHO GOES FIRST And Trys
' ***********************************************************************************
    Try1 = 1   ' Amount of tries for player one
    Try2 = 1   ' Amount of tries for player 2
    
   ' set it up so its nobodys turn
    PLAYERONE = False
    PLAYERTWO = False
    
    ' set player status to off
    PlayerOneDead = False
    PlayerTwoDead = False
    bothdead = False
    
    ' Display label that displays who's turn it is.
    
    LblWinnerbox(0).Visible = True
    
    ' pick a random number to see who goes first
    Dim G As Integer
      G = Int((2 * Rnd) + 1)
        
       ' if player one goes first
        If G = 1 Then
             PLAYERONE = True ' it's player one's turn
             cmdHint(1).Visible = False ' turn off 2 players Hint
             
             ' make hints visible if enabled
             If mnuHints.Checked = True Then cmdHint(0).Visible = True
             
             ' Turn 1 players Solve button on 2 player button off
             cmdSolve2.Visible = False
             cmdSolve1.Visible = True
             
             ' Display player one goes first
             LblWinnerbox(0) = LocF("msg.goesFirst", Player1name)
             
             ' Display name for player one and hide player 2's name
             lblPlayerName(1).Visible = False
             lblPlayerName(0).Visible = True
        
        End If
      ' Player two works the same way as player one but settings are opposite
      
      ' if player two goes first
      If G = 2 Then
             PLAYERTWO = True
             cmdHint(0).Visible = False
             If mnuHints.Checked = True Then cmdHint(1).Visible = True
             cmdSolve1.Visible = False
             cmdSolve2.Visible = True
             LblWinnerbox(0) = LocF("msg.goesFirst", Player2name)
             lblPlayerName(0).Visible = False
             lblPlayerName(1).Visible = True
     End If

GetPhrase ' GETPHRASE and begin game

End Function

Public Function TwoPlayerGame()
    
    Dim O As Integer
    

'*********************************************************************************************************
' Set2playerGame ( Called from cmdStart_click )
' *********************************************************************************************************
' Purpose: To Play a two player Game, switch turns and find out if both players
'               are dead, if so Gameover.
'
' Variables Used: Hints - Amount of hints left for Player 1
'                         Hints2 - Amount of Hints for player 2
'                         try1 - # of trys for playerone. try2 - # of trys for player 2
'                         O-(counter), PLAYERONE - Playerone's turn
'                         PLAYERTWO - Playertwo's turn, BOTHDEAD = Both players Dead
'                         Playeronedead - IF player one is dead
'                         Playertwodead - If player 2 is Dead
'
' Variables returned: Wrongletter = false - Checkletter
'
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
       O = Int((6 * Rnd) + 1) ' get random number betwwen 1 and 5
    LblWinnerbox(0) = ""
     
       ' Play random death sound
       
       If O = 1 And bSound = True Then PlaySound "sounds\Trys\TRY1.wav", 0, 1
       If O = 2 And bSound = True Then PlaySound "sounds\Trys\TRY2.wav", 0, 1
       If O = 3 And bSound = True Then PlaySound "sounds\Trys\TRY3.wav", 0, 1
       If O = 4 And bSound = True Then PlaySound "sounds\Trys\TRY4.wav", 0, 1
       If O = 5 And bSound = True Then PlaySound "sounds\Trys\TRY5.wav", 0, 1
       If O = 6 And bSound = True Then PlaySound "sounds\Trys\TRY7.wav", 0, 1
       
       ' display on screen who's turn it is
        If PLAYERONE Then LblWinnerbox(0) = LocF("msg.turn1", Player2name)
        If PLAYERTWO Then LblWinnerbox(0) = LocF("msg.turn2", Player1name)

' If it is playerone's Turn
If PLAYERONE = True Then
         If Try1 = 1 Then
             PicPlayer1.Picture = LoadPicture("grafix\trys\Try1.jpg")
            End If
          If Try1 = 2 Then
            PicPlayer1.Picture = LoadPicture("grafix\trys\Try2.jpg")
            End If
          If Try1 = 3 Then
            PicPlayer1.Picture = LoadPicture("grafix\trys\Try3.jpg")
            End If
          If Try1 = 4 Then
             PicPlayer1.Picture = LoadPicture("grafix\trys\Try4.jpg")
           End If
          If Try1 = 5 Then
             PicPlayer1.Picture = LoadPicture("grafix\trys\Try5.jpg")
           End If
    ' player dies, If the other player is dead the game is over
           If Try1 = 6 Then
                  If bSound = True Then PlaySound "sounds\Trys\TRY6.wav", 0, 1
                  cmdHint(0).Visible = False
                  cmdSolve1.Visible = False
                  PicPlayer1.Visible = True
                  PlayerOneDead = True ' player one is dead
                  bothdead = False ' both not dead
               If PlayerTwoDead = True Then bothdead = True ' if bothdead then gameover
                  Endtime = 0
                  timerDeath.Enabled = True
             End If
             ' these settings Happen every turn for player one
                 cmdHint(0).Visible = False
             If mnuHints.Checked = True And Hints2 > 0 Then cmdHint(1).Visible = True
                 cmdSolve1.Visible = False
                 cmdSolve2.Visible = True
                 lblPlayerName(0).Visible = False
                 lblPlayerName(1).Visible = True
                 Try1 = Try1 + 1
                 If Try1 = 7 Then Try1 = 6 ' if try1 = 7 then the player is dead
                                                      ' so try1 = 6 to end game
             End If
' player two same as player one but for player two
       If PLAYERTWO = True Then
                  If Try2 = 1 Then
                     PicPlayer2.Picture = LoadPicture("grafix\trys\Try1.jpg")
                   End If
                  If Try2 = 2 Then
                     PicPlayer2.Picture = LoadPicture("grafix\trys\Try2.jpg")
                   End If
                  If Try2 = 3 Then
                      PicPlayer2.Picture = LoadPicture("grafix\trys\Try3.jpg")
                   End If
                  If Try2 = 4 Then
                     PicPlayer2.Picture = LoadPicture("grafix\trys\Try4.jpg")
                     End If
                  If Try2 = 5 Then
                     PicPlayer2.Picture = LoadPicture("grafix\trys\Try5.jpg")
                   End If
                  If Try2 = 6 Then
                     If bSound = True Then PlaySound "sounds\Trys\TRY6.wav", 0, 1
                     cmdSolve2.Visible = False
                     cmdHint(0).Visible = False
                     cmdSolve1.Visible = False
                     bothdead = False
                     If PlayerOneDead = True Then bothdead = True
                     PicPlayer2.Visible = True
                     PlayerTwoDead = True
                     timerDeath.Enabled = True
                End If
                     cmdHint(1).Visible = False
                     If mnuHints.Checked = True And Hints > 0 Then cmdHint(0).Visible = True
                        cmdSolve2.Visible = False
                        cmdSolve1.Visible = True
                        lblPlayerName(0).Visible = True
                        lblPlayerName(1).Visible = False
                        Try2 = Try2 + 1
                        If Try2 = 7 Then Try2 = 6
                End If
  
 ' If playerone just finished his turn it's now player twos turn and vice-vesa
 
 If PLAYERONE = True Then
     histurn = True
     PLAYERONE = False
  End If
  If PLAYERTWO = True Then
     PLAYERTWO = False
     PLAYERONE = True
  End If
    
If histurn = True Then PLAYERTWO = True

WrongLetter = False ' reset Wrongletter to false

End Function
Public Function Player1Game()
  Dim M As Integer
 ' *********************************************************************************************************
' PlayeroneGame
' *********************************************************************************************************
' Purpose: To Play a random sound When the player Wins
'
' Variables Used: Playsound - Used to play sound, randomSound - ( random # )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
'***********************************************
' death scene for i player
' **********************************************
 Turn1player = Turn1player + 1
 
       
      If Turn1player < 6 Then
         
         M = Int((6 * Rnd) + 1) ' get random number betwwen 1 and 5
         ' play random sound
         If M = 1 And bSound = True Then PlaySound "sounds\Trys\TRY1.wav", 0, 1
         If M = 2 And bSound = True Then PlaySound "sounds\Trys\TRY2.wav", 0, 1
         If M = 3 And bSound = True Then PlaySound "sounds\Trys\TRY3.wav", 0, 1
         If M = 4 And bSound = True Then PlaySound "sounds\Trys\TRY4.wav", 0, 1
         If M = 5 And bSound = True Then PlaySound "sounds\Trys\TRY5.wav", 0, 1
         If M = 6 And bSound = True Then PlaySound "sounds\Trys\TRY7.wav", 0, 1
       End If
        
   ' if first turn play the first picture ...
    If Turn1player = 1 And WrongLetter = True Then
      picDude1.Picture = LoadPicture("grafix\trys\Try1.jpg")
     End If
 
    If Turn1player = 2 And WrongLetter = True Then
       picDude1.Picture = LoadPicture("grafix\trys\try2.jpg")
     End If
 
    If Turn1player = 3 And WrongLetter = True Then
       picDude1.Picture = LoadPicture("grafix\trys\try3.jpg")
     End If
  
    If Turn1player = 4 And WrongLetter = True Then
       picDude1.Picture = LoadPicture("grafix\trys\try4.jpg")
     End If
 
    If Turn1player = 5 And WrongLetter = True Then
        picDude1.Picture = LoadPicture("grafix\trys\try5.jpg")
     End If
   
    ' IF it's the Last turn And they picked the wrong letter then game over
    
    If Turn1player = 6 And WrongLetter = True Then
       If bSound = True Then PlaySound "sounds\Trys\TRY6.wav", 0, 1
       Endtime = 0
       Turn1player = 0 ' reset the trys for player one
       ClearGrafix ' clear grafix and restore settings for next game
       picDude1.Visible = True
       timerDeath.Enabled = True
     End If

WrongLetter = False ' reset WrongLetter to False

End Function
