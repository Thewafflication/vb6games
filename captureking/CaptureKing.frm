VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   570
   ClientWidth     =   13770
   Icon            =   "CaptureKing.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   13770
   StartUpPosition =   2  'CenterScreen
   Tag             =   "lline"
   Begin VB.PictureBox picInvalidMoveOver 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   10890
      Picture         =   "CaptureKing.frx":08CA
      ScaleHeight     =   705
      ScaleWidth      =   1800
      TabIndex        =   81
      Top             =   4965
      Width           =   1800
      Visible         =   0   'False
   End
   Begin VB.PictureBox picInvalidMoveOut 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   10530
      Picture         =   "CaptureKing.frx":4B24
      ScaleHeight     =   705
      ScaleWidth      =   1800
      TabIndex        =   80
      Top             =   4800
      Width           =   1800
      Visible         =   0   'False
   End
   Begin VB.PictureBox picInvalidMove 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4785
      Picture         =   "CaptureKing.frx":8D7E
      ScaleHeight     =   705
      ScaleWidth      =   1800
      TabIndex        =   79
      Top             =   5235
      Width           =   1800
   End
   Begin VB.PictureBox picCheckMateOver 
      Height          =   900
      Left            =   10095
      Picture         =   "CaptureKing.frx":CFD8
      ScaleHeight     =   840
      ScaleWidth      =   1785
      TabIndex        =   78
      Top             =   3255
      Visible         =   0   'False
      Width           =   1845
      Visible         =   0   'False
   End
   Begin VB.PictureBox picCheckMateOut 
      Height          =   660
      Left            =   10740
      Picture         =   "CaptureKing.frx":11232
      ScaleHeight     =   600
      ScaleWidth      =   1380
      TabIndex        =   77
      Top             =   3465
      Visible         =   0   'False
      Width           =   1440
      Visible         =   0   'False
   End
   Begin VB.PictureBox picCheckMate 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2955
      Picture         =   "CaptureKing.frx":1548C
      ScaleHeight     =   705
      ScaleWidth      =   1800
      TabIndex        =   76
      Top             =   5235
      Width           =   1800
   End
   Begin VB.PictureBox picCmdCheckOut 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   11865
      Picture         =   "CaptureKing.frx":196E6
      ScaleHeight     =   705
      ScaleWidth      =   1800
      TabIndex        =   75
      Top             =   6000
      Visible         =   0   'False
      Width           =   1800
      Visible         =   0   'False
   End
   Begin VB.PictureBox picCmdCheckOVER 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   1125
      Picture         =   "CaptureKing.frx":1D940
      ScaleHeight     =   705
      ScaleWidth      =   1800
      TabIndex        =   74
      Top             =   5235
      Width           =   1800
   End
   Begin VB.PictureBox picCheckOUT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   11445
      Picture         =   "CaptureKing.frx":21B9A
      ScaleHeight     =   705
      ScaleWidth      =   1800
      TabIndex        =   73
      Top             =   5715
      Visible         =   0   'False
      Width           =   1800
      Visible         =   0   'False
   End
   Begin VB.PictureBox picSendMess 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   30
      Picture         =   "CaptureKing.frx":25DF4
      ScaleHeight     =   705
      ScaleWidth      =   1005
      TabIndex        =   72
      Top             =   5235
      Width           =   1005
   End
   Begin VB.PictureBox picCMDOVER 
      Height          =   495
      Left            =   12210
      Picture         =   "CaptureKing.frx":283AA
      ScaleHeight     =   435
      ScaleWidth      =   660
      TabIndex        =   71
      Top             =   3315
      Visible         =   0   'False
      Width           =   720
      Visible         =   0   'False
   End
   Begin VB.PictureBox picCMDOUT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   12315
      Picture         =   "CaptureKing.frx":2A960
      ScaleHeight     =   705
      ScaleWidth      =   1005
      TabIndex        =   70
      Top             =   3465
      Visible         =   0   'False
      Width           =   1005
      Visible         =   0   'False
   End
   Begin VB.TextBox txtMessageOut 
      BorderStyle     =   0  'None
      Height          =   1470
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      Top             =   3555
      Width           =   1830
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   1875
      Picture         =   "CaptureKing.frx":2CF16
      ScaleHeight     =   1380
      ScaleWidth      =   120
      TabIndex        =   69
      Top             =   3600
      Width           =   120
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1860
      Picture         =   "CaptureKing.frx":2EB78
      ScaleHeight     =   180
      ScaleWidth      =   135
      TabIndex        =   68
      Top             =   3450
      Width           =   135
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   135
      Picture         =   "CaptureKing.frx":2ED0A
      ScaleHeight     =   150
      ScaleWidth      =   1725
      TabIndex        =   67
      Top             =   3450
      Width           =   1725
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   15
      Picture         =   "CaptureKing.frx":2FE2C
      ScaleHeight     =   150
      ScaleWidth      =   135
      TabIndex        =   66
      Top             =   3450
      Width           =   135
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   15
      Picture         =   "CaptureKing.frx":2FF86
      ScaleHeight     =   1350
      ScaleWidth      =   120
      TabIndex        =   65
      Top             =   3585
      Width           =   120
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   120
      Picture         =   "CaptureKing.frx":31BE8
      ScaleHeight     =   150
      ScaleWidth      =   1770
      TabIndex        =   64
      Top             =   4980
      Width           =   1770
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   15
      Picture         =   "CaptureKing.frx":33F52
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   63
      Top             =   4935
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   1875
      Picture         =   "CaptureKing.frx":34100
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   62
      Top             =   4980
      Width           =   120
   End
   Begin VB.TextBox txtMessageIn 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Top             =   315
      Width           =   1845
   End
   Begin VB.PictureBox Picture35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2820
      Left            =   1875
      Picture         =   "CaptureKing.frx":3424A
      ScaleHeight     =   2820
      ScaleWidth      =   120
      TabIndex        =   61
      Top             =   330
      Width           =   120
   End
   Begin VB.PictureBox Picture36 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1860
      Picture         =   "CaptureKing.frx":35EAC
      ScaleHeight     =   180
      ScaleWidth      =   135
      TabIndex        =   60
      Top             =   195
      Width           =   135
   End
   Begin VB.PictureBox Picture37 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   150
      Picture         =   "CaptureKing.frx":3603E
      ScaleHeight     =   150
      ScaleWidth      =   1740
      TabIndex        =   59
      Top             =   195
      Width           =   1740
   End
   Begin VB.PictureBox Picture38 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   30
      Picture         =   "CaptureKing.frx":37160
      ScaleHeight     =   150
      ScaleWidth      =   135
      TabIndex        =   58
      Top             =   195
      Width           =   135
   End
   Begin VB.PictureBox Picture39 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2820
      Left            =   30
      Picture         =   "CaptureKing.frx":372BA
      ScaleHeight     =   2820
      ScaleWidth      =   120
      TabIndex        =   57
      Top             =   330
      Width           =   120
   End
   Begin VB.PictureBox Picture40 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   105
      Picture         =   "CaptureKing.frx":38F1C
      ScaleHeight     =   150
      ScaleWidth      =   1815
      TabIndex        =   56
      Top             =   3150
      Width           =   1815
   End
   Begin VB.PictureBox Picture41 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      Picture         =   "CaptureKing.frx":3B286
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   55
      Top             =   3120
      Width           =   135
   End
   Begin VB.PictureBox Picture42 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   1875
      Picture         =   "CaptureKing.frx":3B434
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   54
      Top             =   3150
      Width           =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7815
      Top             =   5280
   End
   Begin VB.PictureBox picBack2 
      BackColor       =   &H00886D1E&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   2175
      ScaleHeight     =   135
      ScaleWidth      =   4935
      TabIndex        =   52
      Top             =   5040
      Width           =   4935
   End
   Begin MSWinsockLib.Winsock SocketGame 
      Left            =   11010
      Top             =   1965
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   5
      LocalPort       =   5
   End
   Begin VB.TextBox txtGameOut 
      Enabled         =   0   'False
      Height          =   375
      Left            =   10965
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   49
      Top             =   1515
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtGameIn 
      Enabled         =   0   'False
      Height          =   375
      Left            =   10965
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   48
      Top             =   1155
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock SocketChat 
      Left            =   11520
      Top             =   1950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   10
      LocalPort       =   10
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   10950
      TabIndex        =   47
      Text            =   "172.17.0.41"
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdClearSend 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   11235
      TabIndex        =   46
      Top             =   2760
      Width           =   735
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   12165
      TabIndex        =   45
      Top             =   2775
      Width           =   975
      Visible         =   0   'False
   End
   Begin VB.PictureBox picboard 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   0
      Left            =   2175
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdSurrender 
      Caption         =   "Surrender"
      Height          =   615
      Left            =   6645
      TabIndex        =   40
      Top             =   5280
      Width           =   1800
   End
   Begin VB.CommandButton cmdCallCheat 
      Caption         =   "Call Invalid Move"
      Height          =   495
      Left            =   12435
      TabIndex        =   39
      Top             =   4275
      Width           =   975
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdCallCheckMate 
      Caption         =   "Call Check Mate"
      Height          =   495
      Left            =   11355
      TabIndex        =   38
      Top             =   4275
      Width           =   1095
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdCallCheck 
      Caption         =   "Call Check"
      Height          =   495
      Left            =   10635
      TabIndex        =   37
      Top             =   4275
      Width           =   735
      Visible         =   0   'False
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   31
      Left            =   9255
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   36
      Top             =   4440
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   30
      Left            =   9255
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   35
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   29
      Left            =   9255
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   34
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   28
      Left            =   9255
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   33
      Top             =   2640
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   27
      Left            =   9255
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   2040
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   26
      Left            =   9255
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   31
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   25
      Left            =   9255
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   30
      Top             =   840
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   24
      Left            =   9255
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   29
      Top             =   240
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   23
      Left            =   8655
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   28
      Top             =   4440
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   22
      Left            =   8655
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   27
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   21
      Left            =   8655
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   26
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   20
      Left            =   8655
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   25
      Top             =   2640
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   19
      Left            =   8655
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   24
      Top             =   2040
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   18
      Left            =   8655
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   23
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   17
      Left            =   8655
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   22
      Top             =   840
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   16
      Left            =   8655
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   240
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   15
      Left            =   8055
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   4440
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   14
      Left            =   8055
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   19
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   13
      Left            =   8055
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   12
      Left            =   8055
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   2640
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   11
      Left            =   8055
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   2040
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   10
      Left            =   8055
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   9
      Left            =   8055
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   14
      Top             =   840
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   8
      Left            =   8055
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   7
      Left            =   7455
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   4440
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   6
      Left            =   7455
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   3840
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   5
      Left            =   7455
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   4
      Left            =   7455
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   3
      Left            =   7455
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      Left            =   7455
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   7455
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.PictureBox picTAKEN 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   7455
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.PictureBox picTake 
      Height          =   615
      Left            =   12495
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   405
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtTest 
      Height          =   1095
      Left            =   7215
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12825
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBLANK 
      Height          =   255
      Left            =   13095
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   405
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog cdlgFile 
      Left            =   12255
      Top             =   1995
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSELECT 
      Height          =   615
      Left            =   11895
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   405
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picBack4 
      BackColor       =   &H00886D1E&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   7095
      ScaleHeight     =   5175
      ScaleWidth      =   135
      TabIndex        =   51
      Top             =   0
      Width           =   135
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   5040
         Y2              =   120
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00886D1E&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   2055
      ScaleHeight     =   5175
      ScaleWidth      =   135
      TabIndex        =   42
      Top             =   0
      Width           =   135
      Begin VB.Line Line4 
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   2400
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   120
         X2              =   120
         Y1              =   5040
         Y2              =   120
      End
   End
   Begin VB.PictureBox picback3 
      BackColor       =   &H00886D1E&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2115
      ScaleHeight     =   255
      ScaleWidth      =   5055
      TabIndex        =   50
      Top             =   -180
      Width           =   5055
   End
   Begin VB.Label lblColor 
      BackStyle       =   0  'Transparent
      Caption         =   "white"
      Height          =   255
      Left            =   11085
      TabIndex        =   53
      Top             =   900
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   2175
      X2              =   7095
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line5 
      BorderWidth     =   5
      X1              =   7095
      X2              =   2175
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   6000
      Width           =   6000
      ForeColor       =   &H00FFFFFF&
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuBoardColour 
         Caption         =   "Board Colour"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuChat 
         Caption         =   "Chat Room"
      End
      Begin VB.Menu mnuGrave 
         Caption         =   "Grave yard"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuStats 
         Caption         =   "Statistics"
      End
      Begin VB.Menu mnuAchievments 
         Caption         =   "Achievments"
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuLocalGame 
         Caption         =   "New Local Game"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewgame 
         Caption         =   "New LAN Game"
      End
      Begin VB.Menu mnuSavegame 
         Caption         =   "Save Game"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Game As"
      End
      Begin VB.Menu mnuLoadgame 
         Caption         =   "Load Game"
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameAdvanced 
         Caption         =   "&Advanced"
         Begin VB.Menu mnuSetOpIP 
            Caption         =   "Set Oppenent IP"
         End
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim BoardState(7, 7) As String  'Row, Column
Dim MoveFrom As Long, MoveTo As Long
Const white = &H80000009
Const black = &H80000007
'=====variables for game itself
Dim MovesTaken As Integer       'number of moves taken
Dim playerturn As String        'value of W or B, which player's turn is it?
Dim LANMessageOut As String
Dim LANmessageIn As String
Dim checkmatecalled As Boolean  'when local player calls checkmate this is
                                'set to true; stopping both player from loosing
Dim PingCount As Integer        'checks to see if other user is activly reciving messages
Dim playercolor As String       'the color of the local player
'for local games
Dim playername As String

Dim localOrLAN As Boolean 'if 0 then local
'==============================

Dim MyUserID As String
Dim grave As Integer
Dim ticked As Boolean
Dim userlevel As Integer


'====================DECLARES MOUSEMOVE=========================
'WILL BE USED LATER FOR HELP DOCUMENTATION
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Sub cmdCallCheat_Click()
checkcheck = MsgBox("Call an invalid move?", vbYesNo)

If checkcheck = vbYes Then
    If MovesTaken = 0 Then
        SocketGame.SendData "Cheat!"
    Else
        MsgBox ("No moves yet taken!")
    End If
    
End If
End Sub

Private Sub cmdCallCheck_Click()
checkcheck = MsgBox("Call Check?", vbYesNo)

If checkcheck = vbYes Then
    SocketGame.SendData "Check!"
End If

End Sub

Private Sub cmdCallCheckMate_Click()
checkMatecheck = MsgBox("Call Check Mate?", vbYesNo)

If checkMatecheck = vbYes Then
    checkmatecalled = True
    SocketGame.SendData "CheckMate!"
End If

End Sub

Private Sub cmdSend_Click()
'Enables user to chat to oppenent
LANMessageOut = txtMessageOut
SocketChat.SendData LANMessageOut

txtMessageOut.Text = ""
txtMessageOut.SetFocus

End Sub

Private Sub cmdSurrender_Click()
answer = MsgBox("There is no gain in surrender, are you sure you want to fall before you opponent?", vbYesNo)

On Error GoTo MonkeyErr

If answer = vbYes Then
    Call AddSurrender
    Kill App.Path & "\game.tmp"
    Call SetBoardState
Else
Exit Sub
End If

MonkeyErr:
If Err.Number = "13" Then
End If

End Sub


Private Sub Form_Load()
Dim C As Long
Dim sBuffer As String
Dim lSize As Long
Dim moves As Long

On Error GoTo errorhandler

frmGame.BackColor = 8940830
picBack2.Top = picBack2.Top + 50

picback3.Top = picback3.Top + 20

'LOCAL GAMES---
localOrLAN = True

'--------------
ticked = True

Open App.Path & "\User\User.txt" For Input As #2

Input #2, C
    userlevel = C
    MyUserID = frmChat.lblUser

Close #2

frmGame.Caption = MyUserID & " [level " & userlevel & "]     - Capture King"

'Open App.Path & "\user\settings.cpt" For Input As #1
'Input #1, C
'CommonDialog1.Color = C
'frmGame.BackColor = C
'picBack.BackColor = C
'picBack2.BackColor = C
'picback3.BackColor = C
'picBack4.BackColor = C
'frmPawnPromotion.BackColor = C
'frmStats.BackColor = C
'frmChat.lblMessageHere.BackColor = C
'frmChat.lblUserMessages.BackColor = C
'Close #1



errorhandler:
If Err.Number > 0 Then
    MsgBox (Err.Number & "  " & Err.Description)
    CommonDialog1.Color = &H80000006
End If

SocketChat.RemoteHost = txtIP
SocketChat.SendData ""
SocketGame.RemoteHost = txtIP
SocketGame.SendData ""

'sets graveyard
grave = 0

'opens save game file
'Open App.Path & "\BoardState" For Output As #1

'loads board and sets out board
Call BoardLoad
Call SetBoardState

'----creates temperary game file
Open App.Path & "\game.tmp" For Output As #3
    For R = 0 To 7
        For C = 0 To 7
            Print #3, BoardState(R, C)
        Next C
    Next R
Close #3
'---end-------------------------

SocketGame.SendData ""



End Sub
Public Function IsEven(ByVal C As Integer) As Boolean

If C Mod 2 = 0 Then
    IsEven = True
Else
    IsEven = False
End If
End Function

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveMouse 200, 200
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picSendMess.Picture = picCMDOUT.Picture
picCmdCheckOVER.Picture = picCmdCheckOut.Picture
picCheckMate.Picture = picCheckMateOut.Picture
picInvalidMove.Picture = picInvalidMoveOut.Picture
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SocketGame.SendData "leftroom"

Open App.Path & "\user\settings.cpt" For Output As #10
    Print #10, CommonDialog1.Color
Close #5

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo monkey
Kill App.Path & "\game.tmp"
End

Exit Sub
monkey:
If Err.Number = 55 Then
    End
End If
End Sub

Private Sub lblColor_Change()
playercolor = lblColor.Caption
End Sub

Private Sub mnuabout_Click()
MsgBox ("Capture King © - Created, designed and coded by Christopher McKirgan 2007-2008. Designed for A2 computing coursework at Bridgwater College." & vbNewLine & vbNewLine & "Contact: kirgy9@gmail.com, www.mckirgan.com." & vbNewLine & vbNewLine & "This game is open to the public for sharing and distrubution, this game MUST NOT BE SOLD without owners prior concent.")
End Sub

Private Sub mnuAchievments_Click()
frmAchievments.Visible = True
End Sub

Private Sub mnuBKColour_Click()
'On Error GoTo errorhandler
'
'CommonDialog1.CancelError = True
'CommonDialog1.ShowColor
'frmGame.BackColor = CommonDialog1.Color
'picBack.BackColor = CommonDialog1.Color
'picBack2.BackColor = C
'picback3.BackColor = C
'picBack4.BackColor = C
'frmChat.BackColor = CommonDialog1.Color
'frmChat.lblOnline.BackColor = CommonDialog1.Color
'frmChat.lblPrivate.BackColor = CommonDialog1.Color
'frmChat.lblMessageHere.BackColor = CommonDialog1.Color
'frmChat.lblUserMessages.BackColor = CommonDialog1.Color
'frmPawnPromotion.BackColor = CommonDialog1.Color
'
'errorhandler:
'If Err.Number = 32755 Then
'End If

End Sub

Private Sub mnuBoardColour_Click()
On Error GoTo errorhandler

MsgBox ("Please choose the first colour for the board")
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
color1 = CommonDialog1.Color

MsgBox ("Please choose the second colour for the board")
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
color2 = CommonDialog1.Color

For i = 0 To 63
RowC = i \ 8
CelC = i Mod 8
    If IsEven(CelC + RowC) Then
        picboard(i).BackColor = color1
    Else
        picboard(i).BackColor = color2
    End If
Next i

'when cancel is hit on change colour
errorhandler:
If Err.Number = 32755 Then
End If

End Sub

Private Sub mnuChat_Click()
frmChat.Visible = True
End Sub

Private Sub mnuExit_Click()
df = MsgBox("Are you sure you want to exit Capture King and leave the current game? Any unsaved progress will be lost.", vbYesNo)
If df = vbYes Then
    End
Else
    Exit Sub
End If

End Sub

Private Sub mnuGrave_Click()
'==checks the option to make the graveyard invisible/visible==
If ticked = True Then
    ticked = False
Else
    ticked = True
End If

mnuGrave.Checked = ticked
For i = 0 To 31
    picTAKEN(i).Visible = ticked
Next i

End Sub

Private Sub mnuHelp_Click()
MoveMouse 200, 200
End Sub

Private Sub mnuLoadgame_Click()
cdlgFile.Filter = "Capture King files only|*.cpt*"


End Sub

Private Sub mnuLocalGame_Click()
playermove = ""
playername = ""

df = MsgBox("Are you sure you want to start a new Local game and exit this one? All unsaved game data will be deleted.", vbYesNo)
If df = vbYes Then
    Kill App.Path & "\game.tmp"
    Call SetBoardState
    localOrLAN = False
    playerturn = "W"
'----creates temperary game file
Open App.Path & "\game.tmp" For Output As #3
    For R = 0 To 7
        For C = 0 To 7
            Print #3, BoardState(R, C)
        Next C
    Next R
Close #3
'---end-------------------------

    
'--------load images into array of squares------
'Loads PAWNS
For i = 8 To 15
    picboard(i) = LoadPicture(App.Path & "\img\b1.ico")
    picboard(i + 40) = LoadPicture(App.Path & "\img\w1.ico")
Next i

'Loads BLACK - PAWNS
picboard(0) = LoadPicture(App.Path & "\img\b2.ico")
picboard(1) = LoadPicture(App.Path & "\img\b3.ico")
picboard(2) = LoadPicture(App.Path & "\img\b4.ico")
picboard(3) = LoadPicture(App.Path & "\img\b5.ico")
picboard(4) = LoadPicture(App.Path & "\img\b6.ico")
picboard(5) = LoadPicture(App.Path & "\img\b4.ico")
picboard(6) = LoadPicture(App.Path & "\img\b3.ico")
picboard(7) = LoadPicture(App.Path & "\img\b2.ico")

'Loads blanks
For i = 16 To 47
    picboard(i).Picture = picBLANK.Picture
Next i

'Loads WHITE - PAWNS
picboard(56) = LoadPicture(App.Path & "\img\w2.ico")
picboard(57) = LoadPicture(App.Path & "\img\w3.ico")
picboard(58) = LoadPicture(App.Path & "\img\w4.ico")
picboard(59) = LoadPicture(App.Path & "\img\w6.ico")
picboard(60) = LoadPicture(App.Path & "\img\w5.ico")
picboard(61) = LoadPicture(App.Path & "\img\w4.ico")
picboard(62) = LoadPicture(App.Path & "\img\w3.ico")
picboard(63) = LoadPicture(App.Path & "\img\w2.ico")


lblInfo.Caption = "A local game is being set up..."

PlayernameError:
playername = InputBox("Please enter name of [Player 2]", "Player 2 name needed")
If playername = "" Then
    GoTo PlayernameError
End If

PlayerMoveError:
playermove = MsgBox("Will " & playername & " be taking the first move?", vbYesNo)
If playermove = "" Then GoTo PlayerMoveError
Print playermove

If playermove = 7 Then
    MsgBox (MyUserID & " will be taking the first move as the white team")
    lblInfo.Caption = MyUserID & " is taking their move."
Else
    MsgBox (playername & " will be taking the first move as the white team")
    lblInfo.Caption = playername & " is taking their move."
End If



Else
    Exit Sub
End If



End Sub

Private Sub mnuNewgame_Click()
df = MsgBox("Are you sure you want to start a new game and exit this one?", vbYesNo)
If df = vbYes Then
    Kill App.Path & "\game.tmp"
    Call SetBoardState
    localOrLAN = True

    
'--------load images into array of squares------
'Loads PAWNS
For i = 8 To 15
    picboard(i) = LoadPicture(App.Path & "\img\b1.ico")
    picboard(i + 40) = LoadPicture(App.Path & "\img\w1.ico")
Next i

'Loads BLACK - PAWNS
picboard(0) = LoadPicture(App.Path & "\img\b2.ico")
picboard(1) = LoadPicture(App.Path & "\img\b3.ico")
picboard(2) = LoadPicture(App.Path & "\img\b4.ico")
picboard(3) = LoadPicture(App.Path & "\img\b5.ico")
picboard(4) = LoadPicture(App.Path & "\img\b6.ico")
picboard(5) = LoadPicture(App.Path & "\img\b4.ico")
picboard(6) = LoadPicture(App.Path & "\img\b3.ico")
picboard(7) = LoadPicture(App.Path & "\img\b2.ico")

'Loads blanks
For i = 16 To 47
    picboard(i).Picture = picBLANK.Picture
Next i

'Loads WHITE - PAWNS
picboard(56) = LoadPicture(App.Path & "\img\w2.ico")
picboard(57) = LoadPicture(App.Path & "\img\w3.ico")
picboard(58) = LoadPicture(App.Path & "\img\w4.ico")
picboard(59) = LoadPicture(App.Path & "\img\w6.ico")
picboard(60) = LoadPicture(App.Path & "\img\w5.ico")
picboard(61) = LoadPicture(App.Path & "\img\w4.ico")
picboard(62) = LoadPicture(App.Path & "\img\w3.ico")
picboard(63) = LoadPicture(App.Path & "\img\w2.ico")

Else
    Exit Sub
End If

End Sub

Private Sub mnuSaveAs_Click()

'filters the files to CPT
cdlgFile.Filter = "Capture King files only|*.cpt*"
'shows the save diologue box
cdlgFile.ShowSave

Open cdlgFile.FileName & ".cpt" For Output As #4
    For R = 0 To 7
        For C = 0 To 7
            Print #4, BoardState(R, C)
        Next C
    Next R

Close #4


End Sub

Private Sub mnuSavegame_Click()

For R = 0 To 7
    For C = 0 To 7
        Print #1, BoardState(R, C)
    Next C
Next R

Open App.Path & "\game.tmp" For Output As #3
    
    For R = 0 To 7
        For C = 0 To 7
            Print #3, BoardState(R, C)
        Next C
    Next R

Close #3

End Sub

Private Sub mnuSetOpIP_Click()
Dim OpIP As String

OpIP = InputBox("Please enter a valid IP address")
txtIP = OpIP
SocketChat.RemoteHost = OpIP
SocketGame.RemoteHost = OpIP

End Sub

Private Sub mnuStats_Click()
frmStats.Visible = True
End Sub

Private Sub pic_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

End Sub

Private Sub picBack_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Call ReturnToBoard
End Sub


Private Sub picBack2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Call ReturnToBoard
End Sub

Private Sub picback3_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Call ReturnToBoard
End Sub

Private Sub picBack4_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Call ReturnToBoard
End Sub

Private Sub picboard_Click(Index As Integer)
Dim colorname As String

If playercolor = "white" Then
    colorname = "Black"
Else
    colorname = "White"
End If

For i = 0 To 63
    If i = 0 Then
        If playerturn = playercolor Then
            MsgBox ("Please wait your turn." & colorname & " is taking their turn.")
        End If
    End If
Next i
End Sub

Private Sub picboard_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

MoveTo = Index
picTake.Picture = picboard(MoveTo)
picboard(Index) = picSELECT

taken = BoardState(MoveTo \ 8, MoveTo Mod 8)
movingpiece = BoardState(MoveFrom \ 8, MoveFrom Mod 8)


'////---------LOCAL LAN GAME--------------/////


Dim MessageGameOut As String
Dim movefromSEND As String
Dim movetoSEND As String

movetoSEND = MoveTo
movefromSEND = MoveFrom

'@@@@@@@@@@@@@@@@@ SENDS MOVES ACCROSS NETWORK @@@@@@@@@@@@@@@@@@@@
If localOrLAN = True Then
    If Len(CStr(movefromSEND)) = 1 Then
        movefromSEND = "0" & movefromSEND
    Else
        movefromSEND = movefromSEND
    End If

    If Len(CStr(movetoSEND)) = 1 Then
        movetoSEND = "0" & movetoSEND
    Else
        movetoSEND = movetoSEND
    End If

    MessageGameOut = movefromSEND & ">" & movetoSEND

    SocketGame.SendData MessageGameOut
End If
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@






TESTINGON = False
If TESTINGON = True Then

Else
        '==============FOR LAN GAMES============================
        If localOrLAN = False Then
                If Left(movingpiece, 1) <> playerturn Then
                    MsgBox ("You cannot move your oppenents pieces fool!!")
                    picboard(MoveTo).Picture = picTake.Picture
                    picboard(MoveFrom).Picture = picSELECT.Picture
                Exit Sub
                End If
        End If
        '==============ABOVE IS LAN GAMES=======================
        
        If movingpiece = "" Then
                picboard(MoveTo).Picture = picTake.Picture
                picboard(MoveFrom).Picture = picSELECT.Picture
            Exit Sub
        Else
        
        For i = 0 To 7
            If picboard(MoveTo).Index = i And Mid(movingpiece, 1, 2) = "WP" Then
                frmPawnPromotion.Visible = True
            End If
        Next i
        
        For i = 55 To 63
            If picboard(MoveTo).Index = i And Mid(movingpiece, 1, 2) = "BP" Then
                frmPawnPromotion.Visible = True
            End If
        Next i
        
        
        
            If Left(taken, 1) = Left(movingpiece, 1) Then
                MsgBox ("You cannot take your own piece!!")
                picboard(MoveTo).Picture = picTake.Picture
                picboard(MoveFrom).Picture = picSELECT.Picture
            Exit Sub
            End If
        
        End If
        'TAKEN PIECES MOVE TO GRAVEYARD
        If movingpiece <> "" And taken <> "" Then
            picTAKEN(grave).Picture = picTake.Picture
            grave = grave + 1
        End If
        
        '--------------------------
        BoardState(MoveTo \ 8, MoveTo Mod 8) = BoardState(MoveFrom \ 8, MoveFrom Mod 8)
        Print BoardState(MoveTo \ 8, MoveTo Mod 8) & " moved from cell " & MoveFrom & " to cell " & MoveTo & " holding " & taken
        'prints piece name AND cell moved from address AND cell moved to address AND piece taken name
End If

If MoveTo = MoveFrom Then
    Exit Sub
Else
    BoardState(MoveFrom \ 8, MoveFrom Mod 8) = ""
    Call SaveBoardState
End If

moves = moves + 1

If playeturn = "W" Then
    playerturn = "B"
Else
    playerturn = "W"
End If

MovesTaken = MovesTaken + 1

End Sub

Private Sub picboard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Enables ability to drap pieces
MoveFrom = Index
picSELECT = picboard(Index)
picboard(Index) = LoadPicture()
picboard(Index).DragIcon = picSELECT
picboard(Index).Drag


End Sub

Private Sub BoardLoad()

'loads an array of image boxes in a grid of 8 by 8
For i = 1 To 63
        Load picboard(i)
        picboard(i).Visible = True

    If i Mod 8 = 0 Then
        picboard(i).Left = picboard(0).Left
        picboard(i).Top = picboard(i - 1).Top + picboard(i - 1).Height
    Else
        picboard(i).Left = picboard(i - 1).Left + picboard(i - 1).Width
        picboard(i).Top = picboard(i - 1).Top
    End If
Next i


'colours black and white
For i = 0 To 63
RowC = i \ 8
CelC = i Mod 8
    If IsEven(CelC + RowC) Then
        picboard(i).BackColor = black
    Else
        picboard(i).BackColor = white
    End If
Next i

'--------load images into array of squares------
'Loads PAWNS
For i = 8 To 15
    picboard(i) = LoadPicture(App.Path & "\img\b1.ico")
    picboard(i + 40) = LoadPicture(App.Path & "\img\w1.ico")
Next i

'Loads BLACK - PAWNS
picboard(0) = LoadPicture(App.Path & "\img\b2.ico")
picboard(1) = LoadPicture(App.Path & "\img\b3.ico")
picboard(2) = LoadPicture(App.Path & "\img\b4.ico")
picboard(3) = LoadPicture(App.Path & "\img\b5.ico")
picboard(4) = LoadPicture(App.Path & "\img\b6.ico")
picboard(5) = LoadPicture(App.Path & "\img\b4.ico")
picboard(6) = LoadPicture(App.Path & "\img\b3.ico")
picboard(7) = LoadPicture(App.Path & "\img\b2.ico")

'Loads WHITE - PAWNS
picboard(56) = LoadPicture(App.Path & "\img\w2.ico")
picboard(57) = LoadPicture(App.Path & "\img\w3.ico")
picboard(58) = LoadPicture(App.Path & "\img\w4.ico")
picboard(59) = LoadPicture(App.Path & "\img\w6.ico")
picboard(60) = LoadPicture(App.Path & "\img\w5.ico")
picboard(61) = LoadPicture(App.Path & "\img\w4.ico")
picboard(62) = LoadPicture(App.Path & "\img\w3.ico")
picboard(63) = LoadPicture(App.Path & "\img\w2.ico")

'Loads lines around board

For i = 0 To 63
'picboard(i).Visible = false
Next i


End Sub

Private Sub SetBoardState()
BoardState(0, 0) = "BR1"
BoardState(0, 1) = "BK1"
BoardState(0, 2) = "BB1"
BoardState(0, 3) = "B+1"
BoardState(0, 4) = "BQ1"
BoardState(0, 5) = "BB2"
BoardState(0, 6) = "BK2"
BoardState(0, 7) = "BR2"
BoardState(1, 0) = "BP1"
BoardState(1, 1) = "BP2"
BoardState(1, 2) = "BP3"
BoardState(1, 3) = "BP4"
BoardState(1, 4) = "BP5"
BoardState(1, 5) = "BP6"
BoardState(1, 6) = "BP7"
BoardState(1, 7) = "BP8"
'set up blanks in loop!
'for boardstate'
For i = 2 To 5
    For j = 0 To 7
        BoardState(i, j) = ""
    Next j
Next i

BoardState(7, 0) = "WR1"
BoardState(7, 1) = "WK1"
BoardState(7, 2) = "WB1"
BoardState(7, 3) = "WQ1"
BoardState(7, 4) = "W+1"
BoardState(7, 5) = "WB2"
BoardState(7, 6) = "WK2"
BoardState(7, 7) = "WR2"
BoardState(6, 0) = "WP1"
BoardState(6, 1) = "WP2"
BoardState(6, 2) = "WP3"
BoardState(6, 3) = "WP4"
BoardState(6, 4) = "WP5"
BoardState(6, 5) = "WP6"
BoardState(6, 6) = "WP7"
BoardState(6, 7) = "WP8"

playerturn = "white"
playercolor = lblColor.Caption

End Sub

Private Sub SaveBoardState()
Dim PreviousGame As String
Dim PreviousGameTemp As String
Dim MoveNumber As Integer
'below writes all moves made to a file

Open App.Path & "\game.tmp" For Input As #14
Do Until EOF(14)
    Input #14, PreviousGameTemp
    PreviousGame = PreviousGame & PreviousGameTemp & vbNewLine
Loop
Close #14

Open App.Path & "\game.tmp" For Output As #3
    Print #3, PreviousGame
    For R = 0 To 7
        For C = 0 To 7
            Print #3, BoardState(R, C)
        Next C
    Next R

Close #3

End Sub






Sub MoveMouse(X As Single, Y As Single)
Dim pt As POINTAPI

    pt.X = X
    pt.Y = Y
    ClientToScreen hwnd, pt
    SetCursorPos pt.X, pt.Y
End Sub




Private Sub picCheckMate_Click()
Call cmdCallCheckMate_Click
End Sub

Private Sub picCheckMate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picCheckMate.Picture = picCheckMateOver.Picture
End Sub

Private Sub picCmdCheckOVER_Click()
Call cmdCallCheck_Click
End Sub

Private Sub picCmdCheckOVER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picCmdCheckOVER.Picture = picCheckOUT.Picture
End Sub

Private Sub picInvalidMove_Click()
Call cmdCallCheat_Click
End Sub

Private Sub picInvalidMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picInvalidMove.Picture = picInvalidMoveOver.Picture
End Sub

Private Sub picSendMess_Click()
Call cmdSend_Click
End Sub

Private Sub picSendMess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picSendMess.Picture = picCMDOVER.Picture
End Sub

Private Sub SocketChat_DataArrival(ByVal bytesTotal As Long)
SocketChat.GetData LANmessageIn
txtMessageIn = txtMessageIn & LANmessageIn & vbNewLine

End Sub

Private Sub SocketGame_DataArrival(ByVal bytesTotal As Long)
Dim messageGameIn As String
Dim movetoRECIEVED As String
Dim movefromRECIEVED As String
Dim movefromRECIEVEDint As Integer
Dim movetoRECIEVEDint As Integer


On Error Resume Next
SocketGame.GetData messageGameIn
txtGameIn = messageGameIn

'//SYSTEM MESSAGES FROM NETWORK
'when opponent leaves

If txtGameIn = "leftroom" Then
    MsgBox ("Opponent has left the room!")
    Exit Sub
End If

'when undo is approved
If messageGameIn = "Undo!" Then
    Call UndoMove
    Exit Sub
End If

'when cheat is called
If messageGameIn = "Cheat!" Then
    Cheater = MsgBox("Cheat has been called! Check that move, do you want to undo it?", vbYesNo)
    If Cheater = vbYes Then
        Call UndoMove
        SocketGame.SendData "Undo!"
        Exit Sub
    Else
        Exit Sub
    End If
End If

'when check is called
If txtGameIn = "Check!" Then
    MsgBox ("Check has been called!")
    Exit Sub
End If

'checks win
If txtGameIn = "YouWin!" Then
    MsgBox ("Congradulations, you took the king, you win!")
    Call AddWin
    Exit Sub
End If
    


'When checkmate is called
If txtGameIn = "CheckMate!" And checkmatecalled = False Then
checkmatecalled = False

    endinggame = MsgBox("Check Mate has been called! Is your opponent correct in calling check mate?", vbYesNo)
    If endinggame = vbYes Then
        MsgBox ("Your opponent wins...better luck next time. ")
        SocketGame.SendData "YouWin!"
        'Call Addloss
        Exit Sub
    Else
    Exit Sub
    End If
Else

End If

'//SYSTEM MESSAGES ABOVE

If Mid(messageGameIn, 1, 1) = "0" Then
    movefromRECIEVED = Mid(messageGameIn, 2, 1)
    movefromRECIEVEDint = Int(movefromRECIEVED)
'    MsgBox movefromRECIEVEDint
Else
    movefromRECIEVED = Mid(messageGameIn, 1, 2)
    movefromRECIEVEDint = Int(movefromRECIEVED)
End If
If Mid(messageGameIn, 4, 1) = "0" Then
    movetoRECIEVED = Mid(messageGameIn, 5, 1)
    movetoRECIEVEDint = Int(movetoRECIEVED)
'    MsgBox movefromRECIEVEDint
Else
    movetoRECIEVED = Mid(messageGameIn, 4, 2)
    movetoRECIEVEDint = Int(movetoRECIEVED)
End If


'///////////////////////////////

MoveTo = movetoRECIEVEDint
MoveFrom = movefromRECIEVEDint
picTake.Picture = picboard(MoveTo)
'picboard(MoveFrom) = picSELECT

taken = BoardState(MoveTo \ 8, MoveTo Mod 8)
movingpiece = BoardState(MoveFrom \ 8, MoveFrom Mod 8)

picboard(MoveTo).Picture = picboard(MoveFrom).Picture
picboard(MoveFrom).Picture = picBLANK.Picture

'BoardState(MoveTo \ 8, MoveTo Mod 8) = BoardState(MoveFrom \ 8, MoveFrom Mod 8)
'Print BoardState(MoveTo \ 8, MoveTo Mod 8) & " moved from cell " & MoveFrom & " to cell " & MoveTo & " holding " & taken
''prints piece name AND cell moved from address AND cell moved to address AND piece taken name
'BoardState(MoveFrom \ 8, MoveFrom Mod 8) = ""

BoardState(MoveTo \ 8, MoveTo Mod 8) = BoardState(MoveFrom \ 8, MoveFrom Mod 8)
'Print BoardState(MoveTo \ 8, MoveTo Mod 8) & " moved from cell " & MoveFrom & " to cell " & MoveTo & " holding " & taken
'prints piece name AND cell moved from address AND cell moved to address AND piece taken name


''==============FOR LAN GAMES============================
If localOrLAN = False Then
        If Left(movingpiece, 1) <> playerturn Then
            MsgBox ("You cannot move your oppenents pieces fool!!")
            picboard(MoveTo).Picture = picTake.Picture
            picboard(MoveFrom).Picture = picSELECT.Picture
        Exit Sub
        End If
End If
'==============ABOVE IS LAN GAMES=======================

If movingpiece = "" Then
        picboard(MoveTo).Picture = picTake.Picture
        picboard(MoveFrom).Picture = picSELECT.Picture
    Exit Sub
Else


    If Left(taken, 1) = Left(movingpiece, 1) Then
        MsgBox ("You cannot take your own piece!!")
        picboard(MoveTo).Picture = picTake.Picture
        picboard(MoveFrom).Picture = picSELECT.Picture
    Exit Sub
    End If

End If
'TAKEN PIECES MOVE TO GRAVEYARD
If movingpiece <> "" And taken <> "" Then
    picTAKEN(grave).Picture = picTake.Picture
    grave = grave + 1
End If

'--------------------------


If MoveTo = MoveFrom Then
    Exit Sub
Else
    BoardState(MoveFrom \ 8, MoveFrom Mod 8) = ""
    Call SaveBoardState
End If

moves = moves + 1

If playeturn = "W" Then
    playerturn = "B"
Else
    playerturn = "W"
End If

'////---------LOCAL LAN GAME--------------/////
'Dim MessageGameOut As String
'Dim movefromSEND As String
'Dim movetoSEND As String
'
'movetoSEND = MoveTo
'movefromSEND = MoveFrom
'
'If localOrLAN = True Then
'    If Len(CStr(movefromSEND)) = 1 Then
'        movefromSEND = "0" & movefromSEND
'    Else
'        movefromSEND = movefromSEND
'    End If
'
'    If Len(CStr(movetoSEND)) = 1 Then
'        movetoSEND = "0" & movetoSEND
'    Else
'        movetoSEND = movetoSEND
'    End If
'
'    MessageGameOut = movefromSEND & ">" & movetoSEND
'
'    SocketGame.SendData MessageGameOut
'End If

'////////////////////////////////////

MovesTaken = MovesTaken + 1

End Sub


Public Sub PawnPromotion()
picboard(MoveTo).Picture = frmPawnPromotion.picPromotepiece.Picture
End Sub

Public Sub AddWin()

'adds a win to stats from file - also updates stats form
Open App.Path & "\User\User.txt" For Input As #8
    Input #8, C
        dumpit = C
    Input #8, C
        wins = C + 1
    Input #8, C
        losses = C
    Input #8, C
        games = C + 1
Close #8

userXP = games + wins

frmStats.txtstats = MyUserID & vbNewLine & _
"User XP:" & vbTab & vbTab & userXP & vbNewLine & _
"Wins:" & vbTab & vbTab & vbTab & wins & vbNewLine & _
"Loses:" & vbTab & vbTab & vbTab & losses & vbNewLine & _
"Games:" & vbTab & vbTab & games & vbNewLine

'Prints stats to file
Open App.Path & "\User\User.txt" For Output As #9
Dim StatDump As String
StatDump = MyUserID & vbNewLine & games & vbNewLine & wins & vbNewLine & losses & vbNewLine
    Print #9, StatDump
Close #9

End Sub

Public Sub Addloss()

'adds a loss to stats from file - also updates stats form
Open App.Path & "\User\User.txt" For Input As #9
    Input #8, C
        userlevel = C
    Input #8, C
        wins = C
    Input #8, C
        loses = C + 1
    Input #8, C
        surrenders = C
Close #8
frmStats.txtstats = MyUserID & vbNewLine & _
"User Level:" & vbTab & vbTab & userlevel & vbNewLine & _
"Wins:" & vbTab & vbTab & vbTab & wins & vbNewLine & _
"loses:" & vbTab & vbTab & vbTab & loses & vbNewLine & _
"Surrenders:" & vbTab & vbTab & surrenders & vbNewLine

'Prints stats to file
Open App.Path & "\User\User.txt" For Output As #9
    Print #9, frmStats.txtstats
Close #9

End Sub

Public Sub AddSurrender()

'adds a surrender to stats from file - also updates stats form
'Open App.Path & "\User\User.txt" For Input As #9
'    Input #8, C 'CHANGED
''        userlevel Round((wins / loses) * 100, 0)!!!!!!!!!!!!!!!!!!!!!!!!
'    Input #8, C
'        wins = C
'    Input #8, C
'        loses = C
'    Input #8, C
'        surrenders = C + 1
'Close #8
frmStats.txtstats = MyUserID & vbNewLine & _
"User Level:" & vbTab & vbTab & userlevel & vbNewLine & _
"Wins:" & vbTab & vbTab & vbTab & wins & vbNewLine & _
"loses:" & vbTab & vbTab & vbTab & loses & vbNewLine & _
"Surrenders:" & vbTab & vbTab & surrenders & vbNewLine

'Prints stats to file
Open App.Path & "\User\User.txt" For Output As #9
    Print #9, frmStats.txtstats
Close #9

End Sub

Public Sub ReturnToBoard()
MsgBox ("Please dont drag pieces off the board!")
picboard(MoveFrom).Picture = picSELECT.Picture
End Sub

Private Sub Timer1_Timer()
SocketGame.SendData ""
SocketChat.SendData ""
End Sub

Public Sub UndoMove()
Dim DeadRead As String
Dim MoveUndoPlace As Integer
'movestaken - a single move * number of recorded squares

MoveUndoPlace = ((MovesTaken - 2) * 64)

Open App.Path & "\game.tmp" For Input As #22
    For i = 0 To MoveUndoPlace
            Print #22, DeadRead
    Next i
        
    For R = 0 To 7
        For C = 0 To 7
            Print #22, BoardState(R, C)
        Next C
    Next R

Open App.Path & "\game.tmp" For Output As #3
    For R = 0 To 7
        For C = 0 To 7
            Print #3, BoardState(R, C)
        Next C
    Next R
Close #3


For R = 0 To 7
    For C = 0 To 7
        'WHITE PAWN
        If Mid((BoardState(R, C)), 1, 2) = "WP" Then
            picboard(i) = LoadPicture(App.Path & "\img\w1.ico")
        End If
        'BLACK PAWN
        If Mid((BoardState(R, C)), 1, 2) = "BP" Then
            picboard(i) = LoadPicture(App.Path & "\img\b1.ico")
        End If
        'WHITE ROOK
        If Mid((BoardState(R, C)), 1, 2) = "WR" Then
            picboard(i) = LoadPicture(App.Path & "\img\w2.ico")
        End If
        'BLACK ROOK
        If Mid((BoardState(R, C)), 1, 2) = "BR" Then
            picboard(i) = LoadPicture(App.Path & "\img\b2.ico")
        End If
        'WHITE KING
        If Mid((BoardState(R, C)), 1, 2) = "W+" Then
            picboard(i) = LoadPicture(App.Path & "\img\b2.ico")
        End If
        'BLACK KING
        If Mid((BoardState(R, C)), 1, 2) = "B+" Then
            picboard(i) = LoadPicture(App.Path & "\img\b2.ico")
        End If
        
        If BoardState(R, C) = "WP1" Then
            picboard(i) = LoadPicture(App.Path & "\img\w1.ico")
        End If
        If BoardState(R, C) = "WP1" Then
            picboard(i) = LoadPicture(App.Path & "\img\w1.ico")
        End If
        If BoardState(R, C) = "WP1" Then
            picboard(i) = LoadPicture(App.Path & "\img\w1.ico")
        End If
        If BoardState(R, C) = "WP1" Then
            picboard(i) = LoadPicture(App.Path & "\img\w1.ico")
        End If
        If BoardState(R, C) = "WP1" Then
            picboard(i) = LoadPicture(App.Path & "\img\w1.ico")
        End If
        If BoardState(R, C) = "WP1" Then
            picboard(i) = LoadPicture(App.Path & "\img\w1.ico")
        End If
        
        
        i = i + 1
        
    Next C
Next R


Call SaveBoardState


Close #15

End Sub





