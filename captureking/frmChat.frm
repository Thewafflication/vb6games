VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmChat 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "v"
   ClientHeight    =   7200
   ClientLeft      =   4545
   ClientTop       =   2145
   ClientWidth     =   13000
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   13000
   Begin VB.PictureBox Picture54 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11070
      Picture         =   "frmChat.frx":08CA
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   111
      Top             =   5145
      Width           =   375
   End
   Begin VB.PictureBox Picture53 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11445
      Picture         =   "frmChat.frx":1078
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   110
      Top             =   5145
      Width           =   375
   End
   Begin VB.PictureBox picTick2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11445
      Picture         =   "frmChat.frx":1826
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   109
      Top             =   4395
      Width           =   375
   End
   Begin VB.PictureBox picCross2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11070
      Picture         =   "frmChat.frx":1FD4
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   108
      Top             =   4395
      Width           =   375
   End
   Begin VB.PictureBox picTick1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11445
      Picture         =   "frmChat.frx":2782
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   107
      Top             =   3645
      Width           =   375
   End
   Begin VB.PictureBox picCross1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11070
      Picture         =   "frmChat.frx":2F30
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   106
      Top             =   3645
      Width           =   375
   End
   Begin VB.PictureBox picInviteAct 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   9885
      Picture         =   "frmChat.frx":36DE
      ScaleHeight     =   705
      ScaleWidth      =   1800
      TabIndex        =   105
      Top             =   2745
      Width           =   1800
   End
   Begin VB.PictureBox picInviteDis 
      Height          =   735
      Left            =   3675
      Picture         =   "frmChat.frx":7938
      ScaleHeight     =   675
      ScaleWidth      =   1830
      TabIndex        =   104
      Top             =   6645
      Width           =   1890
   End
   Begin VB.PictureBox picInvite 
      Height          =   585
      Left            =   2085
      Picture         =   "frmChat.frx":BB92
      ScaleHeight     =   525
      ScaleWidth      =   1440
      TabIndex        =   103
      Top             =   6615
      Width           =   1500
   End
   Begin VB.PictureBox picInviteOver 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   255
      Picture         =   "frmChat.frx":FDEC
      ScaleHeight     =   705
      ScaleWidth      =   1800
      TabIndex        =   102
      Top             =   6615
      Width           =   1800
   End
   Begin VB.PictureBox Picture50 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7890
      Picture         =   "frmChat.frx":14046
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   96
      Top             =   2535
      Width           =   135
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5370
      Picture         =   "frmChat.frx":141F4
      ScaleHeight     =   180
      ScaleWidth      =   135
      TabIndex        =   49
      Top             =   315
      Width           =   135
   End
   Begin VB.TextBox txtstat 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0  'None
      Height          =   2160
      Left            =   8010
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   435
      Width           =   3570
   End
   Begin VB.PictureBox Picture52 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   9630
      Picture         =   "frmChat.frx":14386
      ScaleHeight     =   150
      ScaleWidth      =   1950
      TabIndex        =   98
      Top             =   2580
      Width           =   1950
   End
   Begin VB.PictureBox Picture51 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   7995
      Picture         =   "frmChat.frx":166F0
      ScaleHeight     =   150
      ScaleWidth      =   1950
      TabIndex        =   97
      Top             =   2580
      Width           =   1950
   End
   Begin VB.PictureBox Picture49 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   7890
      Picture         =   "frmChat.frx":18A5A
      ScaleHeight     =   2160
      ScaleWidth      =   120
      TabIndex        =   95
      Top             =   450
      Width           =   120
   End
   Begin VB.PictureBox Picture48 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   7890
      Picture         =   "frmChat.frx":1A6BC
      ScaleHeight     =   150
      ScaleWidth      =   135
      TabIndex        =   94
      Top             =   300
      Width           =   135
   End
   Begin VB.PictureBox Picture47 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   7995
      Picture         =   "frmChat.frx":1A816
      ScaleHeight     =   150
      ScaleWidth      =   2160
      TabIndex        =   93
      Top             =   300
      Width           =   2160
   End
   Begin VB.PictureBox Picture46 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   9435
      Picture         =   "frmChat.frx":1B938
      ScaleHeight     =   150
      ScaleWidth      =   2160
      TabIndex        =   92
      Top             =   300
      Width           =   2160
   End
   Begin VB.PictureBox Picture45 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   11580
      Picture         =   "frmChat.frx":1CA5A
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   91
      Top             =   2565
      Width           =   120
   End
   Begin VB.PictureBox Picture44 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   11580
      Picture         =   "frmChat.frx":1CBA4
      ScaleHeight     =   2220
      ScaleWidth      =   120
      TabIndex        =   90
      Top             =   435
      Width           =   120
   End
   Begin VB.PictureBox Picture43 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   11565
      Picture         =   "frmChat.frx":1E806
      ScaleHeight     =   180
      ScaleWidth      =   135
      TabIndex        =   89
      Top             =   300
      Width           =   135
   End
   Begin VB.PictureBox picCMDOVER2 
      Height          =   855
      Left            =   2490
      Picture         =   "frmChat.frx":1E998
      ScaleHeight     =   795
      ScaleWidth      =   900
      TabIndex        =   88
      Top             =   5910
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picCMDOUT2 
      Height          =   870
      Left            =   3330
      Picture         =   "frmChat.frx":20F4E
      ScaleHeight     =   810
      ScaleWidth      =   1095
      TabIndex        =   87
      Top             =   5910
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox picSendPriv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   4590
      Picture         =   "frmChat.frx":23504
      ScaleHeight     =   705
      ScaleWidth      =   1005
      TabIndex        =   86
      Top             =   5085
      Width           =   1005
   End
   Begin VB.PictureBox picCMDOUT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   1380
      Picture         =   "frmChat.frx":25ABA
      ScaleHeight     =   705
      ScaleWidth      =   1005
      TabIndex        =   85
      Top             =   5865
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picCMDOVER 
      Height          =   900
      Left            =   165
      Picture         =   "frmChat.frx":28070
      ScaleHeight     =   840
      ScaleWidth      =   1515
      TabIndex        =   84
      Top             =   5865
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picSendMess 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   4575
      Picture         =   "frmChat.frx":2A626
      ScaleHeight     =   705
      ScaleWidth      =   1005
      TabIndex        =   83
      Top             =   4230
      Width           =   1005
   End
   Begin VB.PictureBox Picture42 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   7650
      Picture         =   "frmChat.frx":2CBDC
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   82
      Top             =   5670
      Width           =   120
   End
   Begin VB.PictureBox Picture41 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5640
      Picture         =   "frmChat.frx":2CD26
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   81
      Top             =   5640
      Width           =   135
   End
   Begin VB.PictureBox Picture40 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   5730
      Picture         =   "frmChat.frx":2CED4
      ScaleHeight     =   150
      ScaleWidth      =   1950
      TabIndex        =   80
      Top             =   5685
      Width           =   1950
   End
   Begin VB.PictureBox Picture39 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   5640
      Picture         =   "frmChat.frx":2F23E
      ScaleHeight     =   1350
      ScaleWidth      =   120
      TabIndex        =   79
      Top             =   4335
      Width           =   120
   End
   Begin VB.PictureBox Picture38 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   5640
      Picture         =   "frmChat.frx":30EA0
      ScaleHeight     =   150
      ScaleWidth      =   135
      TabIndex        =   78
      Top             =   4185
      Width           =   135
   End
   Begin VB.PictureBox Picture37 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   5775
      Picture         =   "frmChat.frx":30FFA
      ScaleHeight     =   150
      ScaleWidth      =   1890
      TabIndex        =   77
      Top             =   4185
      Width           =   1890
   End
   Begin VB.PictureBox Picture36 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7635
      Picture         =   "frmChat.frx":3211C
      ScaleHeight     =   180
      ScaleWidth      =   135
      TabIndex        =   76
      Top             =   4185
      Width           =   135
   End
   Begin VB.PictureBox Picture35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   7650
      Picture         =   "frmChat.frx":322AE
      ScaleHeight     =   1380
      ScaleWidth      =   120
      TabIndex        =   75
      Top             =   4305
      Width           =   120
   End
   Begin VB.PictureBox Picture34 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   7620
      Picture         =   "frmChat.frx":33F10
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   74
      Top             =   3825
      Width           =   120
   End
   Begin VB.PictureBox Picture33 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   5745
      ScaleHeight     =   105
      ScaleWidth      =   1905
      TabIndex        =   73
      Top             =   3810
      Width           =   1905
   End
   Begin VB.PictureBox Picture32 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   5715
      Picture         =   "frmChat.frx":3405A
      ScaleHeight     =   150
      ScaleWidth      =   1950
      TabIndex        =   72
      Top             =   3840
      Width           =   1950
   End
   Begin VB.PictureBox Picture31 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5625
      Picture         =   "frmChat.frx":363C4
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   71
      Top             =   3810
      Width           =   135
   End
   Begin VB.PictureBox Picture30 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   5625
      Picture         =   "frmChat.frx":36572
      ScaleHeight     =   150
      ScaleWidth      =   135
      TabIndex        =   70
      Top             =   315
      Width           =   135
   End
   Begin VB.PictureBox Picture29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3390
      Left            =   5625
      Picture         =   "frmChat.frx":366CC
      ScaleHeight     =   3390
      ScaleWidth      =   120
      TabIndex        =   69
      Top             =   450
      Width           =   120
   End
   Begin VB.PictureBox Picture28 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   5730
      ScaleHeight     =   30
      ScaleWidth      =   1890
      TabIndex        =   68
      Top             =   465
      Width           =   1890
   End
   Begin VB.PictureBox Picture27 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   5745
      Picture         =   "frmChat.frx":3832E
      ScaleHeight     =   150
      ScaleWidth      =   1890
      TabIndex        =   67
      Top             =   315
      Width           =   1890
   End
   Begin VB.PictureBox Picture26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   7620
      Picture         =   "frmChat.frx":39450
      ScaleHeight     =   3405
      ScaleWidth      =   120
      TabIndex        =   66
      Top             =   495
      Width           =   120
   End
   Begin VB.PictureBox Picture25 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7605
      Picture         =   "frmChat.frx":3B0B2
      ScaleHeight     =   180
      ScaleWidth      =   135
      TabIndex        =   65
      Top             =   315
      Width           =   135
   End
   Begin VB.TextBox txtMySend 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4305
      Width           =   4215
   End
   Begin VB.PictureBox Picture24 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      Picture         =   "frmChat.frx":3B244
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   64
      Top             =   5610
      Width           =   135
   End
   Begin VB.PictureBox Picture23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   120
      Picture         =   "frmChat.frx":3B3F2
      ScaleHeight     =   1260
      ScaleWidth      =   120
      TabIndex        =   63
      Top             =   4380
      Width           =   120
   End
   Begin VB.PictureBox Picture22 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      Picture         =   "frmChat.frx":3D054
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   62
      Top             =   4200
      Width           =   315
   End
   Begin VB.PictureBox Picture21 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   270
      Picture         =   "frmChat.frx":3D3D6
      ScaleHeight     =   150
      ScaleWidth      =   2160
      TabIndex        =   61
      Top             =   4200
      Width           =   2160
   End
   Begin VB.PictureBox Picture20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   2235
      Picture         =   "frmChat.frx":3E4F8
      ScaleHeight     =   150
      ScaleWidth      =   2160
      TabIndex        =   60
      Top             =   4200
      Width           =   2160
   End
   Begin VB.PictureBox Picture19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4380
      Picture         =   "frmChat.frx":3F61A
      ScaleHeight     =   180
      ScaleWidth      =   135
      TabIndex        =   59
      Top             =   4200
      Width           =   135
   End
   Begin VB.PictureBox Picture18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   4395
      Picture         =   "frmChat.frx":3F7AC
      ScaleHeight     =   1305
      ScaleWidth      =   120
      TabIndex        =   58
      Top             =   4350
      Width           =   120
   End
   Begin VB.PictureBox Picture17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   4395
      Picture         =   "frmChat.frx":4140E
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   57
      Top             =   5640
      Width           =   120
   End
   Begin VB.PictureBox Picture16 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   240
      Picture         =   "frmChat.frx":41558
      ScaleHeight     =   150
      ScaleWidth      =   4230
      TabIndex        =   56
      Top             =   5655
      Width           =   4230
   End
   Begin VB.TextBox txtMyArrival 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   435
      Width           =   5175
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   5385
      Picture         =   "frmChat.frx":438C2
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   55
      Top             =   3870
      Width           =   120
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   4725
      Picture         =   "frmChat.frx":43A0C
      ScaleHeight     =   150
      ScaleWidth      =   660
      TabIndex        =   54
      Top             =   3885
      Width           =   660
   End
   Begin VB.PictureBox Picture13 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   225
      Picture         =   "frmChat.frx":45D76
      ScaleHeight     =   150
      ScaleWidth      =   4500
      TabIndex        =   53
      Top             =   3885
      Width           =   4500
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3360
      Left            =   90
      Picture         =   "frmChat.frx":480E0
      ScaleHeight     =   3360
      ScaleWidth      =   120
      TabIndex        =   52
      Top             =   495
      Width           =   120
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      Picture         =   "frmChat.frx":49D42
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   51
      Top             =   3840
      Width           =   135
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   5385
      Picture         =   "frmChat.frx":49EF0
      ScaleHeight     =   3495
      ScaleWidth      =   120
      TabIndex        =   50
      Top             =   465
      Width           =   120
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   4665
      Picture         =   "frmChat.frx":4BB52
      ScaleHeight     =   150
      ScaleWidth      =   720
      TabIndex        =   48
      Top             =   315
      Width           =   720
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   2535
      Picture         =   "frmChat.frx":4CC74
      ScaleHeight     =   150
      ScaleWidth      =   2160
      TabIndex        =   47
      Top             =   315
      Width           =   2160
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   390
      Picture         =   "frmChat.frx":4DD96
      ScaleHeight     =   150
      ScaleWidth      =   2160
      TabIndex        =   46
      Top             =   315
      Width           =   2160
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      Picture         =   "frmChat.frx":4EEB8
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   45
      Top             =   315
      Width           =   315
   End
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   8790
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   5550
      Width           =   2160
   End
   Begin VB.TextBox txtInvite 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   2
      Left            =   8550
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   5205
      Width           =   2145
   End
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   8775
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4830
      Width           =   2160
   End
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   8805
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   4110
      Width           =   2160
   End
   Begin VB.TextBox txtInvite 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   1
      Left            =   8565
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   4485
      Width           =   2145
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   5
      Left            =   8655
      Picture         =   "frmChat.frx":4F23A
      ScaleHeight     =   375
      ScaleWidth      =   2400
      TabIndex        =   44
      Top             =   5475
      Width           =   2400
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   4
      Left            =   8415
      Picture         =   "frmChat.frx":5215C
      ScaleHeight     =   375
      ScaleWidth      =   2400
      TabIndex        =   43
      Top             =   5115
      Width           =   2400
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   8655
      Picture         =   "frmChat.frx":5507E
      ScaleHeight     =   375
      ScaleWidth      =   2400
      TabIndex        =   42
      Top             =   4755
      Width           =   2400
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   8655
      Picture         =   "frmChat.frx":57FA0
      ScaleHeight     =   375
      ScaleWidth      =   2400
      TabIndex        =   41
      Top             =   4035
      Width           =   2400
   End
   Begin VB.TextBox txtInvite 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   8550
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   3765
      Width           =   2145
   End
   Begin VB.PictureBox picTMR3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   11820
      Picture         =   "frmChat.frx":5AEC2
      ScaleHeight     =   540
      ScaleWidth      =   555
      TabIndex        =   33
      Top             =   5055
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picTMR2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   11820
      Picture         =   "frmChat.frx":5BEC4
      ScaleHeight     =   540
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   4335
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   8415
      Picture         =   "frmChat.frx":5CEC6
      ScaleHeight     =   375
      ScaleWidth      =   2400
      TabIndex        =   31
      Top             =   3675
      Width           =   2400
   End
   Begin VB.PictureBox picTMR1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   11820
      Picture         =   "frmChat.frx":5FDE8
      ScaleHeight     =   540
      ScaleWidth      =   555
      TabIndex        =   30
      Top             =   3570
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Timer tmrInvite3 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   12105
      Top             =   4920
   End
   Begin VB.Timer tmrInvite2 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   12090
      Top             =   4455
   End
   Begin VB.Timer tmrInvite1 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   12090
      Top             =   3975
   End
   Begin VB.CommandButton cmdDecline3 
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
      Caption         =   "X"
      Height          =   255
      Left            =   14115
      TabIndex        =   29
      Top             =   3270
      Width           =   255
   End
   Begin VB.CommandButton cmdAccept3 
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
      Caption         =   "/"
      Height          =   255
      Left            =   13755
      TabIndex        =   28
      Top             =   3270
      Width           =   255
   End
   Begin VB.CommandButton cmdDecline2 
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
      Caption         =   "X"
      Height          =   255
      Left            =   14115
      TabIndex        =   27
      Top             =   2790
      Width           =   255
   End
   Begin VB.CommandButton cmdAccept2 
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
      Caption         =   "/"
      Height          =   255
      Left            =   13755
      TabIndex        =   26
      Top             =   2790
      Width           =   255
   End
   Begin VB.CommandButton cmdAccept1 
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
      Caption         =   "/"
      Height          =   255
      Left            =   13755
      TabIndex        =   21
      Top             =   2310
      Width           =   255
   End
   Begin VB.CommandButton cmdDecline1 
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
      Caption         =   "X"
      Height          =   255
      Left            =   14115
      TabIndex        =   20
      Top             =   2310
      Width           =   255
   End
   Begin VB.Timer tmrReject 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8340
      Top             =   6270
   End
   Begin MSWinsockLib.Winsock socket1 
      Left            =   9825
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton cmdMyPrivateSend 
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
      Caption         =   "Private Message"
      Height          =   495
      Left            =   5565
      TabIndex        =   7
      Top             =   6045
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdInvite 
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
      Caption         =   "Invite to game"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13110
      TabIndex        =   6
      Top             =   3555
      Width           =   1455
   End
   Begin VB.Timer tmrCheckUsers 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   10305
      Top             =   6840
   End
   Begin VB.Timer tmrCheckOnLine 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   10665
      Top             =   6840
   End
   Begin MSWinsockLib.Winsock socketGetStat 
      Left            =   9105
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   5730
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdMySend 
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
      Caption         =   "&Send Message"
      Height          =   615
      Left            =   4635
      TabIndex        =   1
      Top             =   6060
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
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
      Caption         =   "&Exit"
      Height          =   375
      Left            =   10500
      TabIndex        =   0
      Top             =   6270
      Width           =   495
   End
   Begin MSWinsockLib.Winsock gamesocket 
      Left            =   9465
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   73
      LocalPort       =   73
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   8430
      Picture         =   "frmChat.frx":60DEA
      ScaleHeight     =   375
      ScaleWidth      =   2400
      TabIndex        =   34
      Top             =   4395
      Width           =   2400
   End
   Begin VB.ListBox lstPrivateUsers 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   5745
      TabIndex        =   2
      Top             =   4305
      Width           =   1920
   End
   Begin VB.Label lblStats 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0  'Transparent
      Caption         =   "User Biography"
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   7935
      TabIndex        =   15
      Top             =   90
      Width           =   1185
   End
   Begin VB.Label lblUserSelect 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0  'Transparent
      Caption         =   "<< select a user to view >>"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   9105
      TabIndex        =   14
      Top             =   90
      Width           =   3495
   End
   Begin VB.Label lblinvite3 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0  'Transparent
      Caption         =   "Invite 3:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7830
      TabIndex        =   101
      Top             =   5190
      Width           =   645
   End
   Begin VB.Label lblinvite2 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0  'Transparent
      Caption         =   "Invite 2:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7830
      TabIndex        =   100
      Top             =   4455
      Width           =   645
   End
   Begin VB.Label lblInvite1 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0  'Transparent
      Caption         =   "Invite 1:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7815
      TabIndex        =   99
      Top             =   3750
      Width           =   645
   End
   Begin VB.Label lblPrivate 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0  'Transparent
      Caption         =   "Private Users:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5700
      TabIndex        =   9
      Top             =   3990
      Width           =   1935
   End
   Begin VB.Label lblMessageHere 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   &H00886D1E&
      BackStyle       =   0  'Transparent
      Caption         =   "Type your message here:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   255
      TabIndex        =   12
      Top             =   4005
      Width           =   2535
   End
   Begin VB.Label lblInvite 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   &H8000000B&
      Height          =   255
      Index           =   0
      Left            =   12585
      TabIndex        =   17
      Top             =   375
      Width           =   2175
   End
   Begin VB.Label lblIP 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   &H8000000A&
      Height          =   255
      Index           =   0
      Left            =   12930
      TabIndex        =   22
      Top             =   705
      Width           =   1935
   End
   Begin VB.Label lblReject 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11250
      TabIndex        =   25
      Top             =   6420
      Width           =   1335
   End
   Begin VB.Label lblIP 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   &H8000000A&
      Height          =   255
      Index           =   2
      Left            =   12960
      TabIndex        =   24
      Top             =   1935
      Width           =   1935
   End
   Begin VB.Label lblIP 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   &H8000000A&
      Height          =   255
      Index           =   1
      Left            =   12945
      TabIndex        =   23
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblInvite 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   &H8000000A&
      Height          =   255
      Index           =   2
      Left            =   12555
      TabIndex        =   19
      Top             =   1620
      Width           =   2175
   End
   Begin VB.Label lblInvite 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   &H8000000A&
      Height          =   255
      Index           =   1
      Left            =   12555
      TabIndex        =   18
      Top             =   1035
      Width           =   2175
   End
   Begin VB.Label lblInvites 
      BackStyle       =   0  'Transparent
      Caption         =   "Invites:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7815
      TabIndex        =   16
      Top             =   3315
      Width           =   1935
   End
   Begin VB.Label lblUserMessages 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   &H00886D1E&
      BackStyle       =   0  'Transparent
      Caption         =   "Messages from users:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   210
      TabIndex        =   13
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblUser 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblOnline 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0  'Transparent
      Caption         =   "Online Users:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5700
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnugameboard 
         Caption         =   "Game Board"
      End
   End
   Begin VB.Shape shpA 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpAS1 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpAS2 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpB 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpBS1 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpBS2 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpC 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpCS1 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpCS2 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim MyUserID As String

Dim Users() As Record 'a list of the users in the user lsitbox on chat form
Dim PrivateUsers() As Record 'a list of private users in the private user list box

Dim MyStats As String 'mystats is the value of what is stored in the "biography" section

Private Type Record 'these store value of online users
    UserID As String 'indicates the name of the user which is displayed to the user
    UserIP As String 'stores the IP address of the user who is online
    OnLine As Boolean 'gives a value indicating if the user is online, true = online
End Type

Dim UserSelect As Integer 'the position in listbox of selected user
Dim UserNameSelect As String 'the Name of selected user
Dim UserIpSelect As String 'the IP of the selected user
Dim countSec As Integer 'the counter to flash rejected user

Dim InviteDis As Boolean

Private Sub cmdAccept1_Click()
If lblIP(0).Caption = "" Then
    Exit Sub
End If

'this accepts an invite, pushes other queued invites up list.
frmGame.lblColor.Caption = "white"

picTMR1.Visible = False

frmGame.txtIP = lblIP(0).Caption
frmGame.SocketGame.RemoteHost = lblIP(0).Caption
frmGame.SocketChat.RemoteHost = lblIP(0).Caption

gamesocket.SendData "accept*" & lblInvite(0).Caption & "*" & MyUserID & "*" & gamesocket.LocalIP

frmGame.Show
End Sub

Private Sub cmdAccept2_Click()
If lblIP(1).Caption = "" Then
    Exit Sub
End If

'this accepts an invite, pushes other queued invites up list.
frmGame.lblColor.Caption = "white"

picTMR2.Visible = False

frmGame.txtIP = lblIP(1).Caption
frmGame.SocketGame.RemoteHost = lblInvite(1).Caption
frmGame.SocketChat.RemoteHost = lblIP(1).Caption

gamesocket.SendData "accept*" & lblInvite(1).Caption & "*" & MyUserID & "*" & gamesocket.LocalIP

frmGame.Show
End Sub

Private Sub cmdaccept3_Click()
If lblIP(2).Caption = "" Then
    Exit Sub
End If

'this accepts an invite, pushes other queued invites up list.
frmGame.lblColor.Caption = "white"

picTMR3.Visible = False

frmGame.txtIP = lblIP(2).Caption
frmGame.SocketGame.RemoteHost = lblIP(2).Caption
frmGame.SocketChat.RemoteHost = lblIP(2).Caption

gamesocket.SendData "accept*" & lblInvite(2).Caption & "*" & MyUserID & "*" & gamesocket.LocalIP

frmGame.Show
End Sub

Private Sub cmdDecline1_Click()
If lblIP(0).Caption = "" Then
    Exit Sub
End If

'this declines an invite, pushes other queued invites up list.
gamesocket.SendData "reject" & lblInvite(0).Caption & MyUserID

picTMR1.Visible = False

lblInvite(0).Caption = ""
lblIP(0).Caption = ""

'moves second and third place invites up one place each
lblInvite(0).Caption = lblInvite(1).Caption
lblIP(0).Caption = lblInvite(1).Caption
lblInvite(1).Caption = ""
lblInvite(1).Caption = ""

lblInvite(1).Caption = lblInvite(2).Caption
lblIP(1).Caption = lblInvite(2).Caption
lblInvite(2).Caption = ""
lblInvite(2).Caption = ""

End Sub

Private Sub cmdDecline2_Click()
If lblIP(1).Caption = "" Then
    Exit Sub
End If

'this declines an invite, pushes other queued invites up list.
gamesocket.SendData "reject" & lblInvite(1).Caption & MyUserID

picTMR2.Visible = False

lblInvite(1).Caption = ""
lblIP(1).Caption = ""

'moves second and third place invites up one place each
lblInvite(1).Caption = lblInvite(2).Caption
lblIP(1).Caption = lblInvite(2).Caption
lblInvite(2).Caption = ""
lblInvite(2).Caption = ""

End Sub

Private Sub cmdDecline3_Click()
If lblIP(2).Caption = "" Then
    Exit Sub
End If

'this declines an invite, pushes other queued invites up list.
gamesocket.SendData "reject" & lblInvite(2).Caption & MyUserID

picTMR3.Visible = False

lblInvite(2).Caption = ""
lblIP(2).Caption = ""

'deletes 3rd invite from que
lblInvite(2).Caption = ""
lblIP(2).Caption = ""

End Sub

Private Sub cmdExit_Click()
'this exits the form and informs online users that the player has left the chatroom
'so that the systm may remove them from record of users
Dim Message As String
Message = MyUserID & "Has Left the Room"

'sends the message myuserid"has left the room" to all users.
For i = 0 To UBound(Users)
    socket1.RemoteHost = Users(i).UserIP
    socket1.SendData Message
Next i

txtMySend = ""

End
End Sub

Private Sub cmdInvite_Click()
'sends an invite to the user that is selected.
gamesocket.SendData "Invite" & UserNameSelect & "*" & MyUserID & "*" & gamesocket.LocalIP
End Sub

Private Sub cmdMyPrivateSend_Click()
'sends message to all private users

If txtMySend.Text = "" Then
    Exit Sub
End If

'sends in form "[Mckic1909]Hello" where McKic1909 is a username,
'and "Hello" is a message
Dim Message As String
Message = "[" & MyUserID & "] " & txtMySend.Text

'sends to all private users
For i = 0 To UBound(PrivateUsers)
    socket1.RemoteHost = PrivateUsers(i).UserIP
    socket1.SendData Message
Next i

txtMySend = ""
txtMySend.SetFocus
End Sub

Private Sub cmdMySend_Click()
'sends a message to all online users.

If txtMySend.Text = "" Then
    Exit Sub
End If

'sends in form "[Mckic1909]Hello" where McKic1909 is a username,
'and "Hello" is a message

Dim Message As String
Message = "[" & MyUserID & "] " & txtMySend.Text

'sends to all online users
For i = 0 To UBound(Users)
    socket1.RemoteHost = Users(i).UserIP
    socket1.SendData Message
Next i

txtMySend = ""
txtMySend.SetFocus

End Sub
Private Sub Form_Activate()
lstUsers.ListIndex = 0
UserSelect = 0 'the username that is selected, top of list
InviteDis = False
End Sub



Private Sub ApplyLocale()
If LocLanguage = "en" Then Exit Sub
cmdMyPrivateSend.Caption = Loc("chat.privateSend")
cmdInvite.Caption = Loc("chat.invite")
cmdMySend.Caption = Loc("chat.send")
cmdExit.Caption = Loc("chat.exit")
lblStats.Caption = Loc("chat.bio")
lblUserSelect.Caption = Loc("chat.select")
lblInvite1.Caption = Loc("chat.invite1")
lblinvite2.Caption = Loc("chat.invite2")
lblinvite3.Caption = Loc("chat.invite3")
lblPrivate.Caption = Loc("chat.private")
lblMessageHere.Caption = Loc("chat.typeHere")
lblInvites.Caption = Loc("chat.invites")
lblUserMessages.Caption = Loc("chat.messages")
lblOnline.Caption = Loc("chat.online")
mnugameboard.Caption = Loc("chat.gameBoard")
End Sub

' A clean full-screen layout: messages, the people, and the invitations.
' The art around the boxes and the pictured buttons are hidden; the real
' buttons stand in for them.
Private Sub ChatLayout()
Dim c As Control
Dim n As String
Dim i As Long
Dim y As Single
For Each c In Controls
  n = c.Name
  If Left(n, 7) = "Picture" Or Left(n, 4) = "picC" Or Left(n, 4) = "picS" Or Left(n, 4) = "picI" Or Left(n, 4) = "picT" Or n = "txtIP" Or n = "lblIP" Or n = "lblInvite" Or n = "lblUser" Then
    c.Visible = False
  End If
Next c
ThemeFullScreen Me
BackColor = ThBack
ThemeCard shpA, shpAS1, shpAS2, 60, 60, 5100, 7080, 15
ThemeCard shpB, shpBS1, shpBS2, 5300, 60, 2300, 7080, 15
ThemeCard shpC, shpCS1, shpCS2, 7740, 60, 5200, 7080, 15
lblUserMessages.Move 160, 130, 4900, 300
txtMyArrival.Move 160, 470, 4900, 3600
lblMessageHere.Move 160, 4200, 4900, 300
txtMySend.Move 160, 4540, 4900, 1500
cmdMySend.Move 160, 6200, 2400, 650
cmdMyPrivateSend.Move 2660, 6200, 2400, 650
cmdMySend.Visible = True
cmdMyPrivateSend.Visible = True
cmdMyPrivateSend.Caption = Loc("chat.sendPrivate")
cmdMySend.BackColor = ThAccent2
lblOnline.Move 5400, 130, 2100, 300
lstUsers.Move 5400, 470, 2100, 3000
lblPrivate.Move 5400, 3600, 2100, 300
lstPrivateUsers.Move 5400, 3940, 2100, 2900
lblStats.Move 7840, 130, 2400, 300
lblUserSelect.Move 7840, 450, 4900, 300
txtstat.Move 7840, 800, 4900, 1900
lblInvites.Move 7840, 2850, 4900, 300
y = 3250
For i = 0 To 2
  txtInvite(i).Move 8800, y, 2500, 420
  y = y + 600
Next i
lblInvite1.Move 7840, 3300, 900, 300
lblinvite2.Move 7840, 3900, 900, 300
lblinvite3.Move 7840, 4500, 900, 300
cmdAccept1.Move 11400, 3250, 600, 420
cmdDecline1.Move 12100, 3250, 600, 420
cmdAccept2.Move 11400, 3850, 600, 420
cmdDecline2.Move 12100, 3850, 600, 420
cmdAccept3.Move 11400, 4450, 600, 420
cmdDecline3.Move 12100, 4450, 600, 420
cmdAccept1.Caption = ChrW(10003)
cmdAccept2.Caption = ChrW(10003)
cmdAccept3.Caption = ChrW(10003)
cmdDecline1.Caption = ChrW(10005)
cmdDecline2.Caption = ChrW(10005)
cmdDecline3.Caption = ChrW(10005)
cmdInvite.Move 7840, 5350, 4900, 650
cmdInvite.Visible = True
cmdExit.Move 7840, 6200, 2400, 650
cmdExit.BackColor = ThBad
lblReject.Move 10400, 6250, 2400, 550
lblStats.BackColor = ThPanel
lblUserSelect.BackColor = ThPanel
lblInvites.BackColor = ThPanel
lblInvite1.BackColor = ThPanel
lblinvite2.BackColor = ThPanel
lblinvite3.BackColor = ThPanel
lblOnline.BackColor = ThPanel
lblPrivate.BackColor = ThPanel
lblMessageHere.BackColor = ThPanel
lblUserMessages.BackColor = ThPanel
lblReject.BackColor = ThPanel
lblStats.ForeColor = ThAccent
lblOnline.ForeColor = ThAccent
lblPrivate.ForeColor = ThAccent
lblInvites.ForeColor = ThAccent
lblUserMessages.ForeColor = ThAccent
lblMessageHere.ForeColor = ThAccent
ThemeShift Me, (ScaleWidth - 13000) / 2, (ScaleHeight - 7200) / 2
ThemeFitAll Me
End Sub

Private Sub Form_Load()
ApplyLocale

'sets ports
socket1.RemoteHost = "255.255.255.255"
socket1.RemotePort = 72
socket1.LocalPort = 72
frmChat.BackColor = ThBack
ThemeApply Me
mnuView.Visible = False
ChatLayout
' the other states of the Invite button are only picture sources
picInvite.Visible = False
picInviteDis.Visible = False
picInviteOver.Visible = False
Dim pc As Control
For Each pc In Controls
    If TypeName(pc) = "PictureBox" Then
        If pc.BackColor = 8940830 Then pc.BackColor = ThBack
    End If
Next pc

gamesocket.SendData ""


'sets user name to blank
UserNameSelect = ""

lblMessageHere.BackColor = ThPanel
lblUserMessages.BackColor = ThPanel

gamesocket.RemoteHost = "255.255.255.255"

'loads stats
MyStats = "" 'frmStats.txtstats.Text



'used for listbox of users
Dim sBuffer As String
Dim lSize As Long

sBuffer = Space$(255)
lSize = Len(sBuffer)
'sub routine that gets a username of local user
Call GetUserName(sBuffer, lSize)
If lSize > 0 Then
    MyUserID = Mid(Left$(sBuffer, lSize - 1), 1, 9)
    'to make all user ID of equal length.
    If Len(MyUserID) < 9 Then
        Do Until Len(MyUserID) = 9
            MyUserID = MyUserID & "_"
        Loop
    End If
Else
    MyUserID = vbNullString
End If

'Was used for implementation testing purposes (so i could be loged onto two computers)
'If socketGetStat.LocalIP <> "172.17.0.144" Then
'    MyUserID = "TESTINGXX"
'End If

frmChat.Caption = LocF("chat.title", MyUserID, socket1.LocalIP)
lblUser.Caption = MyUserID

socket1.SendData "#" & MyUserID & vbTab & socket1.LocalIP


'Sets user XP etc. #############
Open App.Path & "\User\User.txt" For Input As #18
    Input #18, C
        userlevel = C
    Input #18, C
        games = C
    Input #18, C
        wins = C
    Input #18, C
        losses = C
Close #18

userXP = games + wins

frmStats.txtstats = MyUserID & vbNewLine & _
LocF("stats.c", userXP, wins, losses, games)

MyStats = frmStats.txtstats.Text

'################################


ReDim Users(0)
Users(0).UserID = MyUserID
Users(0).UserIP = socket1.LocalIP

ReDim PrivateUsers(0)
PrivateUsers(0).UserID = MyUserID
PrivateUsers(0).UserIP = socket1.LocalIP
tmrCheckUsers.Enabled = True
tmrCheckOnLine.Enabled = True


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picSendMess.Picture = picCMDOUT.Picture
picSendPriv.Picture = picCMDOUT2.Picture

If cmdInvite.Enabled = True Then
    picInviteAct.Picture = picInvite.Picture
Else
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'before form is closed, the strings used to inform other users that "i" have
'left the room is send.

For i = 1 To UBound(Users)
    socket1.RemoteHost = Users(i).UserIP
    socket1.SendData "~" & MyUserID
Next i

'unloads other forms
Unload frmPawnPromotion
Unload frmStats
frmGame.Show
End Sub

Private Sub gamesocket_DataArrival(ByVal bytesTotal As Long)
Dim Arrival As String
gamesocket.GetData Arrival

If Arrival = "" Then
    Exit Sub
End If

'if it is a reject then display this to the user
If Mid(Arrival, 1, 15) = "reject" & MyUserID Then
    lblReject.Caption = LocF("chat.rejected", Mid(Arrival, 16, Len(Arrival)))
    tmrReject.Enabled = True
    Exit Sub
End If

'when invites are recieved
'if it is an invite, display this to the user
'in invite lables
If Mid(Arrival, 1, 15) = "Invite" & MyUserID Then
    
    For i = 0 To 2
        If lblInvite(i).Caption = Mid(Arrival, 7) Then
        End If
    Next i
    
    invite = Split(Mid(Arrival, 17, Len(Arrival)), "*")
    If lblInvite(0) = "" Then 'for first position when all are free
       
        lblInvite(0) = invite(0)
        lblIP(0) = invite(1)
        
        Call InviteToGame
        Exit Sub
    Else
        If lblInvite(0) = "" Then 'for second position when two are free
        
            lblInvite(1) = invite(0)
            lblIP(1) = invite(1)
            
        Call InviteToGame
        Exit Sub
            Else
                If lblInvite(0) = "" Then 'for last(third) position when 1 place is free
                
                    lblInvite(2) = invite(0)
                    lblIP(2) = invite(1)
                
                    Call InviteToGame
                    Exit Sub
                End If
        End If
    End If
End If

'the oponent has accepted an invite.
'sets ports for oponent, sets local piece color to "black"
If Mid(Arrival, 1, 16) = "accept*" & MyUserID Then
    Accept = Split(Arrival, "*")
    frmGame.lblColor.Caption = "black"
    
    frmGame.SocketGame.RemoteHost = Accept(3)
    frmGame.SocketChat.RemoteHost = Accept(3)

    frmGame.lblColor = "black"
    frmGame.Show
    Unload Me
End If


If Arrival = MyUserID Then
    gamesocket.SendData "*" & MyUserID & frmStats.txtstats.Text
End If

'sets the "biography" sections to the recieved data.
If Mid(Arrival, 1, 1) = "*" Then
    Arrival = Mid(Arrival, 2, Len(Arrival))
    If Mid(Arrival, 1, 9) = UserNameSelect Then
        txtstat = Mid(Arrival, 10, Len(Arrival))
        Exit Sub
    End If
End If

End Sub

Private Sub lblInvite_Change(Index As Integer)
'when a invite is recieved then a timer is started, after 15 secs. the
'invite is rejected.

'change textbox text to same as label
For i = 0 To 2
    txtInvite(i).Text = lblInvite(i)
Next i

If Index = 0 And lblInvite(0).Caption <> "" Then
    tmrInvite1.Enabled = True
    picTMR1.Visible = True
End If
If Index = 1 And lblInvite(1).Caption <> "" Then
    tmrInvite2.Enabled = True
    picTMR2.Visible = True
End If
If Index = 2 And lblInvite(2).Caption <> "" Then
    tmrInvite3.Enabled = True
    picTMR2.Visible = True
End If


End Sub

Private Sub lblIP_Change(Index As Integer)
For i = 0 To 2
    txtIP(i).Text = lblIP(i).Caption
Next i
End Sub

Private Sub lstPrivateUsers_DblClick()
'removes a user to private user list box
RemoveFromPrivateArray lstPrivateUsers
Call UpdateListBox
End Sub



Private Sub lstUsers_DblClick()
'adds a user to private user list box
If lstUsers.Text = lblUser.Caption Then
    Exit Sub
End If

ReDim Preserve PrivateUsers(UBound(PrivateUsers) + 1)
PrivateUsers(UBound(PrivateUsers)).UserID = lstUsers.Text
PrivateUsers(UBound(PrivateUsers)).UserIP = GetUserIP(lstUsers.Text)
Call UpdateListBox
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'adds a user to the invite button and variables.
UserNameSelect = lstUsers.Text
UserIpSelect = GetUserIP(lstUsers.Text)
lblUserSelect = UserNameSelect


gamesocket.SendData ""

If UserNameSelect = MyUserID Then
    txtstat = frmStats.txtstats.Text
    cmdInvite.Enabled = False
    picInviteAct.Picture = picInviteDis.Picture
    cmdInvite.Caption = Loc("chat.invite")
    Exit Sub
Else
    gamesocket.SendData UserNameSelect
    cmdInvite.Caption = LocF("chat.inviteUser", UserNameSelect)
    cmdInvite.Enabled = True
    picInviteAct.Picture = picInvite.Picture
End If


End Sub

Private Sub mnugameboard_Click()
'shows the game form
frmGame.Show
End Sub

Private Sub picCross1_Click()
Call cmdDecline1_Click
End Sub

Private Sub picCross2_Click()
cmdDecline2_Click
End Sub

Private Sub picInviteAct_Click()

If cmdInvite.Enabled = False Then
Else
    Call cmdInvite_Click
    Exit Sub
End If


End Sub

Private Sub picInviteAct_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdInvite.Enabled = True Then
    picInviteAct.Picture = picInviteOver.Picture
Else
End If

End Sub
Private Sub picSendMess_Click()
Call cmdMySend_Click
End Sub

Private Sub picSendMess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picSendMess.Picture = picCMDOVER.Picture
End Sub

Private Sub picSendPriv_Click()
Call cmdMyPrivateSend_Click
End Sub

Private Sub picSendPriv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picSendPriv.Picture = picCMDOVER2.Picture
End Sub

Private Sub picTick1_Click()
Call cmdAccept1_Click
End Sub

Private Sub picTick2_Click()
Call cmdAccept2_Click
End Sub

Private Sub Picture53_Click()
Call cmdaccept3_Click
End Sub

Private Sub Picture54_Click()
Call cmdDecline3_Click
End Sub

Private Sub Socket1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errorhandler

Dim TmpUsers As Record


Dim Arrival As String
socket1.GetData Arrival

If Arrival = "PING" Then
    txtMyArrival = txtMyArrival
    Else


    'checks to see if users are online, this is done periodicaly
    'on a timer event called "tmrusercheck"
    If Left(Arrival, 1) = "!" Then  '
        For i = 1 To UBound(Users)
            If Users(i).UserID = Right(Arrival, Len(Arrival) - 1) Then
               Users(i).OnLine = True
                Exit For
            End If
        Next i
    End If



'is sent my new user on a form load event - when program is ran
    If Left(Arrival, 1) = "#" Then
        TabStop = InStr(Arrival, vbTab)
        
        TmpUsers.UserID = Mid(Arrival, 2, TabStop - 2)
        TmpUsers.UserIP = Mid(Arrival, TabStop + 1)
    
        socket1.RemoteHost = TmpUsers.UserIP
'data is sent when a new user runs program
        socket1.SendData "-" & MyUserID & vbTab & socket1.LocalIP

        If InUserArray(TmpUsers.UserID) = False Then
            ReDim Preserve Users(UBound(Users) + 1)
            Users(UBound(Users)).UserID = TmpUsers.UserID
            Users(UBound(Users)).UserIP = TmpUsers.UserIP
            Users(UBound(Users)).OnLine = True
        End If
    
        Call UpdateListBox
    End If
    
'may be used to stop adding myself to the listbox
'see above comment to see when data is sent
        If Left(Arrival, 1) = "-" Then
        TabStop = InStr(Arrival, vbTab)
    
        TmpUsers.UserID = Mid(Arrival, 2, TabStop - 2)
        TmpUsers.UserIP = Mid(Arrival, TabStop + 1)
    
        If InUserArray(TmpUsers.UserID) = False Then
            ReDim Preserve Users(UBound(Users) + 1)
            Users(UBound(Users)).UserID = TmpUsers.UserID
            Users(UBound(Users)).UserIP = TmpUsers.UserIP
            Users(UBound(Users)).OnLine = True
        End If
    
        Call UpdateListBox
    End If
    
    
    
    If Left(Arrival, 1) = "~" Then          'the user has exited the program
        ' Remove record from the array
        Call RemoveFromArray(Mid(Arrival, 2))
        Call UpdateListBox
    End If
    If Left(Arrival, 1) = "[" Then
        txtMyArrival = txtMyArrival & Arrival & vbNewLine
        txtMyArrival.SelStart = Len(txtMyArrival)
    End If
End If


'lstUsers.ListIndex = 0 ' previously selected user in userlist will be reselected

'below selects the user that was previously selected >
'this could change when a user leaves or is added.
'For i = 0 To lstUsers.ListCount - 1
'    If lstUsers.List(i) = UserNameSelect Then
'        lstUsers.ListIndex = i
'        Exit For
'    End If
'Next i

lstUsers.ListIndex = 0


Exit Sub

errorhandler:
If Err.Number = "10054" Then
End If

MsgBox (LocF("chat.error", Err.Number, Err.Description))

End Sub
Public Sub RemoveFromArray(ByVal ThisUserID As String)
' Remove From First Array
For i = 0 To UBound(Users)
    If Users(i).UserID = ThisUserID Then j = i: Exit For
Next i

For i = j To UBound(Users) - 1
    Users(i) = Users(i + 1)
Next i

ReDim Preserve Users(UBound(Users) - 1)


j = 0

For i = 0 To UBound(PrivateUsers)
    If PrivateUsers(i).UserID = ThisUserID Then j = i: Exit For
Next i

If Not j = 0 Then
    For i = j To UBound(PrivateUsers) - 1
        PrivateUsers(i) = PrivateUsers(i + 1)
    Next i
    
    ReDim Preserve PrivateUsers(UBound(PrivateUsers) - 1)
End If
End Sub

Public Sub UpdateListBox()
lstUsers.Clear

For i = 0 To UBound(Users)
    lstUsers.AddItem Users(i).UserID
Next i


lstPrivateUsers.Clear

For i = 0 To UBound(PrivateUsers)
    lstPrivateUsers.AddItem PrivateUsers(i).UserID
Next i

End Sub
Public Function InUserArray(ByVal ID As String) As Boolean
InUserArray = False
For i = 0 To UBound(Users)
    If Users(i).UserID = ID Then
        InUserArray = True
        Exit For
    End If
Next i
End Function

Public Function GetUserIP(ByVal ID As String) As String
'gets the IP address of a specific user and returns the value
For i = 0 To UBound(Users)
    If Users(i).UserID = ID Then
        GetUserIP = Users(i).UserIP
        Exit For
    End If
Next i
End Function

Public Sub RemoveFromPrivateArray(ByVal ThisUserID As String)
'remove from array and from list box
For i = 0 To UBound(PrivateUsers)
    If PrivateUsers(i).UserID = ThisUserID Then j = i: Exit For
Next i

If Not j = 0 Then
    For i = j To UBound(PrivateUsers) - 1
        PrivateUsers(i) = PrivateUsers(i + 1)
    Next i
    
    ReDim Preserve PrivateUsers(UBound(PrivateUsers) - 1)
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub tmrCheckOnLine_Timer()
'check users online status
For i = 1 To UBound(Users)
    If Users(i).OnLine = False Then
        'remove from arrays and lists
        Call RemoveFromArray(Users(i).UserID)
    End If
    Users(i).OnLine = False
Next i
UpdateListBox



End Sub

Private Sub tmrCheckUsers_Timer()
Message = "!" & MyUserID
For i = 0 To UBound(Users)
    socket1.RemoteHost = Users(i).UserIP
    socket1.SendData Message
Next i
End Sub
Public Sub InviteToGame()
'when an invite is recieved a messagebox is displayed.
MsgBox (Loc("chat.inviteReceived"))
End Sub

Private Sub tmrInvite1_Timer()
'The following 3 routines reject invites after 15 seconds
Call cmdDecline1_Click
tmrInvite1.Enabled = False
End Sub

Private Sub tmrInvite2_Timer()
Call cmdDecline2_Click
tmrInvite2.Enabled = False
End Sub

Private Sub tmrInvite3_Timer()
Call cmdDecline3_Click
tmrInvite3.Enabled = False
End Sub

Private Sub tmrReject_Timer()

countSec = countSec + 1

If countSec = 5 Then
    countSec = 0
    lblReject.Visible = False
    tmrReject.Enabled = False
End If

If countSec Mod 2 = 0 Then
    lblReject.Visible = False
    tmrReject.Enabled = False
End If
End Sub

Private Sub txtMySend_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 13) Then
    cmdMySend_Click
    txtMySend.Text = ""
End If
End Sub
