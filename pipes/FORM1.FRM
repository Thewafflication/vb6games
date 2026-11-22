VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pipes version 1.0"
   ClientHeight    =   5055
   ClientLeft      =   1545
   ClientTop       =   2085
   ClientWidth     =   6525
   DrawWidth       =   3
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   4560
   End
   Begin VB.Label lblTitle 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PIPES"
      Height          =   300
      Left            =   750
      TabIndex        =   9
      Top             =   200
      Width           =   3000
   End
   Begin VB.CommandButton btnQuit 
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
      Appearance      =   0  'Flat
      Caption         =   "&Quit"
      Height          =   372
      Left            =   960
      TabIndex        =   0
      Top             =   3720
      Width           =   972
   End
   Begin VB.Image out 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Stretch         =   -1  'True
      Left            =   240
      Picture         =   "Form1.frx":030A
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image out 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Stretch         =   -1  'True
      Left            =   0
      Picture         =   "Form1.frx":0614
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image inx 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Stretch         =   -1  'True
      Left            =   3000
      Picture         =   "Form1.frx":091E
      Top             =   600
      Width           =   480
   End
   Begin VB.Image inx 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Stretch         =   -1  'True
      Left            =   3360
      Picture         =   "Form1.frx":0C28
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   34
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":0F32
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape boxcurrent 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   408
      Left            =   5640
      Top             =   240
      Width           =   408
   End
   Begin VB.Image current 
      Appearance      =   0  'Flat
      Height          =   480
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":123C
      Top             =   240
      Width           =   480
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   27
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":1546
      Top             =   4440
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   26
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":1850
      Top             =   4440
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   25
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":1B5A
      Top             =   4440
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   24
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":1E64
      Top             =   4440
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   23
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":216E
      Top             =   3840
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   22
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":2478
      Top             =   3840
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   21
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":2782
      Top             =   3840
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   20
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":2A8C
      Top             =   3840
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   19
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":2D96
      Top             =   3240
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   18
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":30A0
      Top             =   3240
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   17
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":33AA
      Top             =   3240
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   16
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":36B4
      Top             =   3240
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   15
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":39BE
      Top             =   2640
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   14
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":3CC8
      Top             =   2640
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   13
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":3FD2
      Top             =   2640
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   12
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":42DC
      Top             =   2640
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   11
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":45E6
      Top             =   2040
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   10
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":48F0
      Top             =   2040
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   9
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":4BFA
      Top             =   2040
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   8
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":4F04
      Top             =   2040
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   7
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":520E
      Top             =   1440
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   6
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":5518
      Top             =   1440
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   5
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":5822
      Top             =   1440
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   4
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":5B2C
      Top             =   1440
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   3
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":5E36
      Top             =   840
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   2
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":6140
      Top             =   840
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   1
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":644A
      Top             =   840
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image piece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Index           =   0
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":6754
      Top             =   840
      Width           =   510
   End
   Begin VB.Image blank 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Stretch         =   -1  'True
      Left            =   5640
      Picture         =   "Form1.frx":6A5E
      Top             =   240
      Width           =   510
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   13
      X1              =   20
      X2              =   212
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   12
      X1              =   20
      X2              =   212
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   11
      X1              =   20
      X2              =   212
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   10
      X1              =   20
      X2              =   212
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   9
      X1              =   20
      X2              =   212
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   8
      X1              =   20
      X2              =   212
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   7
      X1              =   50
      X2              =   242
      Y1              =   50
      Y2              =   50
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   6
      X1              =   20
      X2              =   20
      Y1              =   170
      Y2              =   362
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   5
      X1              =   20
      X2              =   20
      Y1              =   170
      Y2              =   362
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   4
      X1              =   20
      X2              =   20
      Y1              =   170
      Y2              =   362
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   3
      X1              =   20
      X2              =   20
      Y1              =   170
      Y2              =   362
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   20
      X2              =   20
      Y1              =   170
      Y2              =   362
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   20
      X2              =   20
      Y1              =   170
      Y2              =   362
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   40
      X2              =   40
      Y1              =   50
      Y2              =   242
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   35
      Stretch         =   -1  'True
      Left            =   600
      Picture         =   "Form1.frx":6D68
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   33
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":7072
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   32
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":737C
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   31
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":7686
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   30
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":7990
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   29
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":7C9A
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   28
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":7FA4
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   27
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":82AE
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   26
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":85B8
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   25
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":88C2
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   24
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":8BCC
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   23
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":8ED6
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   22
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":91E0
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   21
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":94EA
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   20
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":97F4
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   19
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":9AFE
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   18
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":9E08
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   17
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":A112
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   16
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":A41C
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   15
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":A726
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   14
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":AA30
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   13
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":AD3A
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   12
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":B044
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   11
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":B34E
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   10
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":B658
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   9
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":B962
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   8
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":BC6C
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   7
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":BF76
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   6
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":C280
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":C58A
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":C894
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":CB9E
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":CEA8
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":D1B2
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Stretch         =   -1  'True
      Left            =   720
      Picture         =   "Form1.frx":D4BC
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Water 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   0
      Left            =   600
      Top             =   2856
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   46
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   45
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   44
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   43
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   42
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   41
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   40
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   39
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   38
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   37
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   36
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   35
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   34
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   33
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   32
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   31
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   30
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   29
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   28
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   27
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   26
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   25
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   24
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   23
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   22
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   21
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   20
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   19
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   18
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   17
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   16
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   15
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   14
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   13
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   12
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   11
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   10
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   9
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   8
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   7
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   6
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   5
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   4
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   3
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   2
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   1
      Left            =   360
      Top             =   4200
      Width           =   2304
   End
   Begin VB.Shape Water 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   47
      Left            =   360
      Top             =   4200
      Width           =   2304
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
   Begin VB.Shape shpPal 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpPalS1 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpPalS2 
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
  Dim tb As Single, lb As Single   'board top left coords
  Dim dragging As Boolean  'a piece is being dragged from the palette
  Dim wasSel As Boolean    'the pressed piece was already the selected one
  Dim downX As Single, downY As Single  'where the press was on the piece
  Dim yy As Integer
  Dim irow, orow As Integer 'in and out rows
  Dim si As Single        'icon size
  Dim sb As Single        'size of board
  Dim wt As Integer       'top layer of water
  Dim np As Integer       'no of selected piece
  Dim nb As Integer       'no of board square
  Dim nxt(27) As Integer  'next icon in rotation
  Dim brdarray(35) As brdinfo ' info on each board square
  Dim piecearray(27) As pieceinfo 'info on each piece
  Dim pipearray(35) As Integer 'is there a pipe there?

Private Const PageY As Long = &H4AE2FF      ' sunny yellow
Private Const PipeDark As Long = &H69450E    ' deep blue
Private Const WaterC As Long = &HFFE500      ' cyan water
Private Const CellC As Long = &HFDF8E9
Private Const CellLine As Long = &HE8D296
Private Const EdgeC As Long = &HB28E2F

Private Sub blank_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
   np = -1
   current.Picture = blank.Picture
   boxcurrent.top = blank.top
   boxcurrent.left = blank.left
End Sub

Private Sub blank_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
 ' movepiece blank.Left, blank.Top
  current.left = blank.left
  current.top = blank.top
End Sub

Private Sub Board_MouseDown(index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i, a, t, b, l, r As Integer
  dragging = False
  a = 0
  If button = 1 And brdarray(index).available Then
    If np = -1 Then
      Board(index).Picture = blank.Picture
    Else
     If brdarray(index).side(1).pipe <> 0 Then
      If brdarray(index).side(1).pipe = piecearray(np).top Then
       a = a + 1
      End If
     Else
      a = a + 1
     End If
     If brdarray(index).side(2).pipe <> 0 Then
      If brdarray(index).side(2).pipe = piecearray(np).right Then
       a = a + 1
      End If
     Else
      a = a + 1
     End If
     If brdarray(index).side(3).pipe <> 0 Then
      If brdarray(index).side(3).pipe = piecearray(np).bottom Then
       a = a + 1
      End If
     Else
      a = a + 1
     End If
     If brdarray(index).side(4).pipe <> 0 Then
      If brdarray(index).side(4).pipe = piecearray(np).left Then
       a = a + 1
      End If
     Else
      a = a + 1
     End If
     If a = 4 Then
      Board(index).Picture = piece(np).Picture
      brdarray(index).pieceno = np
      brdarray(index).available = False
      t = piecearray(np).top
      r = piecearray(np).right
      l = piecearray(np).left
      b = piecearray(np).bottom
      If index - 6 >= 0 Then
      brdarray(index - 6).side(3).pipe = t
      End If
      brdarray(index).side(1).empty = t
      If index + 1 <= 35 Then
      brdarray(index + 1).side(4).pipe = r
      End If
      brdarray(index).side(2).empty = r
      If index + 6 <= 35 Then
      brdarray(index + 6).side(1).pipe = b
      End If
      brdarray(index).side(3).empty = b
      If index - 1 >= 0 Then
      brdarray(index - 1).side(2).pipe = l
      End If
      brdarray(index).side(4).empty = l

      ' handle top

      If index - 6 >= 0 Then
       If Not brdarray(index).side(1).empty Then
        i = index - 6
        If brdarray(i).pieceno = -1 Then
          brdarray(i).available = True
        End If
       End If
      End If

      ' handle right

      If index + 1 <= 35 Then
       If Not brdarray(index).side(2).empty Then
        i = index + 1
        If (brdarray(i).pieceno = -1) And Not ((index + 1) Mod 6 = 0) Then
          brdarray(i).available = True
        End If
       End If
      End If

      ' handle bottom

      If index + 6 <= 35 Then
       If Not brdarray(index).side(3).empty Then
        i = index + 6
        If brdarray(i).pieceno = -1 Then
          brdarray(i).available = True
        End If
       End If
      End If

      ' handle left

      If index - 1 >= 0 Then
       If Not brdarray(index).side(4).empty Then
        i = index - 1
        If (brdarray(i).pieceno = -1) And Not (index Mod 6 = 0) Then
          brdarray(i).available = True
        End If
       End If
      End If
      ' is it a winner?
      xx = False
      For i = 0 To 35            'is there a loose end?
      If brdarray(i).available = True Then
      xx = True
      End If
      Next i
      If orow * 6 = index Then
      yy = True 'we have just connected to out so record it incase there are still loose ends to connect
      End If
      If yy = True And xx = False Then 'is source connected to drain and no loose ends?
      Timer1.Enabled = False
      MsgBox Loc("msg.win")
      NewGame
      Exit Sub
      End If
     End If
    End If
  Else
    rotatepiece
  End If
End Sub

' True when the selected piece fits the board square: its open sides meet the
' pipes already there.
Private Function Fits(ByVal index As Integer) As Boolean
  Dim a As Integer
  Fits = False
  If np < 0 Then Exit Function
  If Not brdarray(index).available Then Exit Function
  a = 0
  If brdarray(index).side(1).pipe <> 0 Then
    If brdarray(index).side(1).pipe = piecearray(np).top Then a = a + 1
  Else
    a = a + 1
  End If
  If brdarray(index).side(2).pipe <> 0 Then
    If brdarray(index).side(2).pipe = piecearray(np).right Then a = a + 1
  Else
    a = a + 1
  End If
  If brdarray(index).side(3).pipe <> 0 Then
    If brdarray(index).side(3).pipe = piecearray(np).bottom Then a = a + 1
  Else
    a = a + 1
  End If
  If brdarray(index).side(4).pipe <> 0 Then
    If brdarray(index).side(4).pipe = piecearray(np).left Then a = a + 1
  Else
    a = a + 1
  End If
  Fits = (a = 4)
End Function

Private Sub Board_MouseMove(index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
  nb = index
  current.Visible = Fits(index)
  current.left = Board(index).left
  current.top = Board(index).top
End Sub

Private Sub btnQuit_Click()
  End
End Sub

Private Sub current_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
  If button = 2 Then
    rotatepiece
  ElseIf button = 1 And nb >= 0 Then
    Board_MouseDown nb, button, Shift, x, y
  End If

End Sub

Private Sub ApplyLocale()
  If LocLanguage = "en" Then Exit Sub
  Caption = Loc("app.title")
End Sub

' The pictures come in sets of several sizes: the one for this screen.
Private Sub LoadSprites(ByVal box As Single)
  Dim i As Integer
  For i = 0 To 27
    piece(i).BorderStyle = 0
    piece(i).Picture = LoadPicture(ThSprite("piece" & i, box))
  Next i
  blank.BorderStyle = 0
  blank.Picture = LoadPicture(ThSprite("blank", box))
  current.Picture = blank.Picture
  out(1).Picture = LoadPicture(ThSprite("stub", box))
  out(0).Picture = LoadPicture(ThSprite("arrow", box))
  inx(0).Picture = LoadPicture(ThSprite("stub", box))
  inx(1).Picture = LoadPicture(ThSprite("arrow", box))
End Sub

Private Sub Form_Load()
  ApplyLocale
  ThemeApply Me
  BackColor = PageY
  lblTitle.BackColor = PageY
  btnQuit.BackColor = WaterC
  LoadSprites 510
  Dim shp As Control
  For Each shp In Controls
    If TypeName(shp) = "Line" Then shp.BorderColor = CellLine
  Next shp
  boxcurrent.BorderColor = &H78FF&
  btnQuit.Width = 110
  btnQuit.Height = 40
  Dim i, j, k, n As Integer    'counters
  Dim t, l As Integer       'top,left
  tb = 50
  lb = 50
  blank.BorderStyle = 0
  si = blank.Width
  blank.BorderStyle = 1
  sb = si * 6

  ' set up nxt array

  For i = 0 To 24 Step 4
    For j = 0 To 2
     nxt(i + j) = i + j + 1
    Next j
    nxt(i + 3) = i
  Next i
        
  ' place pieces on board with blank icon

  For j = 0 To 5
    t = tb + si * j
    For i = 0 To 5
      n = i + j * 6
      l = lb + si * i
      Board(n).top = t
      Board(n).left = l
      Board(n).Visible = True
      Board(n).Picture = blank.Picture
      brdarray(n).pieceno = -1
      For k = 1 To 4
        brdarray(n).side(k).edge = False
        brdarray(n).side(k).empty = True
        brdarray(n).side(k).pipe = 0
      Next k
      brdarray(n).available = False
    Next i
  Next j

  ' mark edge of board

   For i = 0 To 5
     brdarray(i).side(1).edge = True     'top
     brdarray(i).side(1).pipe = -1
     brdarray(i * 6 + 5).side(2).edge = True 'right
     brdarray(i * 6 + 5).side(2).pipe = -1
     brdarray(i + 30).side(3).edge = True
     brdarray(i + 30).side(3).pipe = -1
     brdarray(i * 6).side(4).edge = True
     brdarray(i * 6).side(4).pipe = -1
   Next i

  ' draw board grid

  For j = 0 To 6
    t = tb + si * j
    Line1(j).X1 = t
    Line1(j).Y1 = lb
    Line1(j).X2 = t
    Line1(j).Y2 = lb + sb
    Line1(j + 7).X1 = tb
    Line1(j + 7).Y1 = t
    Line1(j + 7).X2 = tb + sb
    Line1(j + 7).Y2 = t
  
  Next j

  ' place water strips on board as invisible

  
  For j = 0 To 47
    Water(j).top = tb + j * sb / 48
    Water(j).left = lb
    Water(j).Height = 5
    Water(j).Width = sb
    Water(j).FillColor = CellC
    Water(j).Visible = True
  Next j
  wt = 47      ' first strip to turn on
  
  ' box to highlight selected piece

  boxcurrent.Height = blank.Height
  boxcurrent.Width = blank.Width
   
  ' place quit button

  btnQuit.top = 0.85 * Form1.ScaleHeight
  btnQuit.left = lb + (sb - btnQuit.Width) / 2

  ' title, and cards behind the board and the piece palette
  lblTitle.Font.Size = 26
  lblTitle.Font.Bold = True
  lblTitle.ForeColor = PipeDark
  lblTitle.BackStyle = 1
  lblTitle.BackColor = PageY
  lblTitle.Alignment = 2
  lblTitle.Left = lb
  lblTitle.Width = sb
  lblTitle.Top = 0
  lblTitle.Height = 46
  ThemeCardEx shpBoard, shpBoardS1, shpBoardS2, lb - 6, tb - 6, sb + 12, sb + 12, 1, CellC, EdgeC, &H28AAD6, &H38C6E8
  Dim pb As Single
  pb = blank.top
  For i = 0 To 27
    If piece(i).Visible And piece(i).left = blank.left Then
      If piece(i).top + piece(i).Height > pb Then pb = piece(i).top + piece(i).Height
    End If
  Next i
  ThemeCardEx shpPal, shpPalS1, shpPalS2, blank.left - 6, blank.top - 6, blank.Width + 12, pb - blank.top + 12, 1, &HD6FAFF, EdgeC, &H28AAD6, &H38C6E8

  ' place in and out pipes

  Randomize
  orow = Int(Rnd * 6)
  irow = Int(Rnd * 6)

  out(1).top = tb + orow * si
  out(1).left = lb - si
  out(0).top = out(1).top
  out(0).left = out(1).left - si / 2
  inx(1).top = tb + irow * si
  inx(1).left = lb + si * 7
  inx(0).top = inx(1).top
  inx(0).left = inx(1).left - si
  ScaleToScreen

  ' update info for adjacent squares


  i = orow * 6                    ' out on left
  'brdarray(i).available = True
  brdarray(i).side(4).pipe = 1
  brdarray(i).side(4).edge = False

  i = irow * 6 + 5                 ' in on right
  brdarray(i).available = True
  brdarray(i).side(2).pipe = 1
  brdarray(i).side(2).edge = False

  ' fill pipearray
  
  For i = 0 To 35
  pipearray(i) = False
  Next i

  ' fill piece info array
  ' -1 = no pipe
  ' 1 = pipe

  piecearray(0).top = 1
  piecearray(0).bottom = 1
  piecearray(0).left = -1
  piecearray(0).right = -1
  piecearray(1).top = -1
  piecearray(1).bottom = -1
  piecearray(1).left = 1
  piecearray(1).right = 1
  piecearray(2).top = 1
  piecearray(2).bottom = 1
  piecearray(2).left = -1
  piecearray(2).right = -1
  piecearray(3).top = -1
  piecearray(3).bottom = -1
  piecearray(3).left = 1
  piecearray(3).right = 1
  piecearray(4).top = 1
  piecearray(4).bottom = 1
  piecearray(4).left = 1
  piecearray(4).right = 1
  piecearray(5).top = 1
  piecearray(5).bottom = 1
  piecearray(5).left = 1
  piecearray(5).right = 1
  piecearray(6).top = 1
  piecearray(6).bottom = 1
  piecearray(6).left = 1
  piecearray(6).right = 1
  piecearray(7).top = 1
  piecearray(7).bottom = 1
  piecearray(7).left = 1
  piecearray(7).right = 1
  piecearray(8).top = 1
  piecearray(8).bottom = 1
  piecearray(8).left = 1
  piecearray(8).right = 1
  piecearray(9).top = 1
  piecearray(9).bottom = 1
  piecearray(9).left = 1
  piecearray(9).right = 1
  piecearray(10).top = 1
  piecearray(10).bottom = 1
  piecearray(10).left = 1
  piecearray(10).right = 1
  piecearray(11).top = 1
  piecearray(11).bottom = 1
  piecearray(11).left = 1
  piecearray(11).right = 1
  piecearray(12).top = 1
  piecearray(12).bottom = -1
  piecearray(12).left = 1
  piecearray(12).right = 1
  piecearray(13).top = 1
  piecearray(13).bottom = 1
  piecearray(13).left = 1
  piecearray(13).right = -1
  piecearray(14).top = -1
  piecearray(14).bottom = 1
  piecearray(14).left = 1
  piecearray(14).right = 1
  piecearray(15).top = 1
  piecearray(15).bottom = 1
  piecearray(15).left = -1
  piecearray(15).right = 1
  piecearray(16).top = 1
  piecearray(16).bottom = -1
  piecearray(16).left = -1
  piecearray(16).right = 1
  piecearray(17).top = 1
  piecearray(17).bottom = -1
  piecearray(17).left = 1
  piecearray(17).right = -1
  piecearray(18).top = -1
  piecearray(18).bottom = 1
  piecearray(18).left = 1
  piecearray(18).right = -1
  piecearray(19).top = -1
  piecearray(19).bottom = 1
  piecearray(19).left = -1
  piecearray(19).right = 1
  piecearray(20).top = 1
  piecearray(20).bottom = 1
  piecearray(20).left = 1
  piecearray(20).right = 1
  piecearray(21).top = 1
  piecearray(21).bottom = 1
  piecearray(21).left = 1
  piecearray(21).right = 1
  piecearray(22).top = 1
  piecearray(22).bottom = 1
  piecearray(22).left = 1
  piecearray(22).right = 1
  piecearray(23).top = 1
  piecearray(23).bottom = 1
  piecearray(23).left = 1
  piecearray(23).right = 1
  piecearray(24).top = -1
  piecearray(24).bottom = -1
  piecearray(24).left = 1
  piecearray(24).right = -1
  piecearray(25).top = -1
  piecearray(25).bottom = 1
  piecearray(25).left = -1
  piecearray(25).right = -1
  piecearray(26).top = -1
  piecearray(26).bottom = -1
  piecearray(26).left = -1
  piecearray(26).right = 1
  piecearray(27).top = 1
  piecearray(27).bottom = -1
  piecearray(27).left = -1
  piecearray(27).right = -1
  yy = 0
End Sub

Private Sub form_mousedown(button As Integer, Shift As Integer, x As Single, y As Single)

  If button = 2 Then
    rotatepiece
  End If
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
  current.Visible = False
  nb = -1
  If button = 2 Then
    rotatepiece
  End If

End Sub

Private Sub movepiece(x As Single, y As Single)
  current.top = y
  current.left = x

End Sub

Private Sub piece_MouseDown(index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
    
  nb = -1          ' not on board
  current.Visible = False
  dragging = (button = 1)
  wasSel = (button = 1 And np = index And boxcurrent.top = piece(index).top And boxcurrent.left = piece(index).left)
  downX = x
  downY = y
  If button > 0 Then
    np = index
    current.Picture = piece(index).Picture
    boxcurrent.top = piece(index).top
    boxcurrent.left = piece(index).left
    If button = 2 Then
      rotatepiece
    End If
  End If

End Sub

Private Sub rotatepiece()
  If np >= 0 Then
    piece(np).Visible = False
    np = nxt(np)
    piece(np).Visible = True
    current.Picture = piece(np).Picture

  End If

End Sub

Private Sub Timer1_Timer()
  If wt >= 0 Then
    Water(wt).FillColor = WaterC
    wt = wt - 1
  Else
    Timer1.Enabled = False
    MsgBox Loc("msg.lose")
    NewGame
  End If
End Sub

' While a piece is dragged over the palette it is over no square.
Private Sub piece_MouseMove(index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
  nb = -1
  current.Visible = False
End Sub

' The board square nearest a point that the selected piece fits, or -1 when
' none is within three quarters of a square (a finger covers what it points at).
Private Function CellNear(ByVal fx As Single, ByVal fy As Single) As Integer
  Dim i As Integer, d As Single, best As Single, dx As Single, dy As Single
  CellNear = -1
  best = si * 0.75
  For i = 0 To 35
    If Fits(i) Then
      dx = Abs(fx - (Board(i).left + Board(i).Width / 2))
      dy = Abs(fy - (Board(i).top + Board(i).Height / 2))
      d = dx
      If dy > d Then d = dy
      If d < best Then
        best = d
        CellNear = i
      End If
    End If
  Next i
End Function

' Releasing a dragged piece puts it on the square it fits nearest the point of
' release. A tap on the piece that is already selected turns it.
Private Sub piece_MouseUp(index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
  Dim c As Integer
  If dragging And button = 1 Then
    dragging = False
    c = CellNear(piece(index).left + x, piece(index).top + y)
    If c >= 0 Then
      Board_MouseDown c, 1, 0, 0, 0
    ElseIf wasSel And Abs(x - downX) < si / 4 And Abs(y - downY) < si / 4 Then
      rotatepiece
    End If
  End If
End Sub

' Makes the form fill the screen and the game grow to fit it: every control
' is scaled and centered, the sprites are those of the new size, and the
' board's own measures follow.
Private Sub ScaleToScreen()
  Dim c As Control, f As Single, n As Integer
  Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single
  Dim l As Single, t As Single, w As Single, h As Single
  Dim dx As Single, dy As Single
  x0 = 1E+30: y0 = 1E+30: x1 = -1E+30: y1 = -1E+30
  On Error Resume Next
  btnQuit.Visible = False
  For Each c In Controls
    w = 0: h = 0
    l = c.Left: t = c.Top: w = c.Width: h = c.Height
    If w > 0 And h > 0 And c.Visible Then
      If l < x0 Then x0 = l
      If t < y0 Then y0 = t
      If l + w > x1 Then x1 = l + w
      If t + h > y1 Then y1 = t + h
    End If
  Next c
  btnQuit.Visible = True
  If x1 < x0 Then Exit Sub
  ThemeFullScreen Me
  f = ScaleWidth * 0.94 / (x1 - x0)
  If ScaleHeight * 0.94 / (y1 - y0) < f Then f = ScaleHeight * 0.94 / (y1 - y0)
  dx = (ScaleWidth - (x1 - x0) * f) / 2 - x0 * f
  dy = (ScaleHeight - (y1 - y0) * f) / 2 - y0 * f
  For Each c In Controls
    If TypeName(c) = "Line" Then
      c.X1 = c.X1 * f + dx
      c.Y1 = c.Y1 * f + dy
      c.X2 = c.X2 * f + dx
      c.Y2 = c.Y2 * f + dy
    Else
      If TypeName(c) = "Image" Then c.Stretch = True
      w = 0: h = 0
      l = c.Left: t = c.Top: w = c.Width: h = c.Height
      If w > 0 And h > 0 Then
        c.Width = w * f
        c.Height = h * f
        c.Left = l * f + dx
        c.Top = t * f + dy
      End If
    End If
  Next c
  lblTitle.Font.Size = 26 * f
  ' the way out is an X in the upper right corner, big enough for a finger
  Dim side As Single
  side = ScaleHeight * 0.07
  btnQuit.Caption = "X"
  btnQuit.BackColor = &H7A7AF0
  btnQuit.Width = side
  btnQuit.Height = side
  btnQuit.Left = ScaleWidth - side - side / 4
  btnQuit.Top = side / 4
  btnQuit.Font.Size = side * 0.35
  btnQuit.Font.Bold = True
  tb = Board(0).Top
  lb = Board(0).Left
  si = Board(0).Width
  sb = si * 6
  LoadSprites 510 * f
  For n = 0 To 35
    Board(n).Picture = blank.Picture
  Next n
End Sub

' Starts another game after a win or a loss: an empty board, the water back
' at the top, and the inlet and outlet at new rows.
Private Sub NewGame()
  Dim i As Integer, k As Integer, n As Integer
  For n = 0 To 35
    Board(n).Picture = blank.Picture
    Board(n).Visible = True
    brdarray(n).pieceno = -1
    For k = 1 To 4
      brdarray(n).side(k).edge = False
      brdarray(n).side(k).empty = True
      brdarray(n).side(k).pipe = 0
    Next k
    brdarray(n).available = False
    pipearray(n) = False
  Next n
  For i = 0 To 5
    brdarray(i).side(1).edge = True
    brdarray(i).side(1).pipe = -1
    brdarray(i * 6 + 5).side(2).edge = True
    brdarray(i * 6 + 5).side(2).pipe = -1
    brdarray(i + 30).side(3).edge = True
    brdarray(i + 30).side(3).pipe = -1
    brdarray(i * 6).side(4).edge = True
    brdarray(i * 6).side(4).pipe = -1
  Next i
  orow = Int(Rnd * 6)
  irow = Int(Rnd * 6)
  out(1).top = tb + orow * si
  out(1).left = lb - si
  out(0).top = out(1).top
  out(0).left = out(1).left - si / 2
  inx(1).top = tb + irow * si
  inx(1).left = lb + si * 7
  inx(0).top = inx(1).top
  inx(0).left = inx(1).left - si
  i = orow * 6
  brdarray(i).side(4).pipe = 1
  brdarray(i).side(4).edge = False
  i = irow * 6 + 5
  brdarray(i).available = True
  brdarray(i).side(2).pipe = 1
  brdarray(i).side(2).edge = False
  For i = 0 To 47
    Water(i).FillColor = CellC
  Next i
  wt = 47
  yy = 0
  nb = -1
  dragging = False
  current.Visible = False
  Timer1.Enabled = True
End Sub
