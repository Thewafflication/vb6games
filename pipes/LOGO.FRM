VERSION 5.00
Begin VB.Form Logo 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Pipes version 1.0"
   ClientHeight    =   3375
   ClientLeft      =   1455
   ClientTop       =   1635
   ClientWidth     =   2880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   192
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   372
      Left            =   960
      TabIndex        =   0
      Top             =   2880
      Width           =   852
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Height          =   3372
      Left            =   0
      Top             =   0
      Width           =   2892
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kym and Jonathan Wilson"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2208
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   4500
      Left            =   0
      Picture         =   "Logo.frx":0000
      Top             =   0
      Width           =   3600
   End
End
Attribute VB_Name = "Logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Form_Unload (False)
End Sub

Private Sub Form_Load()
  Dim sw, sh As Integer
  Dim fw, fh As Integer
  sw = Screen.Width
  sh = Screen.Height

  fw = Image1.Width
  fh = Image1.Height
  Logo.Width = fw * Screen.TwipsPerPixelX
  Logo.Height = fh * Screen.TwipsPerPixelY
  Shape1.Width = fw
  Shape1.Height = fh
  Label1.left = fw / 2
  Label1.top = fh / 10
  Label2.top = fh * 0.7
  Label2.left = (fw - Label2.Width) / 2
  Command1.top = fh * 0.8
  Command1.left = (fw - Command1.Width) / 2
  Load Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Form1.Show
  Form1.Timer1.Enabled = True
  Unload Logo
End Sub

