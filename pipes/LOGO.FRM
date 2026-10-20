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
      Style           =   1  'Graphical
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
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PIPES"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   0
      TabIndex        =   3
      Top             =   70
      Width           =   192
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

Private Sub ApplyLocale()
  If LocLanguage = "en" Then Exit Sub
  Caption = Loc("app.title")
  Label1.Caption = Loc("logo.version")
  Command1.Caption = Loc("button.ok")
End Sub

Private Sub Form_Load()
  ApplyLocale
  ThemeApply Me
  Image1.Visible = False
  Shape1.BorderColor = ThLine
  Shape1.Width = ScaleWidth
  Shape1.Height = ScaleHeight
  lblTitle.ForeColor = ThAccent
  lblTitle.Font.Size = 36
  lblTitle.Left = 0
  lblTitle.Width = ScaleWidth
  lblTitle.Top = 45
  Label1.Alignment = 2
  Label1.Left = 0
  Label1.Width = ScaleWidth
  Label1.Top = 115
  Label2.Alignment = 2
  Label2.Left = 0
  Label2.Width = ScaleWidth
  Label2.Top = 145
  Command1.Width = 90
  Command1.Height = 34
  Command1.Left = (ScaleWidth - Command1.Width) / 2
  Command1.Top = 180
  Load Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Form1.Show
  Form1.Timer1.Enabled = True
  Unload Logo
End Sub

