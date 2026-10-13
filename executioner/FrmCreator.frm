VERSION 5.00
Begin VB.Form FrmCreator 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   188
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "    creptic@gmail.com or visit  http://cmcinroy.comze.com"
      Top             =   4780
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "FrmCreator.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   5175
      TabIndex        =   3
      Top             =   4680
      Width           =   5175
   End
   Begin VB.PictureBox Littleabout 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   960
      Picture         =   "FrmCreator.frx":0DB6
      ScaleHeight     =   615
      ScaleWidth      =   3735
      TabIndex        =   2
      Top             =   0
      Width           =   3735
   End
   Begin VB.PictureBox About 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   480
      Picture         =   "FrmCreator.frx":1B63
      ScaleHeight     =   1695
      ScaleWidth      =   4815
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.PictureBox PicAbout 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1560
      ScaleHeight     =   2295
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   90
      Left            =   360
      Top             =   2160
   End
End
Attribute VB_Name = "FrmCreator"
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
' Variables Used in frmCreator
' *********************************************************************************************************
Dim Endtime As Integer ' Used as a counter to unload the form if no action is made
Dim x As Integer
' The Following Variable ( Playsound ) is used to setup windows sound for playback
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
(ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Sub Form_Load()

' *********************************************************************************************************
' FORM LOAD
' *********************************************************************************************************
' Purpose: To play a Sound when form is loaded and to set the endtime to 0
'
' Variables Used: Endtime (reset Timer1 to 0. Playsound ( used to play sound
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

Endtime = 0
If bSound = True Then PlaySound ("sounds\hey.wav"), 0, 1

End Sub

Private Sub PicAbout_Click()

' *********************************************************************************************************
' WHEN PIC CLICKED EXIT
' *********************************************************************************************************
' Purpose: To exit when Picture os clicked by the user
'
' Variables Used: none
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************

Unload FrmCreator
End Sub

Private Sub Timer1_Timer()
 
' *********************************************************************************************************
' TIMER1_TIMER
' *********************************************************************************************************
' Purpose: To display Form and unload form after Endtime = 90
'
' Variables Used: Endtime, x (both as integer type Variables )
'
' Written By: Corey McInroy
' Date Modified: May 29/99
' *********************************************************************************************************
  
 ' Dim x As Integer
   '  x = Int((12 * Rnd) + 1)
     x = x + 1
    If x = 2 Then
       PicAbout.Picture = LoadPicture("grafix\about\1.jpg")
       PicAbout.Refresh
     End If
    If x = 4 Then
       PicAbout.Picture = LoadPicture("grafix\about\2.jpg")
       PicAbout.Refresh
     End If
    If x = 6 Then
       PicAbout.Picture = LoadPicture("grafix\about\3.jpg")
       PicAbout.Refresh
     End If
    If x = 8 Then
       PicAbout.Picture = LoadPicture("grafix\about\4.jpg")
       PicAbout.Refresh
     End If
    If x = 10 Then
       PicAbout.Picture = LoadPicture("grafix\about\5.jpg")
       PicAbout.Refresh
     End If
    If x = 12 Then
       PicAbout.Picture = LoadPicture("grafix\about\6.jpg")
       PicAbout.Refresh
     End If
If x = 13 Then x = 0
    If Endtime = 120 Then Unload FrmCreator ' if endtime = 90 then unload form
    
       Endtime = Endtime + 1 ' increase endtime by one
       
 
 End Sub

