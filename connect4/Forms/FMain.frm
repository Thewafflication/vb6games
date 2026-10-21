VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Connect4"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12255
   FillStyle       =   0  'Ausgefüllt
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3915
      ScaleWidth      =   6555
      TabIndex        =   7
      Top             =   360
      Width           =   6615
   End
   Begin VB.PictureBox PBBackBuffer 
      AutoRedraw      =   -1  'True
      Height          =   7740
      Left            =   120
      Picture         =   "FMain.frx":8023
      ScaleHeight     =   7680
      ScaleWidth      =   9360
      TabIndex        =   8
      Top             =   480
      Width           =   9420
   End
   Begin VB.PictureBox PBNextPlayer 
      Height          =   375
      Left            =   4800
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   0
      Width           =   855
      Begin VB.Label LbPlayerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "______"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   60
         Width           =   450
      End
   End
   Begin VB.CommandButton BtnRedo 
      Style           =   1  \'Graphical
      Caption         =   "Redo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton BtnUndo 
      Style           =   1  \'Graphical
      Caption         =   "Undo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton BtnNewGame 
      Style           =   1  \'Graphical
      Caption         =   "New Game"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton BtnInfo 
      Style           =   1  \'Graphical
      Caption         =   "Info"
      Height          =   375
      Left            =   11400
      TabIndex        =   6
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Next Player:"
      Height          =   225
      Left            =   3600
      TabIndex        =   3
      Top             =   60
      Width           =   945
   End
   Begin VB.Shape shpHead 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpHeadS1 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape shpHeadS2 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Connect4   As Connect4
Private m_NextPlayer As Player
Private m_Player1    As Player
Private m_Player2    As Player
Private m_Undo       As UndoRedo
Private m_LastButton As Integer
Private m_LastX      As Single
Private m_LastY      As Single
'https://de.wikipedia.org/wiki/Vier_gewinnt
'https://de.wikipedia.org/wiki/Gel%C3%B6ste_Spiele
'https://de.wikihow.com/Bei-Vier-Gewinnt-gewinnen

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

Private Sub ApplyLocale()
    If LocLanguage = "en" Then Exit Sub
    Caption = Loc("app.title")
    Label1.Caption = Loc("main.nextPlayer")
    BtnNewGame.Caption = Loc("main.newGame")
    BtnUndo.Caption = Loc("main.undo")
    BtnRedo.Caption = Loc("main.redo")
    BtnInfo.Caption = Loc("main.info")
End Sub

Private Sub Form_Load()
    ApplyLocale
    ThemeFullScreen Me
    LbPlayerName.Tag = "keep"
    ThemeApply Me
    PBBackBuffer.ScaleMode = vbTwips
    Picture1.ScaleMode = vbTwips
    LbPlayerName.ForeColor = ThBack
    LbPlayerName.Font.Bold = True
    Picture1.BorderStyle = 0
    PBBackBuffer.BorderStyle = 0
    PBNextPlayer.BorderStyle = 0
    Label1.BackColor = ThPanel
    LbPlayerName.AutoSize = False
    LbPlayerName.Alignment = 2
    Set PBBackBuffer.Picture = Nothing
    Set m_Player1 = MApp.DefaultPlayer1.Clone
    Set m_Player2 = MApp.DefaultPlayer2.Clone
    Set m_NextPlayer = m_Player1
    Set m_Connect4 = MNew.Connect4(MApp.DefaultWidth, MApp.DefaultHeight)
    Set m_Undo = New UndoRedo
    SaveUndo
End Sub

Private Sub Form_Resize()
    ' A header card with buttons big enough for a finger, and the board below.
    Const M As Single = 240
    Const Th As Single = 720
    Const Pad As Single = 150
    Dim hw As Single: hw = Me.ScaleWidth - 2 * M
    Dim hh As Single: hh = Th + 2 * Pad
    Dim y As Single: y = M + Pad
    Dim x As Single: x = M + Pad
    Dim bw As Single: bw = (hw - 2 * Pad - 6140) / 3.3
    If bw > 1800 Then bw = 1800
    If bw < 700 Then bw = 700
    ThemeCard shpHead, shpHeadS1, shpHeadS2, M, M, hw, hh, 15
    BtnNewGame.Move x, y, bw * 1.3, Th
    x = x + bw * 1.3 + 120
    BtnUndo.Move x, y, bw, Th
    x = x + bw + 120
    BtnRedo.Move x, y, bw, Th
    x = x + bw + 360
    Label1.Move x, y + 200, 1500, 320
    x = x + 1500
    PBNextPlayer.Move x, y, 2300, Th
    LbPlayerName.Move 0, 200, 2300, 320
    BtnInfo.Move Me.ScaleWidth - M - Pad - 1500, y, 1500, Th
    Dim ty As Single: ty = M + hh + M
    Dim h As Single: h = Me.ScaleHeight - ty - M
    If Me.ScaleWidth > 0 And h > 0 Then
        PBBackBuffer.Move M, ty, hw, h
        Picture1.Move M, ty, hw, h
    End If
End Sub

Private Sub BtnNewGame_Click()
    If FNewGame.ShowDialog(Me, m_Connect4, m_Player1, m_Player2) = vbCancel Then Exit Sub
    Set m_Undo = New UndoRedo
    SaveUndo
    Picture1_Paint
End Sub

Private Sub PBBackBuffer_Resize()
    Picture1_Paint
End Sub

Private Sub Picture1_Click()
    'Debug.Print "Picture1_Click"
    PictureMouseClick m_LastButton, m_LastX, m_LastY
End Sub

Private Sub Picture1_DblClick()
    'Debug.Print "Picture1_DblClick"
    'PictureMouseClick m_LastButton, m_LastX, m_LastY
    PictureMouseClick m_LastButton, m_LastX, m_LastY
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Debug.Print "Picture1_MouseDown"
    m_LastButton = Button: m_LastX = X: m_LastY = Y
    'PictureMouseClick Button, X, Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Debug.Print "Picture1_MouseMove"
    'm_LastButton = Button:
    m_LastX = X: m_LastY = Y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Debug.Print "Picture1_MouseUp"
    'm_LastButton = Button: m_LastX = X: m_LastY = Y
End Sub

Private Sub PictureMouseClick(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> MouseButtonConstants.vbLeftButton Then Exit Sub
    If m_Connect4.IsGameOver Then
        MsgBox Loc("msg.gameOver")
        Exit Sub
    End If
    
    Dim ci As Byte: ci = CalcColumnIndex(Picture1, X)
    m_Connect4.NextPlayerDropChip ci
    
    Set m_NextPlayer = IIf(m_Connect4.NextPlayer = Connect4Chip.Player1, m_Player1, m_Player2)
        
    UpdateView
    
    Dim p As Connect4Chip: p = m_Connect4.Check4
    If p <> None Then
        Picture1_Paint
        Dim S As String: S = LocF("player.label", IIf(p = Player1, "1", "2"), IIf(p = Player1, m_Player1.Name, m_Player2.Name))
        MsgBox LocF("msg.won", S)
    End If
    If m_Connect4.CountDown = 0 And m_Connect4.IsGameOver Then
        MsgBox Loc("msg.drawn")
    End If
    SaveUndo
End Sub

Private Sub Picture1_Paint()
    UpdateView
End Sub

Sub SaveUndo()
    m_Undo.SaveUndo m_Connect4.Clone
    EnableUndoRedoButtons
End Sub

Sub EnableUndoRedoButtons()
    BtnUndo.Enabled = m_Undo.CanUndo
    BtnRedo.Enabled = m_Undo.CanRedo
End Sub

Private Sub BtnUndo_Click()
    If m_Undo.CanUndo Then
        'Dim c4 As Connect4: Set c4 = m_Undo.Undo
        m_Connect4.NewC m_Undo.Undo 'c4
    End If
    UpdateView
End Sub

Private Sub BtnRedo_Click()
    If m_Undo.CanRedo Then
        'Dim c4 As Connect4: Set c4 = m_Undo.Redo
        m_Connect4.NewC m_Undo.Redo 'c4
    End If
    UpdateView
End Sub

Sub UpdateView()
    DrawConnect4Field PBBackBuffer
    Set Picture1.Picture = PBBackBuffer.Image
    If m_NextPlayer Is Nothing Then
        PBNextPlayer.BackColor = ThBoard
        LbPlayerName.Caption = "________"
    Else
        Set m_NextPlayer = IIf(m_Connect4.NextPlayer = Connect4Chip.Player1, m_Player1, m_Player2)
        PBNextPlayer.BackColor = m_NextPlayer.Color
        LbPlayerName.Caption = m_NextPlayer.Name
    End If
    EnableUndoRedoButtons
End Sub

Public Sub DrawConnect4Field(aPB As PictureBox)
    aPB.Cls
    aPB.BackColor = ThBack
    Dim w As Byte: w = m_Connect4.Width
    Dim h As Byte: h = m_Connect4.Height
    'maximum chip diameter
    Dim d As Double: d = GetChipDiameter(aPB, w, h)
    Dim r As Double: r = d / 2# - d / 14
    'left most point and the top of the board
    Dim l0 As Double: l0 = aPB.ScaleWidth / 2# - d * w / 2#
    Dim t0 As Double: t0 = aPB.ScaleHeight / 2# + d * h / 2#
    Dim pad As Double: pad = d / 5
    Dim bl As Double, bt As Double, bw As Double, bh As Double
    bl = l0 - pad: bt = t0 - d * h - pad: bw = d * w + 2 * pad: bh = d * h + 2 * pad
    ' a soft shadow, the board with a lighter edge, and its inner bevel
    aPB.Line (bl + d / 5, bt + d / 4)-(bl + bw + d / 5, bt + bh + d / 4), ThShadow2, BF
    aPB.Line (bl + d / 9, bt + d / 8)-(bl + bw + d / 9, bt + bh + d / 8), ThShadow1, BF
    aPB.Line (bl, bt)-(bl + bw, bt + bh), ThShade(ThBoard, 60), BF
    aPB.Line (bl + d / 40, bt + d / 40)-(bl + bw - d / 40, bt + bh - d / 40), ThShade(ThBoard, 130), BF
    aPB.Line (bl + d / 20, bt + d / 20)-(bl + bw - d / 20, bt + bh - d / 20), ThBoard, BF
    Dim l  As Double:  l = l0
    Dim t  As Double: t = t0
    Dim i As Long, j As Long, fc As Long
    Dim aChip As Connect4Chip
    Dim cx As Double, cy As Double
    aPB.FillStyle = vbFSSolid
    aPB.DrawWidth = Max(d / 300, 1)
    For j = 1 To h
        For i = 1 To w
            aChip = m_Connect4.Field(i, j)
            cx = l + d / 2
            cy = t - d / 2
            ' the hole, shaded on its upper edge
            aPB.FillColor = ThShade(ThBoard, 55)
            aPB.Circle (cx, cy), r + d / 40, ThShade(ThBoard, 55)
            aPB.FillColor = ThBack
            aPB.Circle (cx + d / 50, cy + d / 50), r - d / 50, ThBack
            If aChip <> None Then
                If (aChip And Player1) = Player1 Then
                    fc = m_Player1.Color
                Else
                    fc = m_Player2.Color
                End If
                aPB.FillColor = ThShade(fc, 70)
                aPB.Circle (cx, cy), r - d / 40, ThShade(fc, 70)
                aPB.FillColor = fc
                aPB.Circle (cx, cy), r - d / 18, fc
                aPB.FillStyle = vbFSTransparent
                aPB.Circle (cx, cy), r * 0.62, ThShade(fc, 82)
                aPB.FillStyle = vbFSSolid
                aPB.FillColor = ThShade(fc, 130)
                aPB.Circle (cx - r * 0.38, cy - r * 0.38), r * 0.13, ThShade(fc, 130)
                If aChip And Highlighted Then
                    aPB.FillStyle = vbFSTransparent
                    aPB.DrawWidth = Max(d / 160, 2)
                    aPB.Circle (cx, cy), r + d / 40, ThGood
                    aPB.DrawWidth = Max(d / 300, 1)
                    aPB.FillStyle = vbFSSolid
                End If
            End If
            l = l + d
        Next
        l = l0
        t = t - d
    Next
End Sub

Function GetChipDiameter(aPB As PictureBox, ByVal w As Byte, ByVal h As Byte) As Double
    ' leave room around the board for its frame and shadow
    Dim m As Double: m = Min(aPB.ScaleWidth, aPB.ScaleHeight) * 0.06
    GetChipDiameter = Min((aPB.ScaleWidth - 2 * m) / w, (aPB.ScaleHeight - 2 * m) / h)
End Function

Public Function CalcColumnIndex(aPB As PictureBox, ByVal X As Single) As Byte
    'maximum chip-diameter
    Dim w As Byte: w = m_Connect4.Width
    Dim h As Byte: h = m_Connect4.Height
    Dim d As Double: d = GetChipDiameter(aPB, w, h)
    'left most point
    Dim l0 As Double: l0 = aPB.ScaleWidth / 2# - d * w / 2#
    Dim l  As Double: l = l0
    Dim l1 As Double
    Dim i As Long
    For i = 1 To w '- 1
        l1 = l + d
        If l <= X And X <= l1 Then
            CalcColumnIndex = i
            Exit Function
        End If
        l = l0 + i * d
    Next
End Function

Private Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

