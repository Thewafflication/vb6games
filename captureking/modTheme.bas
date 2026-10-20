Attribute VB_Name = "modTheme"
Option Explicit
'
' One look for every game: a dark navy background, raised panels, light text
' and a teal accent. Colors are VB values (&HBBGGRR).
'
' ui/modTheme.bas is the source; ui/sync.py copies it into each game.
'
Public Const ThBack As Long = &H1F1411     ' window background
Public Const ThPanel As Long = &H33201B    ' raised surface
Public Const ThPanel2 As Long = &H452B25   ' buttons and inputs
Public Const ThLine As Long = &H5C3C35     ' borders
Public Const ThText As Long = &HF5EBE8     ' text
Public Const ThMuted As Long = &HBEA098    ' secondary text
Public Const ThAccent As Long = &HC6D63D   ' teal
Public Const ThAccent2 As Long = &H54B4FF  ' amber
Public Const ThGood As Long = &H8DD35F
Public Const ThBad As Long = &H816BFF
Public Const ThP1 As Long = &H4DD3FF       ' player one: yellow
Public Const ThP2 As Long = &H7A5CFF       ' player two: coral
Public Const ThBoard As Long = &H8F3E2D    ' game board blue
Public Const ThFontSize As Single = 11      ' text size in points

' Colors a form and its controls. A control whose Tag is "keep" is left alone
' (game boards, pictures); one whose Tag is "accent" is a highlighted button.
' Push buttons need Style = 1 (Graphical) in the form to take BackColor.
Public Sub ThemeApply(frm As Form)
    Dim c As Control
    frm.BackColor = ThBack
    On Error Resume Next
    For Each c In frm.Controls
        If c.Tag <> "keep" Then
            Select Case TypeName(c)
            Case "Label"
                c.ForeColor = ThText
                c.BackColor = ThBack
                c.BackStyle = 1
                c.Font.Size = ThFontSize
            Case "CommandButton"
                ' A VB6 push button has no ForeColor, so its text stays dark
                ' and the buttons are light (teal; amber when tagged accent).
                If c.Tag = "accent" Then
                    c.BackColor = ThAccent2
                Else
                    c.BackColor = ThAccent
                End If
                c.Font.Size = ThFontSize
                c.Font.Bold = True
            Case "TextBox", "ComboBox", "ListBox"
                c.ForeColor = ThText
                c.BackColor = ThPanel2
                c.Font.Size = ThFontSize
            Case "Frame", "CheckBox", "OptionButton"
                c.ForeColor = ThText
                c.BackColor = ThBack
                c.BackStyle = 1
                c.Font.Size = ThFontSize
            End Select
        End If
    Next c
End Sub

' Makes the form as big as the screen (it only grows). With WFC's full-screen
' mode Screen.Width and Screen.Height are the screen in the program's own
' scaled units, so the form fills the screen at any resolution.
Public Sub ThemeFullScreen(frm As Form)
    If Screen.Width > frm.Width Or Screen.Height > frm.Height Then
        frm.Move 0, 0, Screen.Width, Screen.Height
    End If
End Sub
