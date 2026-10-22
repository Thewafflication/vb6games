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
Public Const ThShadow1 As Long = &H0F0A08  ' near shadow
Public Const ThShadow2 As Long = &H17100D  ' far shadow
Public Const ThEdge As Long = &H6B4A40     ' light edge of a raised panel
Public Const ThFontSize As Single = 11      ' text size in points
Public Const ThFontName As String = "Segoe UI"  ' scalable, unlike MS Sans Serif

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
                c.Font.Name = ThFontName
                c.Font.Size = ThFontSize
                ' an auto-sized label grows to its new font
                If c.AutoSize Then
                    c.AutoSize = False
                    c.AutoSize = True
                    c.Caption = c.Caption
                End If
            Case "CommandButton"
                ' A VB6 push button has no ForeColor, so its text stays dark
                ' and the buttons are light (teal; amber when tagged accent).
                If c.Tag = "accent" Then
                    c.BackColor = ThAccent2
                Else
                    c.BackColor = ThAccent
                End If
                c.Font.Name = ThFontName
                c.Font.Size = ThFontSize
                c.Font.Bold = True
            Case "TextBox", "ComboBox", "ListBox"
                c.ForeColor = ThText
                c.BackColor = ThPanel2
                c.Font.Name = ThFontName
                c.Font.Size = ThFontSize
            Case "Frame", "CheckBox", "OptionButton"
                c.ForeColor = ThText
                c.BackColor = ThBack
                c.BackStyle = 1
                c.Font.Name = ThFontName
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

' A raised panel with a thin border and a two-step shadow. The three shapes
' sit at the back of the form (shadows last); u is the size of one pixel in
' the form's own units (1 for pixel forms, 15 for twips).
Public Sub ThemeCard(card As Object, shadowA As Object, shadowB As Object, ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single, ByVal u As Single)
    shadowB.Left = L + 7 * u
    shadowB.Top = T + 9 * u
    shadowB.Width = W
    shadowB.Height = H
    shadowB.FillStyle = 0
    shadowB.FillColor = ThShadow2
    shadowB.BorderStyle = 0
    shadowA.Left = L + 3 * u
    shadowA.Top = T + 4 * u
    shadowA.Width = W
    shadowA.Height = H
    shadowA.FillStyle = 0
    shadowA.FillColor = ThShadow1
    shadowA.BorderStyle = 0
    card.Left = L
    card.Top = T
    card.Width = W
    card.Height = H
    card.FillStyle = 0
    card.FillColor = ThPanel
    card.BorderColor = ThEdge
    card.BorderWidth = 1
End Sub

' A color made lighter (pct > 100) or darker (pct < 100).
Public Function ThShade(ByVal c As Long, ByVal pct As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = (c And &HFF&) * pct \ 100
    g = ((c \ &H100&) And &HFF&) * pct \ 100
    b = ((c \ &H10000) And &HFF&) * pct \ 100
    If r > 255 Then r = 255
    If g > 255 Then g = 255
    If b > 255 Then b = 255
    ThShade = RGB(r, g, b)
End Function

' Moves every control of a form by dx, dy (to center a small layout on a
' screen-sized form). Timers, menus and the like have no position.
Public Sub ThemeShift(frm As Form, ByVal dx As Single, ByVal dy As Single)
    Dim c As Control
    On Error Resume Next
    For Each c In frm.Controls
        c.Left = c.Left + dx
        c.Top = c.Top + dy
    Next c
End Sub

' Shows a picture centered in the upper part of a screen-sized form, on a
' card with a shadow (ratio is its width over its height, u the size of a
' pixel in the form's units). Returns the bottom edge of the card.
Public Function ThemeShowArt(frm As Form, img As Object, card As Object, shadowA As Object, shadowB As Object, ByVal ratio As Single, ByVal u As Single) As Single
    Dim sw As Single, sh As Single, ah As Single, aw As Single
    sw = frm.ScaleWidth
    sh = frm.ScaleHeight
    ah = sh * 0.6
    aw = ah * ratio
    If aw > sw * 0.8 Then
        aw = sw * 0.8
        ah = aw / ratio
    End If
    img.Stretch = True
    img.Left = (sw - aw) / 2
    img.Top = sh * 0.07
    img.Width = aw
    img.Height = ah
    img.Visible = True
    ThemeCard card, shadowA, shadowB, img.Left - 4 * u, img.Top - 4 * u, aw + 8 * u, ah + 8 * u, u
    ThemeShowArt = img.Top + ah + 4 * u
End Function

' True when a control sits directly on the form (not inside a frame).
Private Function OnForm(frm As Form, c As Object) As Boolean
    On Error Resume Next
    OnForm = True
    OnForm = (c.Container Is frm)
End Function

' Makes a form fill the screen and shows its controls as one dialog: the
' group is centered on a card with a shadow (the card shapes are the form's
' own, named shpDlg, shpDlgS1 and shpDlgS2). u is one pixel in the form's units.
Public Sub ThemeDialog(frm As Form, card As Object, shadowA As Object, shadowB As Object, ByVal u As Single, ByVal pad As Single)
    Dim c As Control
    Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single
    Dim l As Single, t As Single, w As Single, h As Single
    Dim dx As Single, dy As Single
    x0 = 1E+30: y0 = 1E+30: x1 = -1E+30: y1 = -1E+30
    frm.BackColor = ThBack
    ThemeFullScreen frm
    On Error Resume Next
    For Each c In frm.Controls
        If Left(c.Name, 6) <> "shpDlg" And OnForm(frm, c) Then
            w = 0: h = 0
            l = c.Left: t = c.Top: w = c.Width: h = c.Height
            If w > 0 And h > 0 And c.Visible Then
                If l < x0 Then x0 = l
                If t < y0 Then y0 = t
                If l + w > x1 Then x1 = l + w
                If t + h > y1 Then y1 = t + h
            End If
        End If
    Next c
    If x1 < x0 Then Exit Sub
    dx = (frm.ScaleWidth - (x1 - x0)) / 2 - x0
    dy = (frm.ScaleHeight - (y1 - y0)) / 2 - y0
    For Each c In frm.Controls
        If Left(c.Name, 6) <> "shpDlg" And OnForm(frm, c) Then
            c.Left = c.Left + dx
            c.Top = c.Top + dy
        End If
    Next c
    ThemeCard card, shadowA, shadowB, x0 + dx - pad, y0 + dy - pad, (x1 - x0) + 2 * pad, (y1 - y0) + 2 * pad, u
    ' what sits on the card takes the card's color
    For Each c In frm.Controls
        Select Case TypeName(c)
        Case "Label", "OptionButton", "CheckBox", "Frame"
            If c.Tag <> "keep" Then c.BackColor = ThPanel
        End Select
    Next c
End Sub

' Shrinks a control's font until its caption fits (no clipping in any
' language or at any size). Uses the form's own text measuring.
Public Sub ThemeFitText(frm As Form, c As Object)
    Dim t As String, sz As Single, w As Single, th As Single
    Dim avail As Single, wrap As Boolean, lines As Single, ok As Boolean, au As Boolean
    Dim words() As String, k As Long
    On Error Resume Next
    t = c.Caption
    If Len(t) = 0 Then Exit Sub
    If InStr(t, vbLf) > 0 Or InStr(t, vbCr) > 0 Then Exit Sub
    wrap = (TypeName(c) = "Label")  ' a label wraps in its box
    au = False
    au = c.AutoSize
    If au Then Exit Sub
    sz = c.Font.Size
    words = Split(t, " ")
    frm.Font.Name = c.Font.Name
    frm.Font.Bold = c.Font.Bold
    avail = c.Width * 0.88
    If TypeName(c) = "OptionButton" Or TypeName(c) = "CheckBox" Then avail = c.Width - 24 * (c.Width / 100)
    Do
        frm.Font.Size = sz
        w = frm.TextWidth(t)
        th = frm.TextHeight("Ag")
        If wrap Then
            lines = Int(w / (avail * 0.85)) + 1
            ok = (lines * th * 1.15 <= c.Height)
            ' a single long word must fit on a line too
            For k = 0 To UBound(words)
                If frm.TextWidth(words(k)) > avail Then ok = False
            Next k
        Else
            ok = (w <= avail And th * 1.1 <= c.Height)
        End If
        If ok Or sz <= 6 Then Exit Do
        sz = sz - 0.5
    Loop
    c.Font.Size = sz
End Sub

' Fits the text of every label, button and option of a form (call after
' the layout and the captions are final).
Public Sub ThemeFitAll(frm As Form)
    Dim c As Control
    On Error Resume Next
    For Each c In frm.Controls
        Select Case TypeName(c)
        Case "Label", "CommandButton", "OptionButton", "CheckBox"
            ThemeFitText frm, c
        End Select
    Next c
End Sub

' A card in any colors (fill, edge, near and far shadow).
Public Sub ThemeCardEx(card As Object, shadowA As Object, shadowB As Object, ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single, ByVal u As Single, ByVal fill As Long, ByVal edge As Long, ByVal near As Long, ByVal far As Long)
    ThemeCard card, shadowA, shadowB, L, T, W, H, u
    shadowB.FillColor = far
    shadowA.FillColor = near
    card.FillColor = fill
    card.BorderColor = edge
End Sub

' The width of the screen in real pixels (Screen reports the program's own
' scaled units in WFC's full-screen mode).
Public Function ThScreenPixels() As Long
    ThScreenPixels = CLng(Screen.Width / Screen.TwipsPerPixelX)
End Function

' The size of the sprite set that suits a picture box of boxTwips on this
' screen: the sprites come in several sizes (32 to 192 pixels) and the one
' nearest the box's size in real pixels is used, so that nothing is enlarged
' or squeezed much.
Public Function ThSpriteSize(ByVal boxTwips As Single) As Long
    Dim want As Single
    want = boxTwips / Screen.TwipsPerPixelX
    If want <= 40 Then
        ThSpriteSize = 32
    ElseIf want <= 56 Then
        ThSpriteSize = 48
    ElseIf want <= 80 Then
        ThSpriteSize = 64
    ElseIf want <= 112 Then
        ThSpriteSize = 96
    ElseIf want <= 144 Then
        ThSpriteSize = 128
    ElseIf want <= 176 Then
        ThSpriteSize = 160
    Else
        ThSpriteSize = 192
    End If
End Function

' The file of a sprite (a GIF in sprites\<size>) in the set for a box of
' boxTwips.
Public Function ThSprite(ByVal name As String, ByVal boxTwips As Single) As String
    ThSprite = App.Path & "\sprites\" & CStr(ThSpriteSize(boxTwips)) & "\" & name & ".gif"
End Function

' The same for a sprite kept as an icon (its directory size is the size the
' program lays it out at; the image has more pixels). Sets stop at 160.
Public Function ThSpriteIcon(ByVal name As String, ByVal boxTwips As Single) As String
    Dim sz As Long
    sz = ThSpriteSize(boxTwips)
    If sz > 160 Then sz = 160
    ThSpriteIcon = App.Path & "\sprites\" & CStr(sz) & "\" & name & ".ico"
End Function

' Makes a form fill the screen with one color and centers its controls on it
' (for art-driven dialogs that bring their own frames).
Public Sub ThemeCenter(frm As Form, ByVal back As Long)
    Dim c As Control
    Dim x0 As Single, y0 As Single, x1 As Single, y1 As Single
    Dim l As Single, t As Single, w As Single, h As Single
    Dim dx As Single, dy As Single
    x0 = 1E+30: y0 = 1E+30: x1 = -1E+30: y1 = -1E+30
    frm.BackColor = back
    ThemeFullScreen frm
    On Error Resume Next
    For Each c In frm.Controls
        If OnForm(frm, c) Then
            w = 0: h = 0
            l = c.Left: t = c.Top: w = c.Width: h = c.Height
            If w > 0 And h > 0 And c.Visible Then
                If l < x0 Then x0 = l
                If t < y0 Then y0 = t
                If l + w > x1 Then x1 = l + w
                If t + h > y1 Then y1 = t + h
            End If
        End If
    Next c
    If x1 < x0 Then Exit Sub
    dx = (frm.ScaleWidth - (x1 - x0)) / 2 - x0
    dy = (frm.ScaleHeight - (y1 - y0)) / 2 - y0
    For Each c In frm.Controls
        If OnForm(frm, c) Then
            c.Left = c.Left + dx
            c.Top = c.Top + dy
        End If
    Next c
End Sub
