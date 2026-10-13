# RestaUm
Jogo resta 1, contém apenas 13 linhas de código :)


``` 
Private Sub imgGoti_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

    If (imgGoti(Index).Picture <> 0) _
        Or Abs(Source.Index - Index) <> 2 _
       And Abs(Source.Index - Index) <> 20 Then
       Exit Sub
    End If
    
    If (imgGoti(Source.Index - ((Source.Index - Index) / 2)).Picture = 0) Then
        Exit Sub
    End If
    
    Source.Picture = LoadPicture
    imgGoti(Source.Index - ((Source.Index - Index) / 2)).Picture = LoadPicture
    imgGoti(Index).Picture = imgGotipic.Picture

End Sub
```


![Capa](/capture.png?raw=true "Capa")
