Attribute VB_Name = "modEditor"
Public Enum Editor
    Name = 0
    Map
    X
    Y
    Login
    Senha
    Sprite
    Acesso
    Level
    Experiencia
    HP
    MP
    PDL
    Força
    Constituiçao
    KI
    Destreza
    Tecnica
    Pontos
    Bonus
    Tempo
    LevelVIP
    
    TotalTexts
End Enum

Sub AddToList(ByVal Name As String)
    Dim i As Long, n As Long
    
    Do While GetVar(App.Path & "\recents.ini", "RECENTS", Val(i)) <> ""
        i = i + 1
    Loop
    
    For n = 1 To i
        If LCase(GetVar(App.Path & "\recents.ini", "RECENTS", Val(n))) = LCase(Name) Then
            Exit Sub
        End If
    Next n

    Call PutVar(App.Path & "\recents.ini", "RECENTS", Val(i), Name)
End Sub

Sub LoadList()
    Dim i As Long, n As Long
    
    frmEditor.txtFile.Clear
    
    Do While GetVar(App.Path & "\recents.ini", "RECENTS", Val(i)) <> ""
        i = i + 1
    Loop
    
    For n = 1 To i
        frmEditor.txtFile.AddItem GetVar(App.Path & "\recents.ini", "RECENTS", Val(n))
    Next n
End Sub
