Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function ShellExecute _
                         Lib "shell32.dll" _
                         Alias "ShellExecuteA" ( _
                         ByVal hwnd As Long, _
                         ByVal lpOperation As String, _
                         ByVal lpFile As String, _
                         ByVal lpParameters As String, _
                         ByVal lpDirectory As String, _
                         ByVal nShowCmd As Long) _
                         As Long

Public Sub CheckKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_UP) >= 0 And GetAsyncKeyState(VK_LEFT) >= 0 Then DirUpLeft = False
    If GetAsyncKeyState(VK_UP) >= 0 And GetAsyncKeyState(VK_RIGHT) >= 0 Then DirUpRight = False
    If GetAsyncKeyState(VK_DOWN) >= 0 And GetAsyncKeyState(VK_LEFT) >= 0 Then DirDownLeft = False
    If GetAsyncKeyState(VK_DOWN) >= 0 And GetAsyncKeyState(VK_RIGHT) >= 0 Then DirDownRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If
    
    If GetKeyState(vbKeyTab) < 0 Then
        tabDown = True
    Else
        tabDown = False
    End If
    
    If Not chatOn Then
        If GetKeyState(vbKeySpace) < 0 Then
            CheckMapGetItem
        End If
        'Move Up Left

If GetAsyncKeyState(VK_UP) < 0 And GetAsyncKeyState(VK_LEFT) < 0 Then

     DirUp = False

     DirDown = False

     DirLeft = False

     DirRight = False

     DirUpLeft = True

     DirUpRight = False

     DirDownLeft = False

     DirDownRight = False

     Exit Sub

Else

     DirUpLeft = False

End If



'Move Up Right

If GetAsyncKeyState(VK_UP) < 0 And GetAsyncKeyState(VK_RIGHT) < 0 Then

     DirUp = False

     DirDown = False

     DirLeft = False

     DirRight = False

     DirUpLeft = False

     DirUpRight = True

     DirDownLeft = False

     DirDownRight = False

     Exit Sub

Else

     DirUpRight = False

End If



'Move Down Left

If GetAsyncKeyState(VK_DOWN) < 0 And GetAsyncKeyState(VK_LEFT) < 0 Then

     DirUp = False

     DirDown = False

     DirLeft = False

     DirRight = False

     DirUpLeft = False

     DirUpRight = False

     DirDownLeft = True

     DirDownRight = False

     Exit Sub

Else

     DirDownLeft = False

End If



'Move Down Right

If GetAsyncKeyState(VK_DOWN) < 0 And GetAsyncKeyState(VK_RIGHT) < 0 Then

     DirUp = False

     DirDown = False

     DirLeft = False

     DirRight = False

     DirUpLeft = False

     DirUpRight = False

     DirDownLeft = False

     DirDownRight = True

     Exit Sub

Else

     DirDownRight = False

End If



'Move Up

If GetAsyncKeyState(VK_UP) < 0 Then

     DirUp = True

     DirDown = False

     DirLeft = False

     DirRight = False

     DirUpLeft = False

     DirUpRight = False

     DirDownLeft = False

     DirDownRight = False

     Exit Sub

Else

     DirUp = False

End If



'Move Right

If GetAsyncKeyState(VK_RIGHT) < 0 Then

     DirUp = False

     DirDown = False

     DirLeft = False

     DirRight = True

     DirUpLeft = False

     DirUpRight = False

     DirDownLeft = False

     DirDownRight = False

     Exit Sub

Else

     DirRight = False

End If



'Move down

If GetAsyncKeyState(VK_DOWN) < 0 Then

     DirUp = False

     DirDown = True

     DirLeft = False

     DirRight = False

     DirUpLeft = False

     DirUpRight = False

     DirDownLeft = False

     DirDownRight = False

     Exit Sub

Else

     DirDown = False

End If



'Move left

If GetAsyncKeyState(VK_LEFT) < 0 Then

     DirUp = False

     DirDown = False

     DirLeft = True

     DirRight = False

     DirUpLeft = False

     DirUpRight = False

     DirDownLeft = False

     DirDownRight = False

     Exit Sub

Else

     DirLeft = False

End If

    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim chatText As String
Dim Name As String
Dim i As Long
Dim n As Long
Dim Command() As String
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If isLogging = True Then
        If NewGUIWindow(TEXTPASSWORD).visible Then
            If (KeyAscii = vbKeyBack) Then
                If LenB(NewGUIWindow(TEXTPASSWORD).value) > 0 Then NewGUIWindow(TEXTPASSWORD).value = Mid$(NewGUIWindow(TEXTPASSWORD).value, 1, Len(NewGUIWindow(TEXTPASSWORD).value) - 1)
            End If
                
            ' And if neither, then add the character to the user's text buffer
            If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) And Len(NewGUIWindow(TEXTPASSWORD).value) < 20 Then
                NewGUIWindow(TEXTPASSWORD).value = NewGUIWindow(TEXTPASSWORD).value & ChrW$(KeyAscii)
            End If
            
            If (KeyAscii = vbKeyTab) Then
                NewGUIWindow(TEXTPASSWORD).visible = False
            End If
        End If
        If NewGUIWindow(TEXTLOGIN).visible Then
            If (KeyAscii = vbKeyBack) Then
                If LenB(NewGUIWindow(TEXTLOGIN).value) > 0 Then NewGUIWindow(TEXTLOGIN).value = Mid$(NewGUIWindow(TEXTLOGIN).value, 1, Len(NewGUIWindow(TEXTLOGIN).value) - 1)
            End If
                
            ' And if neither, then add the character to the user's text buffer
            If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) And Len(NewGUIWindow(TEXTLOGIN).value) < 12 Then
                NewGUIWindow(TEXTLOGIN).value = NewGUIWindow(TEXTLOGIN).value & ChrW$(KeyAscii)
            End If
            
            If (KeyAscii = vbKeyTab) Then
                NewGUIWindow(TEXTPASSWORD).visible = True
                NewGUIWindow(TEXTLOGIN).visible = False
            End If
        End If
        If NewGUIWindow(TEXTCHARNAME).visible Then
            If (KeyAscii = vbKeyBack) Then
                If LenB(NewGUIWindow(TEXTCHARNAME).value) > 0 Then NewGUIWindow(TEXTCHARNAME).value = Mid$(NewGUIWindow(TEXTCHARNAME).value, 1, Len(NewGUIWindow(TEXTCHARNAME).value) - 1)
            End If
                
            ' And if neither, then add the character to the user's text buffer
            If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) And Len(NewGUIWindow(TEXTCHARNAME).value) < 12 Then
                NewGUIWindow(TEXTCHARNAME).value = NewGUIWindow(TEXTCHARNAME).value & ChrW$(KeyAscii)
            End If
        End If
        Exit Sub
    End If
    
    chatText = MyText
    
    If GUIWindow(GUI_CURRENCY).visible Then
        If (KeyAscii = vbKeyBack) Then
            If LenB(sDialogue) > 0 Then sDialogue = Mid$(sDialogue, 1, Len(sDialogue) - 1)
        End If
            
        ' And if neither, then add the character to the user's text buffer
        If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
            sDialogue = sDialogue & ChrW$(KeyAscii)
        End If
        Exit Sub
    End If
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
        chatOn = Not chatOn
        
        ' Broadcast message
        If Left$(chatText, 1) = "'" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call BroadcastMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Emote message
        If Left$(chatText, 1) = "-" Then
            MyText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call EmoteMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Player message
        If Left$(chatText, 1) = "!" Then
            If Mid$(chatText, 1, 2) = "! " Then GoTo Continue
            Name = vbNullString

            ' Get the desired player from the user text
            For i = 1 To Len(chatText)

                If Mid$(chatText, i, 1) <> Space(1) Then
                    Name = Name & Mid$(chatText, i, 1)
                Else
                    Exit For
                End If

            Next

            chatText = Mid$(chatText, i, Len(chatText) - 1)

            ' Make sure they are actually sending something
            If Len(chatText) - i > 0 Then
                chatText = Mid$(chatText, i + 1, Len(chatText) - i)
                ' Send the message to the player
                Call PlayerMsg(chatText, Name)
            Else
                Call AddText(printf("Use: !nomedojogador (mensagem)"), AlertColor)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case Command(0)
                Case "/ajuda"
                    Call AddText(printf("Comandos do jogo:"), HelpColor)
                    Call AddText(printf("'mensagemaqui = Envia a mensagem no chat global"), HelpColor)
                    Call AddText(printf("!nomedojogador mensagemaqui = Mensagem privada"), HelpColor)
                    Call AddText(printf("Comandos comuns: /gui (Desativa os menus do jogo), /trade (Envia uma solicitação de troca para outro jogador), /fps (Mostra os frames por segundo que o jogo está sendo renderizado), /reportar [Motivo] (Reporta o jogador marcado como alvo)"), HelpColor)
                    Set buffer = Nothing
                Case "/gui"
                    hideGUI = Not hideGUI
                Case "/fucktutorial"
                    SendCompleteTutorial
                    Player(MyIndex).InTutorial = 1
                    ' Whos Online
                Case "/fps"
                    BFPS = Not BFPS
                    ' Request stats
                    ' // Monitor Admin Commands //
                    ' Admin Help
                    ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    BLoc = Not BLoc
                    ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If

                    ' Setting sprite
                Case "/setsprite"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    SendSetSprite CLng(Command(1))
                    ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(chatText, Len(chatText) - 5)
                    ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendBanList
                    ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo Continue
                    End If

                    SendBan Command(1)
                    ' // Developer Admin Commands //
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                    ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    SendBanDestroy
                    ' Textures debug mode
                Case "/debug"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    DEBUG_MODE = Not DEBUG_MODE
                Case "/quests"
                    GUIWindow(GUI_QUESTS).visible = Not GUIWindow(GUI_QUESTS).visible
                Case "/trade"
                    If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                        SendTradeRequest
                    Else
                        AddText printf("Antes de fazer uma troca, clique em cima do jogador com quem você deseja efetuá-la."), BrightRed
                    End If
                Case "/reportar"
                    If myTarget > 0 Or myTargetType = TARGET_TYPE_PLAYER Then
                        Call SayMsg(chatText)
                    Else
                        Call AddText(printf("É necessário ter alguém como alvo para reportar!"), BrightRed)
                    End If
                Case "/sair"
                    Call SayMsg(chatText)
                Case "/suporte"
                    SendSupportMsg "open"
                Case Else
                    AddText printf("Comando inválido!"), HelpColor
            End Select

            'continue label where we go instead of exiting the sub
Continue:
            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Say message
        If Len(chatText) > 0 Then
            Call SayMsg(MyText)
        End If

        MyText = vbNullString
        UpdateShowChatText
        Exit Sub
    End If
    If Not chatOn Then Exit Sub
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
        UpdateShowChatText
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            MyText = MyText & ChrW$(KeyAscii)
            UpdateShowChatText
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub HandleMouseMove(ByVal X As Long, ByVal Y As Long, ByVal Button As Long)
Dim i As Long
    ' Set the global cursor position
    
    GlobalX = X
    GlobalY = Y
    GlobalX_Map = GlobalX + (TileView.Left * PIC_X) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.Top * PIC_Y) + Camera.Top
    ' GUI processing
    If Not hideGUI Then
        For i = 1 To Gui_Count - 1
            If (X >= GUIWindow(i).X And X <= GUIWindow(i).X + GUIWindow(i).Width) And (Y >= GUIWindow(i).Y And Y <= GUIWindow(i).Y + GUIWindow(i).Height) Then
                If GUIWindow(i).visible Then
                    Select Case i
                        Case GUI_CHAT, GUI_BARS, GUI_MENU
                            ' Put nothing here and we can click through them!
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
    
    CheckNewGuiMove
    
    ' Handle the events
    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((Y + Camera.Top) \ PIC_Y)
End Sub
Public Sub HandleMouseDown(ByVal Button As Long)
Dim i As Long

    If InGame Then CloseDaily = True

    ' GUI processing
    If Not hideGUI Then
        For i = 1 To Gui_Count - 1
            If i = GUI_MENU Then Menu_MouseDown Button
            If (GlobalX >= GUIWindow(i).X And GlobalX <= GUIWindow(i).X + GUIWindow(i).Width) And (GlobalY >= GUIWindow(i).Y And GlobalY <= GUIWindow(i).Y + GUIWindow(i).Height) Then
                If GUIWindow(i).visible Then
                    Select Case i
                        Case GUI_CHAT, GUI_BARS
                            ' Put nothing here and we can click through the
                        Case GUI_INVENTORY
                            Inventory_MouseDown Button
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_MouseDown Button
                            Exit Sub
                        Case GUI_MENU
                            'Menu_MouseDown Button
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_MouseDown Button
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_MouseDown
                            Exit Sub
                        Case GUI_CURRENCY
                            Currency_MouseDown
                            Exit Sub
                        Case GUI_DIALOGUE
                            Dialogue_MouseDown
                            Exit Sub
                        Case GUI_SHOP
                            Shop_MouseDown
                            Exit Sub
                        Case GUI_PARTY
                            Party_MouseDown
                            Exit Sub
                        Case GUI_OPTIONS
                            Options_MouseDown
                            Exit Sub
                        Case GUI_TRADE
                            Trade_MouseDown
                            Exit Sub
                        Case GUI_EVENTCHAT
                            Chat_MouseDown
                            Exit Sub
                        Case GUI_DEATH
                            Death_MouseDown
                            Exit Sub
                        Case GUI_CONQUISTAS
                            Conquistas_MouseDown Button
                            Exit Sub
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
        ' check chat buttons
        If Not inChat Then
            ChatScroll_MouseDown
        End If
        If UZ And VIAGEMMAP <> GetPlayerMap(MyIndex) And MatchActive = 0 And InTutorial = False Then
        If GlobalX >= 640 And GlobalX <= 800 Then
            If GlobalY >= 470 And GlobalY <= 520 Then
                If Buttons(6).visible = False Then frmFeedback.Show
            End If
        End If
        If GlobalX >= 580 And GlobalX <= 640 Then
            If GlobalY >= 470 And GlobalY <= 520 Then
                Dim r As Long
                r = ShellExecute(0, "open", "http://goplaygames.com.br/universoz/forum/viewtopic.php?f=12&t=11", 0, 0, 1)
            End If
        End If
        End If
    End If
    
    CheckNewGui
    ' left click
    If Button = vbLeftButton Then
        ' targetting
        FindTarget
    End If
End Sub

Public Sub HandleMouseUp(ByVal Button As Long)
Dim i As Long

    If UZ And GetPlayerMap(MyIndex) = VIAGEMMAP Then
        For i = 1 To MAX_PLANETS
            Dim X As Long, Y As Long
            X = (ConvertMapX(Planets(i).X * PIC_X)) - (Planets(i).Size / 2) + 16
            Y = (ConvertMapY(Planets(i).Y * PIC_Y)) - (Planets(i).Size / 2) + 16
        
            If GlobalX >= X And GlobalX <= X + Planets(i).Size Then
                If GlobalY >= Y And GlobalY <= Y + Planets(i).Size Then
                    If PlanetTarget = i Then
                        PlanetTarget = 0
                    Else
                        PlanetTarget = i
                    End If
                End If
            End If
        Next i
    End If
    
    If GetPlayerMap(MyIndex) = VirgoMap Then
        For i = 1 To MAX_PLAYER_PLANETS
            X = (ConvertMapX(PlayerPlanet(i).PlanetData.X * PIC_X)) - (PlayerPlanet(i).PlanetData.Size / 2) + 16
            Y = (ConvertMapY(PlayerPlanet(i).PlanetData.Y * PIC_Y)) - (PlayerPlanet(i).PlanetData.Size / 2) + 16
        
            If GlobalX >= X And GlobalX <= X + PlayerPlanet(i).PlanetData.Size Then
                If GlobalY >= Y And GlobalY <= Y + PlayerPlanet(i).PlanetData.Size Then
                    If PlanetTarget = i Then
                        PlanetTarget = 0
                    Else
                        PlanetTarget = i
                    End If
                End If
            End If
        Next i
    End If
    
    If InOwnPlanet And dialogueIndex = 0 Then
        For i = 55 To 60
            X = Buttons(i).X
            Y = Buttons(i).Y
            ' check if we're on the button
            If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
                If i = 55 And Buttons(55).visible = True Then
                    If EditTargetX >= 0 And EditTargetY >= 0 Then
                        If Map.Tile(EditTargetX, EditTargetY).Type = TILE_TYPE_BLOCKED Then SendRemoveBlock
                        If Map.Tile(EditTargetX, EditTargetY).Type = TILE_TYPE_RESOURCE Or Map.Tile(EditTargetX, EditTargetY).Type = TILE_TYPE_NPCSPAWN Then
                            Dialogue "Confirmação", "Você tem certeza que deseja remover esta construção?", DIALOGUE_TYPE_CONFIRM, True, 1
                            Exit Sub
                        End If
                        Dim n As Long
                        'For N = 1 To MapSaibamans(GetPlayerMap(MyIndex)).TotalSaibamans
                        '    If MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(N).Working = 1 Then
                        '        If EditTargetX = MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(N).X And EditTargetY = MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(N).Y Then
                        '            Dialogue "Confirmação", "Você tem certeza que deseja cancelar esta construção?", DIALOGUE_TYPE_CONFIRM, True, 1
                        '            Exit Sub
                        '        End If
                        '    End If
                        'Next N
                    End If
                End If
                If i = 56 And Buttons(56).visible = True Then
                    For n = 1 To MapSaibamans(GetPlayerMap(MyIndex)).TotalSaibamans
                        If MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(n).Working = 1 Then
                            If EditTargetX = MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(n).X And EditTargetY = MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(n).Y Then
                                SendAccelerate
                            End If
                        End If
                    Next n
                End If
                If i = 57 And Buttons(57).visible = True Then
                    If IsMovingObject = False Then
                        IsMovingObject = True
                        Exit Sub
                    End If
                End If
                If i = 58 And Buttons(58).visible = True Then
                    If EditTargetX > 0 And EditTargetY > 0 Then
                        SendEvolute
                    End If
                End If
                If i = 59 And Buttons(59).visible = True Then
                    If EditTargetX > 0 And EditTargetY > 0 Then
                        SendOpenBuilding
                        Exit Sub
                    End If
                End If
                If i = 60 And Buttons(60).visible = True Then
                    If EditTargetX > 0 And EditTargetY > 0 Then
                        SendAcelerar
                        Exit Sub
                    End If
                End If
            End If
            Buttons(i).visible = False
        Next i
        If IsMovingObject Then
            SendMoveBlock CurX, CurY
            IsMovingObject = False
        End If
        If Not frmMain.picFabrica.visible Then
            EditTargetX = CurX
            EditTargetY = CurY
        End If
    End If
    
    If InTutorial Then
        If GlobalX >= TutorialX + 300 And GlobalX <= TutorialX + 300 + 100 Then
            If GlobalY >= TutorialY + 240 And GlobalY <= TutorialY + 240 + 16 Then
                HandleTutorialClick
            End If
        End If
    End If
    
    'Pick menu
    Dim ButtonX As Long, ButtonY As Long
    ButtonX = GUIWindow(GUI_BARS).X + GUIWindow(GUI_BARS).Width - 96
    ButtonY = GUIWindow(GUI_BARS).Y + GUIWindow(GUI_BARS).Height - 24
    If GlobalX >= ButtonX And GlobalX <= ButtonX + 12 Then
        If GlobalY >= ButtonY And GlobalY <= ButtonY + 12 Then
            If Options.PickMenu = 0 Then
                Options.PickMenu = 1
            Else
                Options.PickMenu = 0
            End If
            SaveOptions
        End If
    End If

    ' GUI processing
    If Not hideGUI Then
        For i = 1 To Gui_Count - 1
            If i = GUI_MENU Then Menu_MouseUp
            If (GlobalX >= GUIWindow(i).X And GlobalX <= GUIWindow(i).X + GUIWindow(i).Width) And (GlobalY >= GUIWindow(i).Y And GlobalY <= GUIWindow(i).Y + GUIWindow(i).Height) Then
                If GUIWindow(i).visible Then
                    Select Case i
                        Case GUI_CHAT, GUI_BARS
                            ' Put nothing here and we can click through the
                        Case GUI_INVENTORY
                            Inventory_MouseUp
                        Case GUI_SPELLS
                            Spells_MouseUp
                        Case GUI_MENU
                            'Menu_MouseUp
                        Case GUI_HOTBAR
                            Hotbar_MouseUp
                        Case GUI_CHARACTER
                            Character_MouseUp
                        Case GUI_CURRENCY
                            Currency_MouseUp
                        Case GUI_DIALOGUE
                            Dialogue_MouseUp
                        Case GUI_SHOP
                            Shop_MouseUp
                        Case GUI_PARTY
                            Party_MouseUp
                        Case GUI_OPTIONS
                            Options_MouseUp
                        Case GUI_TRADE
                            Trade_MouseUp
                        Case GUI_EVENTCHAT
                            Chat_MouseUp
                        Case GUI_CONQUISTAS
                            Conquistas_MouseUp
                    End Select
                End If
            End If
        Next
    End If

    ' Stop dragging if we haven't catched it already
    DragInvSlotNum = 0
    DragBankSlotNum = 0
    DragSpell = 0
    ' reset buttons
    resetClickedButtons
    ' stop scrolling chat
    ChatButtonUp = False
    ChatButtonDown = False
End Sub

Public Sub HandleDoubleClick()
Dim i As Long

    ' GUI processing
    If Not hideGUI Then
        For i = 1 To Gui_Count - 1
            If (GlobalX >= GUIWindow(i).X And GlobalX <= GUIWindow(i).X + GUIWindow(i).Width) And (GlobalY >= GUIWindow(i).Y And GlobalY <= GUIWindow(i).Y + GUIWindow(i).Height) Then
                If GUIWindow(i).visible Then
                    Select Case i
                        Case GUI_INVENTORY
                            Inventory_DoubleClick
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_DoubleClick
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_DoubleClick
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_DoubleClick
                            Exit Sub
                        Case GUI_SHOP
                            Shop_DoubleClick
                            Exit Sub
                        Case GUI_BANK
                            Bank_DoubleClick
                            Exit Sub
                        Case GUI_TRADE
                            Trade_DoubleClick
                            Exit Sub
                        Case GUI_QUESTS
                            Quests_DoubleClick
                            Exit Sub
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
End Sub

Public Sub OpenGuiWindow(ByVal Index As Long)
Dim buffer As clsBuffer
    If Index = 1 Then
        GUIWindow(GUI_INVENTORY).visible = Not GUIWindow(GUI_INVENTORY).visible
    Else
        GUIWindow(GUI_INVENTORY).visible = False
    End If
    
    If Index = 2 Then
        GUIWindow(GUI_SPELLS).visible = Not GUIWindow(GUI_SPELLS).visible
        ' Update the spells on the pic
        Set buffer = New clsBuffer
        buffer.WriteLong CSpells
        SendData buffer.ToArray()
        Set buffer = Nothing
    Else
        GUIWindow(GUI_SPELLS).visible = False
    End If
    
    If Index = 3 Then
        GUIWindow(GUI_CHARACTER).visible = Not GUIWindow(GUI_CHARACTER).visible
    Else
        GUIWindow(GUI_CHARACTER).visible = False
    End If
    
    If Index = 4 Then
        GUIWindow(GUI_OPTIONS).visible = Not GUIWindow(GUI_OPTIONS).visible
    Else
        GUIWindow(GUI_OPTIONS).visible = False
    End If
    
    If Index = 6 Then
        GUIWindow(GUI_PARTY).visible = Not GUIWindow(GUI_PARTY).visible
    Else
        GUIWindow(GUI_PARTY).visible = False
    End If
    
    If Index = 5 Then
        GUIWindow(GUI_QUESTS).visible = Not GUIWindow(GUI_QUESTS).visible
    Else
        GUIWindow(GUI_QUESTS).visible = False
    End If
    
    If Index = 7 Then
        GUIWindow(GUI_CONQUISTAS).visible = Not GUIWindow(GUI_CONQUISTAS).visible
    Else
        GUIWindow(GUI_CONQUISTAS).visible = False
    End If
End Sub

Public Sub Currency_MouseDown()
Dim i As Long, X As Long, Y As Long, Width As Long
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        CurrencyAcceptState = 2 ' clicked
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        CurrencyCloseState = 2 ' clicked
    End If
End Sub
Public Sub Currency_MouseUp()
Dim i As Long, X As Long, Y As Long, Width As Long, buffer As clsBuffer
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        If CurrencyAcceptState = 2 Then
            ' do stuffs
            If IsNumeric(sDialogue) Or (CurrencyMenu = 7) Then
                Select Case CurrencyMenu
                    Case 1 ' drop item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        SendDropItem tmpCurrencyItem, Val(sDialogue)
                    Case 2 ' deposit item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        DepositItem tmpCurrencyItem, Val(sDialogue)
                    Case 3 ' withdraw item
                        If Val(sDialogue) > GetBankItemValue(tmpCurrencyItem) Then sDialogue = GetBankItemValue(tmpCurrencyItem)
                        WithdrawItem tmpCurrencyItem, Val(sDialogue)
                    Case 4 ' offer trade item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        TradeItem tmpCurrencyItem, Val(sDialogue)
                    Case 5
                        SelectedGravity = Val(sDialogue)
                        OpenCurrency 6, "Quantidade de horas que deseja ficar na sala (1h - 6h):"
                        CurrencyAcceptState = 0
                        CurrencyCloseState = 0
                        MyText = ""
                        Exit Sub
                    Case 6
                        SelectedHours = Val(sDialogue)
                        SendEnterGravity
                    Case 7
                        SendPlanetName
                End Select
                MyText = ""
            Else
                AddText printf("Por favor, coloque uma quantidade válida"), BrightRed
                Exit Sub
            End If
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    ' check if we're on the button
    If (GlobalX >= X And GlobalX <= X + Buttons(12).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(12).Height) Then
        If CurrencyCloseState = 2 Then
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    
    CurrencyAcceptState = 0
    CurrencyCloseState = 0
    GUIWindow(GUI_CURRENCY).visible = False
    inChat = False
    chatOn = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    sDialogue = vbNullString
    ' reset buttons
    resetClickedButtons
End Sub
Public Sub Dialogue_MouseDown()
Dim i As Long, X As Long, Y As Long, Width As Long
    
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 90
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(1) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(2) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 120
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(3) = 2 ' clicked
        End If
    End If
End Sub

Public Sub Dialogue_MouseUp()
Dim i As Long, X As Long, Y As Long, Width As Long
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        X = GUIWindow(GUI_CHAT).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 90
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(1) = 2 Then
                Dialogue_Button_MouseDown (2)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(1) = 0
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        X = GUIWindow(GUI_CHAT).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(2) = 2 Then
                Dialogue_Button_MouseDown (1)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(2) = 0
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        X = GUIWindow(GUI_CHAT).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 120
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(3) = 2 Then
                Dialogue_Button_MouseDown (3)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(3) = 0
    End If
End Sub

' scroll bar
Public Sub ChatScroll_MouseDown()
Dim i As Long, X As Long, Y As Long, Width As Long
    
    ' find out which button we're clicking
    For i = 34 To 35
        X = GUIWindow(GUI_CHAT).X + Buttons(i).X
        Y = GUIWindow(GUI_CHAT).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            Buttons(i).State = 2 ' clicked
            ' scroll the actual chat
            Select Case i
                Case 34 ' up
                    'ChatScroll = ChatScroll + 1
                    ChatButtonUp = True
                Case 35 ' down
                    'ChatScroll = ChatScroll - 1
                    'If ChatScroll < 8 Then ChatScroll = 8
                    ChatButtonDown = True
            End Select
        End If
    Next
End Sub

' Shop
Public Sub Shop_MouseUp()
Dim i As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For i = 23 To 23
        X = GUIWindow(GUI_SHOP).X + Buttons(i).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            'If Buttons(i).State = 2 Then
                ' do stuffs
                Select Case i
                    Case 23
                        ' exit
                        Set buffer = New clsBuffer
                        buffer.WriteLong CCloseShop
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        GUIWindow(GUI_SHOP).visible = False
                        InShop = 0
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Shop_MouseDown()
Dim i As Long, X As Long, Y As Long

    ' find out which button we're clicking
    For i = 23 To 23
        X = GUIWindow(GUI_SHOP).X + Buttons(i).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            Buttons(i).State = 2 ' clicked
        End If
    Next
End Sub

Public Sub Shop_DoubleClick()
Dim shopSlot As Long

    shopSlot = IsShopItem(GlobalX, GlobalY)

    If shopSlot > 0 Then
        ' buy item code
        BuyItem shopSlot
    End If
End Sub
Public Sub Bank_DoubleClick()
Dim bankNum As Long
    bankNum = IsBankItem(GlobalX, GlobalY)
    If bankNum <> 0 Then
        If GetBankItemNum(bankNum) > 0 Then
        'If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetBankItemNum(bankNum)).Stackable > 0 Then
            CurrencyMenu = 3 ' withdraw
            CurrencyText = "Quanto você gostaria de retirar?"
            tmpCurrencyItem = bankNum
            sDialogue = vbNullString
            GUIWindow(GUI_CURRENCY).visible = True
            inChat = True
            chatOn = True
            Exit Sub
        End If
        WithdrawItem bankNum, 0
        Exit Sub
        End If
    End If
End Sub
Public Sub Trade_DoubleClick()
Dim tradeNum As Long
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum <> 0 Then
        UntradeItem tradeNum
        Exit Sub
    End If
End Sub
Public Sub Trade_MouseDown()
Dim i As Long, X As Long, Y As Long

    ' find out which button we're clicking
    For i = 40 To 41
        X = Buttons(i).X
        Y = Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            Buttons(i).State = 2 ' clicked
        End If
    Next
End Sub
Public Sub Trade_MouseUp()
Dim i As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For i = 40 To 41
        X = Buttons(i).X
        Y = Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            'If Buttons(i).State = 2 Then
                ' do stuffs
                Select Case i
                    Case 40
                        AcceptTrade
                    Case 41
                        DeclineTrade
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

' Party
Public Sub Party_MouseUp()
Dim i As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For i = 24 To 25
        X = GUIWindow(GUI_PARTY).X + Buttons(i).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            'If Buttons(i).State = 2 Then
                ' do stuffs
                Select Case i
                    Case 24 ' invite
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendPartyRequest
                        Else
                            AddText printf("Alvo para convite inválido."), BrightRed
                        End If
                    Case 25 ' leave
                        If Party.Leader > 0 Then
                            SendPartyLeave
                        Else
                            AddText printf("Você não está em uma party."), BrightRed
                        End If
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Party_MouseDown()
Dim i As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For i = 24 To 25
        X = GUIWindow(GUI_PARTY).X + Buttons(i).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            Buttons(i).State = 2 ' clicked
        End If
    Next
End Sub

'Options
Public Sub Options_MouseUp()
Dim i As Long, X As Long, Y As Long, buffer As clsBuffer, layerNum As Long

    ' find out which button we're clicking
    For i = 26 To 31
        X = GUIWindow(GUI_OPTIONS).X + Buttons(i).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            If Buttons(i).State = 3 Then
                ' do stuffs
                Select Case i
                    Case 26 ' music on
                        Options.Music = 1
                        PlayMusic Trim$(Map.Music)
                        SaveOptions
                        Buttons(26).State = 2
                        Buttons(27).State = 0
                    Case 27 ' music off
                        Options.Music = 0
                        StopMusic
                        SaveOptions
                        Buttons(26).State = 0
                        Buttons(27).State = 2
                    Case 28 ' sound on
                        Options.Sound = 1
                        SaveOptions
                        Buttons(28).State = 2
                        Buttons(29).State = 0
                    Case 29 ' sound off
                        Options.Sound = 0
                        StopAllSounds
                        SaveOptions
                        Buttons(28).State = 0
                        Buttons(29).State = 2
                    Case 30 ' debug on
                        Options.Debug = 1
                        SaveOptions
                        Buttons(30).State = 2
                        Buttons(31).State = 0
                    Case 31 ' debug off
                        Options.Debug = 0
                        SaveOptions
                        Buttons(30).State = 0
                        Buttons(31).State = 2
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    For i = 42 To 45
    ' set co-ordinate
        X = GUIWindow(GUI_OPTIONS).X + Buttons(i).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(i).Y
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            If Buttons(i).State = 2 Then
                Select Case i
                    Case 42
                        If Options.FPS = 15 Then Options.FPS = 20
                        SaveOptions
                    Case 43
                        If Options.FPS = 20 Then Options.FPS = 15
                        SaveOptions
                    Case 44
                        If Options.volume - 10 >= 0 Then
                            Options.volume = Options.volume - 10
                            StopMusic
                            PlayMusic Trim$(Map.Music)
                        Else
                            Options.volume = 0
                        End If
                        SaveOptions
                    Case 45
                        If Options.volume + 10 <= 150 Then
                            Options.volume = Options.volume + 10
                            StopMusic
                            PlayMusic Trim$(Map.Music)
                        Else
                            Options.volume = 150
                        End If
                        SaveOptions
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    For i = 46 To 53
        X = GUIWindow(GUI_OPTIONS).X + Buttons(i).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            If Buttons(i).State = 3 Then
                ' do stuffs
                Select Case i
                    Case 46 ' ambiente on
                        Options.Ambiente = 1
                        SaveOptions
                        Buttons(46).State = 2
                        Buttons(47).State = 0
                    Case 47 ' ambiente off
                        Options.Ambiente = 0
                        SaveOptions
                        Buttons(46).State = 0
                        Buttons(47).State = 2
                    Case 48 ' tela on
                        Options.Tela = 1
                        SaveOptions
                        Buttons(48).State = 2
                        Buttons(49).State = 0
                    Case 49 ' tela off
                        Options.Tela = 0
                        StopAllSounds
                        SaveOptions
                        Buttons(48).State = 0
                        Buttons(49).State = 2
                    Case 50 ' clima on
                        Options.Clima = 1
                        SaveOptions
                        Buttons(50).State = 2
                        Buttons(51).State = 0
                    Case 51 ' clima off
                        Options.Clima = 0
                        SaveOptions
                        Buttons(50).State = 0
                        Buttons(51).State = 2
                    Case 52 ' neblina on
                        Options.Neblina = 1
                        SaveOptions
                        Buttons(52).State = 2
                        Buttons(53).State = 0
                    Case 53 ' neblina off
                        Options.Neblina = 0
                        SaveOptions
                        Buttons(52).State = 0
                        Buttons(3).State = 2
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Options_MouseDown()
Dim i As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For i = 26 To 31
        X = GUIWindow(GUI_OPTIONS).X + Buttons(i).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            If Buttons(i).State = 0 Then
                Buttons(i).State = 3 ' clicked
            End If
        End If
    Next
    For i = 42 To 45
    ' set co-ordinate
        X = GUIWindow(GUI_OPTIONS).X + Buttons(i).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(i).Y
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            Buttons(i).State = 2 ' clicked
        End If
    Next
    For i = 46 To 53
        X = GUIWindow(GUI_OPTIONS).X + Buttons(i).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            If Buttons(i).State = 0 Then
                Buttons(i).State = 3 ' clicked
            End If
        End If
    Next
End Sub


' Menu
Public Sub Menu_MouseUp()
Dim i As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For i = 1 To 8
        Dim ButtonIndex As Long
        If i <= 6 Then
            ButtonIndex = i
        Else
            ButtonIndex = 60 - 6 + i
        End If
        X = GUIWindow(GUI_MENU).X + Buttons(ButtonIndex).X
        Y = GUIWindow(GUI_MENU).Y + Buttons(ButtonIndex).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(ButtonIndex).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(ButtonIndex).Height) Then
            If Buttons(ButtonIndex).State = 2 Then
                ' do stuffs
                Select Case i
                    Case 1
                        ' open window
                        OpenGuiWindow 1
                    Case 2
                        ' open window
                        OpenGuiWindow 2
                    Case 3
                        ' open window
                        OpenGuiWindow 3
                    Case 4
                        ' open window
                        OpenGuiWindow 4
                    Case 5
                        OpenGuiWindow 5
                    Case 6
                        ' open window
                        OpenGuiWindow 6
                    Case 7
                        Buttons(6).visible = Not Buttons(6).visible
                        Buttons(62).visible = Not Buttons(62).visible
                    Case 8
                        OpenGuiWindow 7
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Menu_MouseDown(ByVal Button As Long)
Dim i As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For i = 1 To 8
        Dim ButtonIndex As Long
        If i <= 6 Then
            ButtonIndex = i
        Else
            ButtonIndex = 60 - 6 + i
        End If
        If Buttons(ButtonIndex).visible Then
            X = GUIWindow(GUI_MENU).X + Buttons(ButtonIndex).X
            Y = GUIWindow(GUI_MENU).Y + Buttons(ButtonIndex).Y
            ' check if we're on the button
            If (GlobalX >= X And GlobalX <= X + Buttons(ButtonIndex).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(ButtonIndex).Height) Then
                Buttons(ButtonIndex).State = 2 ' clicked
            End If
        End If
    Next
End Sub

'Conquistas
Public Sub Conquistas_MouseDown(ByVal Button As Long)
Dim i As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For i = 63 To 64
        X = Buttons(i).X
        Y = Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            Buttons(i).State = 2 ' clicked
        End If
    Next
End Sub

Public Sub Conquistas_MouseUp()
Dim i As Long, X As Long, Y As Long

    ' find out which button we're clicking
    For i = 63 To 64
        X = Buttons(i).X
        Y = Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            'If Buttons(i).State = 2 Then
                Select Case i
                    Case 63
                        If (PageNum + 1) * 3 < UBound(Conquistas) Then
                            PageNum = PageNum + 1
                        End If
                    Case 64
                        If PageNum - 1 >= 0 Then
                            PageNum = PageNum - 1
                        End If
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        End If
    Next i
    
    resetClickedButtons
End Sub

' Inventory
Public Sub Inventory_MouseUp()
Dim invSlot As Long
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        invSlot = IsInvItem(GlobalX, GlobalY, True)
        If invSlot = 0 Then Exit Sub
        ' change slots
        SendChangeInvSlots DragInvSlotNum, invSlot
    End If

    DragInvSlotNum = 0
End Sub

Public Sub Inventory_MouseDown(ByVal Button As Long)
Dim invNum As Long

    invNum = IsInvItem(GlobalX, GlobalY)

    If Button = 1 Then
        If invNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = invNum
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If invNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                    If GetPlayerInvItemValue(MyIndex, invNum) > 0 Then
                        'CurrencyMenu = 1 ' drop
                        OpenCurrency 1, "Quanto você quer dropar?"
                        'CurrencyText = "Quanto você quer dropar?"
                        tmpCurrencyItem = invNum
                        'sDialogue = vbNullString
                        'GUIWindow(GUI_CURRENCY).visible = True
                        'inChat = True
                        'chatOn = True
                    End If
                Else
                    Call SendDropItem(invNum, 0)
                End If
            End If
        End If
    End If
End Sub

Public Sub Inventory_DoubleClick()
    Dim invNum As Long, value As Long, multiplier As Double, i As Long

    DragInvSlotNum = 0
    invNum = IsInvItem(GlobalX, GlobalY)

    If invNum > 0 Then
        ' are we in a shop?
        If InShop > 0 Then
            SellItem invNum
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                CurrencyMenu = 2 ' deposit
                CurrencyText = "Quanto voce quer depositar?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).visible = True
                inChat = True
                chatOn = True
                Exit Sub
            End If
                
            Call DepositItem(invNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For i = 1 To MAX_INV
                If TradeYourOffer(i).num = invNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Stackable > 0 Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                CurrencyMenu = 4 ' offer in trade
                CurrencyText = "Quanto voce quer trocar?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).visible = True
                inChat = True
                chatOn = True
                Exit Sub
            End If
            
            Call TradeItem(invNum, 0)
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_SCOUTER Then
            PlaySound "ligando scouter.mp3", -1, -1
            ScouterOn = Not ScouterOn
            myTarget = 0
            myTargetType = 0
            Exit Sub
        End If
        
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_RADAR Then
            PlaySound "ligando scouter.mp3", -1, -1
            RadarActive = Not RadarActive
            Exit Sub
        End If
        
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_PLANETCHANGE And Item(GetPlayerInvItemNum(MyIndex, invNum)).data1 = 0 Then
            'Dialogue "Nomear planeta", "Digite o novo nome do seu planeta:", DIALOGUE_TYPE_NAMEPLANET
            OpenCurrency 7, "Digite o novo nome do seu planeta:"
            Exit Sub
        End If
        
        Call SendUseItem(invNum)
        Exit Sub
    End If
End Sub

'Quests
Public Sub Quests_DoubleClick()
Dim QuestNum As Long
    
    QuestNum = IsPlayerQuest(GlobalX, GlobalY)
    
    If QuestNum <> 0 Then
        Call SendQuestInfo(QuestNum)
        Exit Sub
    End If
End Sub

' Spells
Public Sub Spells_DoubleClick()
Dim SpellNum As Long

    SpellNum = IsPlayerSpell(GlobalX, GlobalY)

    If SpellNum <> 0 Then
        Call CastSpell(SpellNum)
        Exit Sub
    End If
    
    SpellNum = IsPlayerEvoluteSpell(GlobalX, GlobalY)

    If SpellNum <> 0 And Not SpellNum = MAX_SPELLS Then
        Dim buffer As clsBuffer
        Set buffer = New clsBuffer
            buffer.WriteLong CUpgrade
            buffer.WriteLong SpellNum
            SendData buffer.ToArray()
        Set buffer = Nothing
        Exit Sub
    End If
End Sub

Public Sub Spells_MouseDown(ByVal Button As Long)
Dim SpellNum As Long

    SpellNum = IsPlayerSpell(GlobalX, GlobalY)
    If Button = 1 Then ' left click
        If SpellNum <> 0 Then
            DragSpell = SpellNum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If SpellNum <> 0 Then
            If PlayerSpells(SpellNum) > 0 Then
                'Dialogue "Forget Spell", "Você tem certeza que deseja esquecer a habilidade " & Trim$(Spell(PlayerSpells(spellnum)).Name) & "?", DIALOGUE_TYPE_FORGET, True, spellnum
            End If
        End If
    End If
End Sub

Public Sub Spells_MouseUp()
Dim spellSlot As Long

    If DragSpell > 0 Then
        spellSlot = IsPlayerSpell(GlobalX, GlobalY, True)
        If spellSlot = 0 Then Exit Sub
        SendChangeSpellSlots DragSpell, spellSlot
    End If

    DragSpell = 0
End Sub

' character
Public Sub Character_DoubleClick()
Dim eqNum As Long

    eqNum = IsEqItem(GlobalX, GlobalY)

    If eqNum <> 0 Then
        SendUnequip eqNum
    End If
End Sub
' hotbar
Public Sub Hotbar_DoubleClick()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarUse slotNum
    End If
End Sub

Public Sub Hotbar_MouseDown(ByVal Button As Long)
Dim slotNum As Long
    
    If Button <> 2 Then Exit Sub ' right click
    
    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarChange 0, 0, slotNum
    End If
End Sub

Public Sub Hotbar_MouseUp()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum = 0 Then Exit Sub
    
    ' inventory
    If DragInvSlotNum > 0 Then
        SendHotbarChange 1, DragInvSlotNum, slotNum
        DragInvSlotNum = 0
        Exit Sub
    End If
    
    ' spells
    If DragSpell > 0 Then
        SendHotbarChange 2, DragSpell, slotNum
        DragSpell = 0
        Exit Sub
    End If
End Sub
Public Sub Dialogue_Button_MouseDown(Index As Integer)
    ' call the handler
    dialogueHandler Index
    GUIWindow(GUI_DIALOGUE).visible = False
    inChat = False
    dialogueIndex = 0
End Sub
Public Sub Character_MouseDown()
Dim i As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For i = 16 To 20
        X = GUIWindow(GUI_CHARACTER).X + Buttons(i).X
        Y = GUIWindow(GUI_CHARACTER).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            Buttons(i).State = 2 ' clicked
        End If
    Next
End Sub

Public Sub Character_MouseUp()
Dim i As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For i = 16 To 20
        X = GUIWindow(GUI_CHARACTER).X + Buttons(i).X
        Y = GUIWindow(GUI_CHARACTER).Y + Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            ' send the level up
            If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
            SendTrainStat (i - 15)
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    Next
End Sub
' Npc Chat
Public Sub Chat_MouseDown()
Dim i As Long, X As Long, Y As Long, Width As Long

Select Case CurrentEvent.Type
    Case Evt_Menu
    For i = 1 To UBound(CurrentEvent.Text) - 1
        If Len(Trim$(CurrentEvent.Text(i + 1))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(CurrentEvent.Text(i + 1)) & "]")
            X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
            Y = GUIWindow(GUI_EVENTCHAT).Y + 115 - ((i - 1) * 15)
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                chatOptState(i) = 2 ' clicked
            End If
        End If
    Next
    Case Evt_Message
    Width = EngineGetTextWidth(Font_Default, "[Continue]")
    X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
    Y = GUIWindow(GUI_EVENTCHAT).Y + 100
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        chatContinueState = 2 ' clicked
    End If
End Select

End Sub
Public Sub Chat_MouseUp()
Dim i As Long, X As Long, Y As Long, Width As Long

Select Case CurrentEvent.Type
    Case Evt_Menu
        For i = 1 To UBound(CurrentEvent.Text) - 1
            If Len(Trim$(CurrentEvent.Text(i + 1))) > 0 Then
                Width = EngineGetTextWidth(Font_Default, "[" & Trim$(CurrentEvent.Text(i + 1)) & "]")
                X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
                Y = GUIWindow(GUI_EVENTCHAT).Y + 115 - ((i - 1) * 15)
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' are we clicked?
                    If chatOptState(i) = 2 Then
                        Events_SendChooseEventOption CurrentEvent.data(i)
                        ' play sound
                        PlaySound Sound_ButtonClick, -1, -1
                    End If
                End If
            End If
        Next
        
        For i = 1 To UBound(CurrentEvent.Text) - 1
            chatOptState(i) = 0 ' normal
        Next
    Case Evt_Message
        Width = EngineGetTextWidth(Font_Default, "[Continue]")
        X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
        Y = GUIWindow(GUI_EVENTCHAT).Y + 100
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' are we clicked?
            If chatContinueState = 2 Then
                Events_SendChooseEventOption CurrentEventIndex + 1
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        
        chatContinueState = 0
End Select
End Sub
Public Sub HandleKeyUp(ByVal KeyCode As Long)
Dim i As Long, buffer As clsBuffer
    
    ' hotbar
    If Not chatOn Then
        For i = 1 To 9
            If KeyCode = 48 + i Then
                SendHotbarUse i
            End If
        Next
        If KeyCode = 48 Then ' 0
            SendHotbarUse 10
        ElseIf KeyCode = 189 Then ' -
            SendHotbarUse 11
        ElseIf KeyCode = 187 Then ' =
            SendHotbarUse 12
        End If
    End If
    
    If InGame Then
        If KeyCode = vbKeyControl Then
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_RESOURCE Then
                If Resource(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).data1).ResourceType = 3 Then
                    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                        If Item(GetPlayerEquipment(MyIndex, Weapon)).data3 = 2 Then
                            If FishingTime < GetTickCount And FishingTime > GetTickCount - 500 And AlertX < GetTickCount Then
                                Set buffer = New clsBuffer
                                buffer.WriteLong CAttack
                                SendData buffer.ToArray()
                                Set buffer = Nothing
                                FishingTime = 0
                                Else
                                AlertX = GetTickCount + 500
                                Call PlaySound("reel.mp3", GetPlayerX(MyIndex), GetPlayerY(MyIndex))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = vbKeyF1 And isLogging = True Then
        DrawMousePosition = Not DrawMousePosition
    End If
    
    If KeyCode = vbKeyF1 Then
        If InGame Then
            If Player(MyIndex).Guild > 0 Then
                ShowGuildPanel
            End If
        End If
    End If
    
    If KeyCode = vbKeyF2 Then
        If GetPlayerAccess(MyIndex) > 0 Then
             frmAdminPanel.Show
        End If
    End If
    
    If isLogging = True Then
        If KeyCode = vbKeyF2 Then
            If ConnectToServer(1) Then
                Call SendLogin(GetVar(App.Path & "\data files\config.ini", "Options", "Username"), GetVar(App.Path & "\data files\config.ini", "Options", "Password"))
                AutoLogin = True
            End If
        End If
        If KeyCode = 39 Then
            SelectedServer = SelectedServer + 1
            If SelectedServer > UBound(Options.Servers) Then SelectedServer = 1
            Call TcpInit(False)
        End If
        If KeyCode = 37 Then
            SelectedServer = SelectedServer - 1
            If SelectedServer < 1 Then SelectedServer = UBound(Options.Servers)
            Call TcpInit(False)
        End If
    End If
    
End Sub

Sub Death_MouseDown()
    Dim X As Long, Y As Long, i As Long
        Y = (frmMain.ScaleHeight / 2) - (GUIWindow(GUI_DEATH).Height / 2)
        X = (frmMain.ScaleWidth / 2) - (GUIWindow(GUI_DEATH).Width / 2)
        
        If GlobalX >= Buttons(54).X + X And GlobalX <= Buttons(54).X + Buttons(54).Width + X Then
            If GlobalY >= Buttons(54).Y + Y And GlobalY <= Buttons(54).Y + Buttons(54).Height + Y Then
                SendOnDeath
            End If
        End If
End Sub

Sub OpenCurrency(Menu As Long, Text As String)
    CurrencyMenu = Menu
    CurrencyText = Text
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = True
    inChat = True
    chatOn = True
End Sub
