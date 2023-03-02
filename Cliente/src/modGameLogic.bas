Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
Dim FrameTime As Long
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim i As Long
Dim WalkTimer As Long
Dim tmr25 As Long
Dim tmr100 As Long
Dim tmr10000 As Long
Dim tmr500, Fadetmr As Long
Dim fogtmr As Long, bartmr As Long, targettmr As Long
Dim chatTmr As Long, surftmr As Long, rendertmr As Long
Dim AntiHackTmr As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    If Options.Debug = 2 Then On Error Resume Next

    If UZ Then PlayMusic Trim$(Map.Music)


    ' *** Start GameLoop ***
1    Do While InGame
        Tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.
        
        If AntiHackTmr < Tick Then
            VerificarAntiHack False
            AntiHackTmr = Tick + 1000
        End If
        
        If Not GetPlayerMap(MyIndex) = VIAGEMMAP And UZ Then RadarActive = False
        
        ' * Check surface timers *
        If surftmr < Tick Then
            For i = 1 To NumTextures
                UnsetTexture (i)
            Next
            surftmr = GetTickCount + 75000
        End If
        
        If Tremor > GetTickCount Then TremorX = Rand(-2, 2)
    
2        If tmr10000 < Tick Then
            ' check ping
            Call GetPing
            tmr10000 = Tick + 10000
            If GetPlayerX(MyIndex) >= 0 And GetPlayerX(MyIndex) <= Map.MaxX Then
            If GetPlayerY(MyIndex) >= 0 And GetPlayerY(MyIndex) <= Map.MaxY Then
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_RESOURCE Then
                If Resource(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).data1).ResourceType = 3 Then
                    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                        If Item(GetPlayerEquipment(MyIndex, Weapon)).data3 = 2 Then
                            If Not isFishing Then isFishing = True
                        End If
                    End If
                End If
            Else
                isFishing = False
            End If
            End If
            End If
        End If

        If tmr25 < Tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hwnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if we need to end the CD icon
3            If NumSpellIcons > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i) > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (Spell(PlayerSpells(i)).CDTime * 1000) < Tick Then
                                SpellCD(i) = 0
                            End If
                        End If
                    End If
                Next
            End If
            
            ' check if we need to unlock the player's spell casting restriction
            'If TempPlayer(MyIndex).SpellBuffer > 0 Then
            '    If TempPlayer(MyIndex).SpellBufferTimer + (Spell(PlayerSpells(TempPlayer(MyIndex).SpellBuffer)).CastTime * 1000) < Tick Then
                    'TempPlayer(MyIndex).SpellBuffer = 0
                    'TempPlayer(MyIndex).SpellBufferTimer = 0
            '    End If
            'End If
            
4            If GUIWindow(GUIType.GUI_SPELLS).visible = False Then IsRefining = False

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If

            ' Change map animation every 250 milliseconds
            If MapAnimTimer < Tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = Tick + 250
            End If
            
            ' Update inv animation
            If numitems > 0 Then
                If tmr100 < Tick Then
                    'If UZ Then
                    '    If PlanetTarget > 0 Then
                    '        If Planets(PlanetTarget).X + 20 < Player(MyIndex).X Or Planets(PlanetTarget).X - 20 > Player(MyIndex).X Then
                    '            PlanetTarget = 0
                    '        ElseIf Planets(PlanetTarget).Y + 16 < GetPlayerY(MyIndex) Or Planets(PlanetTarget).Y - 16 > GetPlayerY(MyIndex) Then
                    '            PlanetTarget = 0
                    '        End If
                    '    End If
                    'End If
                    DrawAnimatedInvItems
                    tmr100 = Tick + 100
                End If
            End If
            
            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next
            
            tmr25 = Tick + 25
        End If
        
5        If UZ Then
            For i = 1 To MAX_PLANETS
                If Planets(i).MoonData.Pic > 0 Then
                    If PlanetMoons(i).Tick < GetTickCount Then
                        If PlanetMoons(i).Local = 0 Then
                            PlanetMoons(i).Position = PlanetMoons(i).Position + 1
                            If PlanetMoons(i).Position > Planets(i).Size Then PlanetMoons(i).Local = 1
                        Else
                            PlanetMoons(i).Position = PlanetMoons(i).Position - 1
                            If PlanetMoons(i).Position < -(Planets(i).Size / 2) Then PlanetMoons(i).Local = 0
                        End If
                        PlanetMoons(i).Tick = GetTickCount + Planets(i).MoonData.speed
                    End If
                End If
            Next i
            For i = 1 To MAX_PLAYER_PLANETS
                If PlayerPlanet(i).PlanetData.MoonData.Pic > 0 Then
                    If PlayerPlanetMoons(i).Tick < GetTickCount Then
                        If PlayerPlanetMoons(i).Local = 0 Then
                            PlayerPlanetMoons(i).Position = PlayerPlanetMoons(i).Position + 1
                            If PlayerPlanetMoons(i).Position > PlayerPlanet(i).PlanetData.Size Then PlayerPlanetMoons(i).Local = 1
                        Else
                            PlayerPlanetMoons(i).Position = PlayerPlanetMoons(i).Position - 1
                            If PlayerPlanetMoons(i).Position < -(PlayerPlanet(i).PlanetData.Size / 2) Then PlayerPlanetMoons(i).Local = 0
                        End If
                        PlayerPlanetMoons(i).Tick = GetTickCount + PlayerPlanet(i).PlanetData.MoonData.speed
                    End If
                End If
            Next i
        End If
        
        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < Tick Then

            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                End If
            Next i

            ' Process npc movements (actually move them)
            For i = 1 To Npc_HighIndex
                If MapNpc(i).num > 0 Then
                    Call ProcessNpcMovement(i)
                End If
            Next i

            WalkTimer = Tick + 30 ' edit this value to change WalkTimer
        End If
        
        If chatTmr < Tick Then
            If ChatButtonUp Then
                ScrollChatBox 0
            End If
            If ChatButtonDown Then
                ScrollChatBox 1
            End If
            chatTmr = Tick + 50
        End If
        
        ' targetting
6        If targettmr < Tick Then
            If tabDown Then
                FindNearestTarget
            End If
            targettmr = Tick + 50
        End If
        
        ' fog scrolling
        If fogtmr < Tick Then
            If CurrentFogSpeed > 0 Then
                If Int(CurrentFogSpeed / 100) < 1 Then
                    ' move
                    If Map.FogDir = 1 Or Map.FogDir = 0 Then fogOffsetX = fogOffsetX - 1
                    If Map.FogDir = 2 Or Map.FogDir = 0 Then fogOffsetY = fogOffsetY - 1
                Else
                    ' move
                    If Map.FogDir = 1 Or Map.FogDir = 0 Then fogOffsetX = fogOffsetX - Int(CurrentFogSpeed / 100)
                    If Map.FogDir = 2 Or Map.FogDir = 0 Then fogOffsetY = fogOffsetY - Int(CurrentFogSpeed / 100)
                End If
                ' reset
                If fogOffsetX < -256 Then fogOffsetX = 0
                If fogOffsetY < -256 Then fogOffsetY = 0
                fogtmr = Tick + 255 - CurrentFogSpeed
            End If
        End If
        
        ' elastic bars
7        If bartmr < Tick Then
            SetBarWidth BarWidth_GuiHP_Max, BarWidth_GuiHP
            SetBarWidth BarWidth_GuiSP_Max, BarWidth_GuiSP
            SetBarWidth BarWidth_GuiEXP_Max, BarWidth_GuiEXP
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    SetBarWidth BarWidth_NpcHP_Max(i), BarWidth_NpcHP(i)
                End If
            Next
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    SetBarWidth BarWidth_PlayerHP_Max(i), BarWidth_PlayerHP(i)
                End If
            Next
            
            ' reset timer
            bartmr = Tick + 10
        End If
        
        ' ****** Parallax Y ******
8        If ParallaxY = 0 Then
            ParallaxY = -600
        Else
            ParallaxY = ParallaxY + 1
        End If
        
        ' ****** Parallax X ******
        If ParallaxX = -800 Then
            ParallaxX = 0
        Else
            If UZ Then
                ParallaxX = ParallaxX - 0.5
            Else
                ParallaxX = ParallaxX - 1
            End If
        End If
        
        If tmr500 < Tick Then
        
            ' animate waterfalls
            Select Case waterfallFrame
                Case 0
                    waterfallFrame = 1
                Case 1
                    waterfallFrame = 2
                Case 2
                    waterfallFrame = 0
            End Select
            
            ' animate autotiles
            Select Case autoTileFrame
                Case 0
                    autoTileFrame = 1
                Case 1
                    autoTileFrame = 2
                Case 2
                    autoTileFrame = 0
            End Select
            
            ' animate textbox
            If chatOn Then
                If chatShowLine = "|" Then
                    chatShowLine = vbNullString
                Else
                    chatShowLine = "|"
                End If
            End If
            
            Call CheckAFK
            
            If GetPlayerMap(MyIndex) > 0 Then
                If MapSaibamans(GetPlayerMap(MyIndex)).TotalSaibamans > 0 Then
                    For i = 1 To MapSaibamans(GetPlayerMap(MyIndex)).TotalSaibamans
                        If MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(i).Working = 1 Then
                            DoAnimation ConstructAnim, MapSaibamans(Player(MyIndex).Map).Saibaman(i).X, MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(i).Y, 0, 0, 0
                        End If
                    Next i
                End If
            End If
            
            tmr500 = Tick + 500
        End If
        
9        ProcessWeather
        
        If Fadetmr < Tick Then
            If FadeType <> 2 Then
                If FadeType = 1 Then
                    If FadeAmount = 255 Then
                        
                    Else
                        FadeAmount = FadeAmount + 15
                    End If
                ElseIf FadeType = 0 Then
                    If FadeAmount = 0 Then
                    
                    Else
                        FadeAmount = FadeAmount - 15
                    End If
                End If
            End If
            Fadetmr = Tick + 30
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
10        If rendertmr < Tick Then
11            Call Render_Graphics
12            Call UpdateSounds
            rendertmr = Tick + 25
        End If
        
        DoEvents

        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < Tick + 20
                DoEvents
                Sleep 1
            Loop
        End If
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            
            
            
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If

    Loop

    frmMain.visible = False

    If isLogging Then
        isLogging = False
        frmMenu.visible = True
        GettingMap = True
        frmMain.lblLoad.Tag = GetTickCount
        StopMusic
        'PlayMusic Options.MenuMusic
    Else
        ' Shutdown the game
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameLoop at " & Erl, "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Call DestroyGame
    Exit Sub
End Sub

Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Double

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    ' Check if player is walking, and if so process moving them over
    Select Case TempPlayer(Index).moving
        Case MOVING_WALKING: MovementSpeed = ((ElapsedTime / 1000) * (WALK_SPEED(Index) * SIZE_X))
        Case MOVING_RUNNING: MovementSpeed = ((ElapsedTime / 1000) * (RUN_SPEED(Index) * SIZE_X))
        Case Else: Exit Sub
    End Select
    
    On Error Resume Next
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            TempPlayer(Index).YOffSet = TempPlayer(Index).YOffSet - MovementSpeed
            If TempPlayer(Index).YOffSet < 0 Then TempPlayer(Index).YOffSet = 0
        Case DIR_DOWN
            TempPlayer(Index).YOffSet = TempPlayer(Index).YOffSet + MovementSpeed
            If TempPlayer(Index).YOffSet > 0 Then TempPlayer(Index).YOffSet = 0
        Case DIR_LEFT
            TempPlayer(Index).XOffSet = TempPlayer(Index).XOffSet - MovementSpeed
            If TempPlayer(Index).XOffSet < 0 Then TempPlayer(Index).XOffSet = 0
        Case DIR_RIGHT
            TempPlayer(Index).XOffSet = TempPlayer(Index).XOffSet + MovementSpeed
            If TempPlayer(Index).XOffSet > 0 Then TempPlayer(Index).XOffSet = 0
        Case DIR_UP_LEFT

                 TempPlayer(Index).YOffSet = TempPlayer(Index).YOffSet - MovementSpeed
        
                 If TempPlayer(Index).YOffSet < 0 Then TempPlayer(Index).YOffSet = 0
        
                 TempPlayer(Index).XOffSet = TempPlayer(Index).XOffSet - MovementSpeed
        
                 If TempPlayer(Index).XOffSet < 0 Then TempPlayer(Index).XOffSet = 0
        
        Case DIR_UP_RIGHT
        
                 TempPlayer(Index).YOffSet = TempPlayer(Index).YOffSet - MovementSpeed
        
                 If TempPlayer(Index).YOffSet < 0 Then TempPlayer(Index).YOffSet = 0
        
                 TempPlayer(Index).XOffSet = TempPlayer(Index).XOffSet + MovementSpeed
        
                 If TempPlayer(Index).XOffSet > 0 Then TempPlayer(Index).XOffSet = 0
        
        Case DIR_DOWN_LEFT
        
                 TempPlayer(Index).YOffSet = TempPlayer(Index).YOffSet + MovementSpeed
        
                 If TempPlayer(Index).YOffSet > 0 Then TempPlayer(Index).YOffSet = 0
        
                 TempPlayer(Index).XOffSet = TempPlayer(Index).XOffSet - MovementSpeed
        
                 If TempPlayer(Index).XOffSet < 0 Then TempPlayer(Index).XOffSet = 0
        
        Case DIR_DOWN_RIGHT
        
                 TempPlayer(Index).YOffSet = TempPlayer(Index).YOffSet + MovementSpeed
        
                 If TempPlayer(Index).YOffSet > 0 Then TempPlayer(Index).YOffSet = 0
        
                 TempPlayer(Index).XOffSet = TempPlayer(Index).XOffSet + MovementSpeed
        
                 If TempPlayer(Index).XOffSet > 0 Then TempPlayer(Index).XOffSet = 0
    End Select

    ' Check if completed walking over to the next tile
    If TempPlayer(Index).moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Or GetPlayerDir(Index) = DIR_DOWN_LEFT Or GetPlayerDir(Index) = DIR_DOWN_RIGHT Then
            If (TempPlayer(Index).XOffSet >= 0) And (TempPlayer(Index).YOffSet >= 0) Then
                TempPlayer(Index).moving = 0
                If VXFRAME = False Then
                    If TempPlayer(Index).Step = 1 Then
                        TempPlayer(Index).Step = 3
                    Else
                        TempPlayer(Index).Step = 1
                    End If
                Else
                    If TempPlayer(Index).Step = 0 Then
                        TempPlayer(Index).Step = 1
                    Else
                        TempPlayer(Index).Step = 0
                    End If
                End If
            End If
        Else
            If (TempPlayer(Index).XOffSet <= 0) And (TempPlayer(Index).YOffSet <= 0) Then
                TempPlayer(Index).moving = 0
                If VXFRAME = False Then
                    If TempPlayer(Index).Step = 1 Then
                        TempPlayer(Index).Step = 3
                    Else
                        TempPlayer(Index).Step = 1
                    End If
                Else
                    If TempPlayer(Index).Step = 0 Then
                        TempPlayer(Index).Step = 1
                    Else
                        TempPlayer(Index).Step = 0
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    If MapNpc(MapNpcNum).num = 0 Then Exit Sub

    ' Check if NPC is walking, and if so process moving them over
    If TempMapNpc(MapNpcNum).moving = MOVING_WALKING Then
        
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                TempMapNpc(MapNpcNum).YOffSet = TempMapNpc(MapNpcNum).YOffSet - ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))
                If TempMapNpc(MapNpcNum).YOffSet < 0 Then TempMapNpc(MapNpcNum).YOffSet = 0
                
            Case DIR_DOWN
                TempMapNpc(MapNpcNum).YOffSet = TempMapNpc(MapNpcNum).YOffSet + ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))
                If TempMapNpc(MapNpcNum).YOffSet > 0 Then TempMapNpc(MapNpcNum).YOffSet = 0
                
            Case DIR_LEFT
                TempMapNpc(MapNpcNum).XOffSet = TempMapNpc(MapNpcNum).XOffSet - ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))
                If TempMapNpc(MapNpcNum).XOffSet < 0 Then TempMapNpc(MapNpcNum).XOffSet = 0
                
            Case DIR_RIGHT
                TempMapNpc(MapNpcNum).XOffSet = TempMapNpc(MapNpcNum).XOffSet + ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))
                If TempMapNpc(MapNpcNum).XOffSet > 0 Then TempMapNpc(MapNpcNum).XOffSet = 0
                
            Case DIR_UP_LEFT

             TempMapNpc(MapNpcNum).YOffSet = TempMapNpc(MapNpcNum).YOffSet - ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))

             If TempMapNpc(MapNpcNum).YOffSet < 0 Then TempMapNpc(MapNpcNum).YOffSet = 0

             TempMapNpc(MapNpcNum).XOffSet = TempMapNpc(MapNpcNum).XOffSet - ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))

             If TempMapNpc(MapNpcNum).XOffSet < 0 Then TempMapNpc(MapNpcNum).XOffSet = 0

        

         Case DIR_UP_RIGHT

             TempMapNpc(MapNpcNum).YOffSet = TempMapNpc(MapNpcNum).YOffSet - ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))

             If TempMapNpc(MapNpcNum).YOffSet < 0 Then TempMapNpc(MapNpcNum).YOffSet = 0

             TempMapNpc(MapNpcNum).XOffSet = TempMapNpc(MapNpcNum).XOffSet + ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))

             If TempMapNpc(MapNpcNum).XOffSet > 0 Then TempMapNpc(MapNpcNum).XOffSet = 0

        

         Case DIR_DOWN_LEFT

             TempMapNpc(MapNpcNum).YOffSet = TempMapNpc(MapNpcNum).YOffSet + ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))

             If TempMapNpc(MapNpcNum).YOffSet > 0 Then TempMapNpc(MapNpcNum).YOffSet = 0

             TempMapNpc(MapNpcNum).XOffSet = TempMapNpc(MapNpcNum).XOffSet - ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))

             If TempMapNpc(MapNpcNum).XOffSet < 0 Then TempMapNpc(MapNpcNum).XOffSet = 0

            

         Case DIR_DOWN_RIGHT

             TempMapNpc(MapNpcNum).YOffSet = TempMapNpc(MapNpcNum).YOffSet + ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))

             If TempMapNpc(MapNpcNum).YOffSet > 0 Then TempMapNpc(MapNpcNum).YOffSet = 0

             TempMapNpc(MapNpcNum).XOffSet = TempMapNpc(MapNpcNum).XOffSet + ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).num).speed * SIZE_X))

             If TempMapNpc(MapNpcNum).XOffSet > 0 Then TempMapNpc(MapNpcNum).XOffSet = 0
        End Select
    
        ' Check if completed walking over to the next tile
        If TempMapNpc(MapNpcNum).moving > 0 Then
            If MapNpc(MapNpcNum).Dir = DIR_RIGHT Or MapNpc(MapNpcNum).Dir = DIR_DOWN Or MapNpc(MapNpcNum).Dir = DIR_DOWN_RIGHT Then
                If (TempMapNpc(MapNpcNum).XOffSet >= 0) And (TempMapNpc(MapNpcNum).YOffSet >= 0) Then
                    TempMapNpc(MapNpcNum).moving = 0
                    If VXFRAME = False Then
                        If TempMapNpc(MapNpcNum).Step = 1 Then
                            TempMapNpc(MapNpcNum).Step = 3
                        Else
                            TempMapNpc(MapNpcNum).Step = 1
                        End If
                    Else
                        If TempMapNpc(MapNpcNum).Step = 0 Then
                            TempMapNpc(MapNpcNum).Step = 1
                        Else
                            TempMapNpc(MapNpcNum).Step = 0
                        End If
                    End If
                End If
            Else
                If (TempMapNpc(MapNpcNum).XOffSet <= 0) And (TempMapNpc(MapNpcNum).YOffSet <= 0) Then
                    TempMapNpc(MapNpcNum).moving = 0
                    If VXFRAME = False Then
                        If TempMapNpc(MapNpcNum).Step = 1 Then
                            TempMapNpc(MapNpcNum).Step = 3
                        Else
                            TempMapNpc(MapNpcNum).Step = 1
                        End If
                    Else
                        If TempMapNpc(MapNpcNum).Step = 0 Then
                            TempMapNpc(MapNpcNum).Step = 1
                        Else
                            TempMapNpc(MapNpcNum).Step = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckMapGetItem()
Dim buffer As New clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer

    If GetTickCount > TempPlayer(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            TempPlayer(MyIndex).MapGetTimer = GetTickCount
            buffer.WriteLong CMapGetItem
            SendData buffer.ToArray()
        End If
    End If

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAttack()
Dim buffer As clsBuffer
Dim AttackSpeed As Long, X As Long, Y As Long, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If ControlDown Then
    
        If TempPlayer(MyIndex).SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If TempPlayer(MyIndex).StunDuration > 0 Then Exit Sub ' stunned, can't attack
        If TempPlayer(MyIndex).Fly = 1 Then Exit Sub

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            AttackSpeed = Item(GetPlayerEquipment(MyIndex, Weapon)).speed - (GetPlayerStat(MyIndex, Agility) * 5)
        Else
            AttackSpeed = 500 - (GetPlayerStat(MyIndex, Agility) * 5)
        End If
        
        If AttackSpeed < 250 Then AttackSpeed = 250
        
        If GetPlayerX(MyIndex) >= 0 And GetPlayerX(MyIndex) <= Map.MaxX Then
            If GetPlayerY(MyIndex) >= 0 And GetPlayerY(MyIndex) <= Map.MaxY Then
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_RESOURCE Then Exit Sub
            End If
        End If

        If TempPlayer(MyIndex).AttackTimer + AttackSpeed < GetTickCount Then
            If TempPlayer(MyIndex).Attacking = 0 Then

                With TempPlayer(MyIndex)
                    '.Attacking = 1
                    .AttackTimer = GetTickCount
                    '.AttackAnim = Rand(0, 1)
                End With

                If Not Map.Moral = 3 Then SendAttack
                If Player(MyIndex).InTutorial = 0 And TutorialStep = 16 Then
                    TutorialProgress = TutorialProgress + 1
                End If
            End If
        End If
        
        Select Case Player(MyIndex).Dir
            Case DIR_UP
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) - 1
            Case DIR_DOWN
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) + 1
            Case DIR_LEFT
                X = GetPlayerX(MyIndex) - 1
                Y = GetPlayerY(MyIndex)
            Case DIR_RIGHT
                X = GetPlayerX(MyIndex) + 1
                Y = GetPlayerY(MyIndex)
        End Select

    End If
    

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function IsTryingToMove() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If DirUp Or DirDown Or DirLeft Or DirRight Or DirUpLeft Or DirUpRight Or DirDownLeft Or DirDownRight Then
        IsTryingToMove = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CanMove() As Boolean
Dim d As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    CanMove = True
    
    If GUIWindow(GUI_DIALOGUE).visible = True And dialogueIndex = DIALOGUE_TYPE_SELLPLANET Then
        CanMove = False
        Exit Function
    End If
    
    If InTrade Then
        CanMove = False
        Exit Function
    End If
    
    If frmMain.picTroca.visible = True Then
        CanMove = False
        Exit Function
    End If
    
    If TutorialBlockWalk Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they aren't trying to move when they are already moving
    If TempPlayer(MyIndex).moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    If Map.Moral = 2 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If TempPlayer(MyIndex).SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not stunned
    If TempPlayer(MyIndex).StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    
    'está vivo
    If Player(MyIndex).IsDead = 1 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' not in bank
    If InBank Then
        'CanMove = False
        'Exit Function
        InBank = False
        GUIWindow(GUI_BANK).visible = False
    End If
    
    If TempPlayer(MyIndex).Fly = 0 And TempPlayer(MyIndex).FlyBalance > 0 Then
        CanMove = False
        Exit Function
    End If
    
    If GUIWindow(GUI_NEWS).visible = True Then
        GUIWindow(GUI_NEWS).visible = False
    End If

    d = GetPlayerDir(MyIndex)
    
        If DirUpLeft Then
    
         Call SetPlayerDir(MyIndex, DIR_UP_LEFT)
    
    
    
         ' Check to see if they are trying to go out of bounds
    
         If GetPlayerY(MyIndex) > 0 And GetPlayerX(MyIndex) > 0 Then
    
             If CheckDirection(DIR_UP_LEFT) Then
    
                 CanMove = False
    
    
    
                 ' Set the new direction if they weren't facing that direction
    
                 If d <> DIR_UP_LEFT Then
    
                     Call SendPlayerDir
    
                 End If
    
    
    
                 Exit Function
    
             End If
    
    
    
         Else
    
    
    

    
    
    
             CanMove = False
    
             Exit Function
    
         End If
    
    End If
    
    
    
    If DirUpRight Then
    
         Call SetPlayerDir(MyIndex, DIR_UP_RIGHT)
    
    
    
         ' Check to see if they are trying to go out of bounds
    
         If GetPlayerY(MyIndex) > 0 And GetPlayerX(MyIndex) < Map.MaxX Then
    
             If CheckDirection(DIR_UP_RIGHT) Then
    
                 CanMove = False
    
    
    
                 ' Set the new direction if they weren't facing that direction
    
                 If d <> DIR_UP_RIGHT Then
    
                     Call SendPlayerDir
    
                 End If
    
    
    
                 Exit Function
    
             End If
    
    
    
         Else
    
    

    
    
    
             CanMove = False
    
             Exit Function
    
         End If
    
    End If
    
    
    
    If DirDownLeft Then
    
         Call SetPlayerDir(MyIndex, DIR_DOWN_LEFT)
    
    
    
         ' Check to see if they are trying to go out of bounds
    
         If GetPlayerY(MyIndex) < Map.MaxY And GetPlayerX(MyIndex) > 0 Then
    
             If CheckDirection(DIR_DOWN_LEFT) Then
    
                 CanMove = False
    
    
    
                 ' Set the new direction if they weren't facing that direction
    
                 If d <> DIR_DOWN_LEFT Then
    
                     Call SendPlayerDir
    
                 End If
    
    
    
                 Exit Function
    
             End If
    
    
    
         Else
    
    
    

    
    
    
             CanMove = False
    
             Exit Function
    
         End If
    
    End If
    
    
    
    If DirDownRight Then
    
         Call SetPlayerDir(MyIndex, DIR_DOWN_RIGHT)
    
    
    
         ' Check to see if they are trying to go out of bounds
    
         If GetPlayerY(MyIndex) < Map.MaxY And GetPlayerX(MyIndex) < Map.MaxX Then
    
             If CheckDirection(DIR_DOWN_RIGHT) Then
    
                 CanMove = False
    
    
    
                 ' Set the new direction if they weren't facing that direction
    
                 If d <> DIR_DOWN_RIGHT Then
    
                     Call SendPlayerDir
    
                 End If
    
    
    
                 Exit Function
    
             End If
    
    
    
         Else
    
    
    

    
    
    
             CanMove = False
    
             Exit Function
    
         End If
    
    End If

    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
                frmMain.lblLoad.Tag = GetTickCount
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
                frmMain.lblLoad.Tag = GetTickCount
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
                frmMain.lblLoad.Tag = GetTickCount
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
                frmMain.lblLoad.Tag = GetTickCount
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean
Dim X As Long
Dim Y As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    CheckDirection = False
    
    If TempPlayer(MyIndex).Fly = 1 Then Exit Function
    
    ' check directional blocking
    If isDirBlocked(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, Direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case Direction
        Case DIR_UP
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
        Case DIR_UP_LEFT

             X = GetPlayerX(MyIndex) - 1
    
             Y = GetPlayerY(MyIndex) - 1
    
         Case DIR_UP_RIGHT
    
             X = GetPlayerX(MyIndex) + 1
    
             Y = GetPlayerY(MyIndex) - 1
    
         Case DIR_DOWN_LEFT
    
             X = GetPlayerX(MyIndex) - 1
    
             Y = GetPlayerY(MyIndex) + 1
    
         Case DIR_DOWN_RIGHT
    
             X = GetPlayerX(MyIndex) + 1
    
             Y = GetPlayerY(MyIndex) + 1
    End Select

    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        If Resource(Map.Tile(X, Y).data1).ResourceType <> 3 Then
            CheckDirection = True
        End If
        Exit Function
    End If
    
    If Map.Tile(X, Y).Type = TILE_TYPE_EVENT Then
        If Map.Tile(X, Y).data1 > 0 Then
            If Events(Map.Tile(X, Y).data1).WalkThrought = NO Then
                If Player(MyIndex).EventOpen(Map.Tile(X, Y).data1) = NO Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Check to see if a player is already on that tile
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If Map.Moral <> MAP_MORAL_OWNER Or Player(i).Instance = Player(MyIndex).Instance Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        If TempPlayer(i).AFK = 0 And CanShow(i) Then
                            CheckDirection = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "checkDirection", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub CheckMovement()
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If IsTryingToMove Then
        If CanMove Then
        
            

            ' Check if player has the shift key down for running
            If ShiftDown Then
                If GetPlayerVital(MyIndex, MP) > 0 Then
                    If Not TempPlayer(MyIndex).HairChange = 5 Then
                        TempPlayer(MyIndex).moving = MOVING_RUNNING
                    Else
                        TempPlayer(MyIndex).moving = MOVING_WALKING
                    End If
                Else
                    TempPlayer(MyIndex).moving = MOVING_WALKING
                End If
            Else
                TempPlayer(MyIndex).moving = MOVING_WALKING
            End If

            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    TempPlayer(MyIndex).YOffSet = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    TempPlayer(MyIndex).YOffSet = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    TempPlayer(MyIndex).XOffSet = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    TempPlayer(MyIndex).XOffSet = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                Case DIR_UP_LEFT

                 Call SendPlayerMove

                 TempPlayer(MyIndex).YOffSet = PIC_Y

                 Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)

                 TempPlayer(MyIndex).XOffSet = PIC_X

                 Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)

             Case DIR_UP_RIGHT

                 Call SendPlayerMove

                 TempPlayer(MyIndex).YOffSet = PIC_Y

                 Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)

                 TempPlayer(MyIndex).XOffSet = PIC_X * -1

                 Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)

             Case DIR_DOWN_LEFT

                 Call SendPlayerMove

                 TempPlayer(MyIndex).YOffSet = PIC_Y * -1

                 Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)

                 TempPlayer(MyIndex).XOffSet = PIC_X

                 Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)

             Case DIR_DOWN_RIGHT

                 Call SendPlayerMove

                 TempPlayer(MyIndex).YOffSet = PIC_Y * -1

                 Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)

                 TempPlayer(MyIndex).XOffSet = PIC_X * -1

                 Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select

            If TempPlayer(MyIndex).XOffSet = 0 Then
                If TempPlayer(MyIndex).YOffSet = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP And Not TempPlayer(MyIndex).Fly = 1 Then
                        GettingMap = True
                        frmMain.lblLoad.Tag = GetTickCount
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isInBounds()
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function


Public Sub UseItem()
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ForgetSpell(ByVal spellSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellSlot) > 0 Then
        AddText printf("Não pode esquecer uma técnica que está em espera!"), BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(MyIndex).SpellBuffer = spellSlot Then
        AddText printf("Não pode esquecer uma técnica que está conjurando!"), BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellSlot) > 0 Then
        Set buffer = New clsBuffer
        buffer.WriteLong CForgetSpell
        buffer.WriteLong spellSlot
        SendData buffer.ToArray()
        Set buffer = Nothing
    Else
        AddText printf("Nenhuma técnica encontrada."), BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal spellSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If Map.Moral = 3 Then Exit Sub
    
    If Not IsRefining Then
        If SpellCD(spellSlot) > 0 Then
            AddText printf("Técnica em espera!"), BrightRed
            Exit Sub
        End If
        
        If TempPlayer(MyIndex).StunDuration > 0 Then Exit Sub
        If PlayerSpells(spellSlot) = 0 Then Exit Sub
        If TempPlayer(MyIndex).SpellBuffer > 0 Then Exit Sub
        
        If TempPlayer(MyIndex).Fly = 1 Then
            If Spell(PlayerSpells(spellSlot)).Type <> SPELL_TYPE_TRANS And Spell(PlayerSpells(spellSlot)).Type <> SPELL_TYPE_VOAR Then
                Exit Sub
            End If
        End If
    
        ' Check if player has enough MP
        If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellSlot)).MPCost Then
            Call AddText(printf("Sem MP para conjurar a técnica %s.", Trim$(Spell(PlayerSpells(spellSlot)).name)), BrightRed)
            Exit Sub
        End If
    End If

    If PlayerSpells(spellSlot) > 0 Then
        If Not IsRefining Then
            If GetTickCount > TempPlayer(MyIndex).AttackTimer + 1000 Then
                If TempPlayer(MyIndex).moving = 0 Then
                    Set buffer = New clsBuffer
                    buffer.WriteLong CCast
                    buffer.WriteLong spellSlot
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    'TempPlayer(MyIndex).SpellBuffer = spellSlot
                    'TempPlayer(MyIndex).SpellBufferTimer = GetTickCount
                    'TempPlayer(MyIndex).SpellBufferNum = PlayerSpells(spellSlot)
                Else
                    Call AddText(printf("Você não pode efetuar uma técnica enquanto anda!"), BrightRed)
                End If
            End If
        Else
            Set buffer = New clsBuffer
                buffer.WriteLong CUpgrade
                buffer.WriteLong spellSlot
                SendData buffer.ToArray()
            Set buffer = Nothing
        End If
    Else
        Call AddText(printf("Nenhuma técnica aqui."), BrightRed)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CastSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DevMsg(ByVal Text As String, ByVal color As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(Text, color)
        End If
    End If

    Debug.Print Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DevMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "TwipsToPixels", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "PixelsToTwips", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) <= 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) <= 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub CacheResources()
Dim X As Long, Y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Resource_Count = 0

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If
        Next
    Next

    Resource_Index = Resource_Count
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateActionMsg(ByVal Message As String, ByVal color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .Message = Message
        .color = color
        .Type = MsgType
        .Created = GetTickCount
        .Scroll = 1
        .X = X
        .Y = Y
        .Alpha = 255
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    End If
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ActionMsg(Index).Message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).X = 0
    ActionMsg(Index).Y = 0
    ActionMsg(Index).Alpha = 0
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
Dim looptime As Long
Dim Layer As Long
Dim FrameCount As Long
Dim lockindex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub
    
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(Layer) = 0 Then AnimInstance(Index).frameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).Timer(Layer) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimInstance(Index).frameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                        If AnimInstance(Index).ReturnAnim > 0 Then DoAnimation AnimInstance(Index).ReturnAnim, AnimInstance(Index).X, AnimInstance(Index).Y, AnimInstance(Index).LockType, AnimInstance(Index).lockindex, AnimInstance(Index).Dir, AnimInstance(Index).ReturnAnim
                    Else
                        AnimInstance(Index).frameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).frameIndex(Layer) = AnimInstance(Index).frameIndex(Layer) + 1
                End If
                AnimInstance(Index).Timer(Layer) = GetTickCount
            End If
        End If
    Next
    
    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "checkAnimInstance", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenShop(ByVal shopnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    InShop = shopnum
    ShopAction = 0

    GUIWindow(GUI_SHOP).visible = True
    GUIWindow(GUI_INVENTORY).visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    GetBankItemNum = Bank.Item(bankslot).num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal ItemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).num = ItemNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    GetBankItemValue = Bank.Item(bankslot).value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef Dir As Byte, ByVal block As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If block Then
        blockvar = blockvar Or (2 ^ Dir)
    Else
        blockvar = blockvar And Not (2 ^ Dir)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        'isDirBlocked = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long, Optional CastAnim As Byte = 0)
Dim soundName As String

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If entityNum <= 0 Then Exit Sub
    
    ' find the sound
    Select Case entityType
        ' animations
        Case SoundEntity.seAnimation
            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).Sound)
            
        ' items
        Case SoundEntity.seItem
            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).Sound)
        ' npcs
        Case SoundEntity.seNpc
            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(Npc(entityNum).Sound)
        ' resources
        Case SoundEntity.seResource
            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).Sound)
        ' spells
        Case SoundEntity.seSpell
            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(Spell(entityNum).Sound)
        ' effects
        Case SoundEntity.seEffect
            If entityNum > MAX_EFFECTS Then Exit Sub
            soundName = Trim$(Effect(entityNum).Sound)
        ' other
        Case Else
            Exit Sub
    End Select
    
    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub
    If InGame = False Then Exit Sub

    ' play the sound
    PlaySound soundName, X, Y, CastAnim
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal data1 As Long = 0)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' exit out if we've already got a dialogue open
    If dialogueIndex > 0 Then Exit Sub
    
    ' set global dialogue index
    dialogueIndex = diIndex
    
    ' set the global dialogue data
    dialogueData1 = data1

    ' set the captions
    Dialogue_TitleCaption = diTitle
    Dialogue_TextCaption = diText
    
    ' show/hide buttons
    If Not isYesNo Then
        Dialogue_ButtonVisible(1) = False ' Yes button
        Dialogue_ButtonVisible(2) = True ' Okay button
        Dialogue_ButtonVisible(3) = False ' No button
    Else
        Dialogue_ButtonVisible(1) = True ' Yes button
        Dialogue_ButtonVisible(2) = False ' Okay button
        Dialogue_ButtonVisible(3) = True ' No button
    End If
    
    ' show the dialogue box
    GUIWindow(GUI_DIALOGUE).visible = True
    inChat = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Dialogue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleDialogueConfirm()
    Select Case dialogueData1
        Case 1 'Remove construção
            SendRemoveBlock
    End Select
End Sub

Sub HandleDialogueConfirmation(ByVal YesNo As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CConfirmation
    buffer.WriteByte YesNo
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub dialogueHandler(ByVal Index As Long)
    ' find out which button
    If Index = 1 Then ' okay button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_NAMEPLANET
                SendPlanetName
        End Select
    ElseIf Index = 2 Then ' yes button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendAcceptTradeRequest
            Case DIALOGUE_TYPE_FORGET
                ForgetSpell dialogueData1
            Case DIALOGUE_TYPE_PARTY
                SendAcceptParty
            Case DIALOGUE_TYPE_SELLPLANET
                SendAcceptSell
            Case DIALOGUE_TYPE_GUILDINVITE
                SendGuildInviteAnswer 1
            Case DIALOGUE_TYPE_ARENAINVITE
                SendArenaAccept
            Case DIALOGUE_TYPE_CONFIRM
                HandleDialogueConfirm
            Case DIALOGUE_TYPE_CONFIRMATION
                HandleDialogueConfirmation 1
        End Select
    ElseIf Index = 3 Then ' no button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendDeclineTradeRequest
            Case DIALOGUE_TYPE_PARTY
                SendDeclineParty
            Case DIALOGUE_TYPE_SELLPLANET
                SendDeclineSell
            Case DIALOGUE_TYPE_GUILDINVITE
                SendGuildInviteAnswer 0
            Case DIALOGUE_TYPE_ARENAINVITE
                SendArenaDeny
            Case DIALOGUE_TYPE_CONFIRMATION
                HandleDialogueConfirmation 0
        End Select
    End If
End Sub

Public Function GetColorString(color As Long)
    Select Case color
        Case 0
            GetColorString = "Black"
        Case 1
            GetColorString = "Blue"
        Case 2
            GetColorString = "Green"
        Case 3
            GetColorString = "Cyan"
        Case 4
            GetColorString = "Red"
        Case 5
            GetColorString = "Magenta"
        Case 6
            GetColorString = "Brown"
        Case 7
            GetColorString = "Grey"
        Case 8
            GetColorString = "Dark Grey"
        Case 9
            GetColorString = "Bright Blue"
        Case 10
            GetColorString = "Bright Green"
        Case 11
            GetColorString = "Bright Cyan"
        Case 12
            GetColorString = "Bright Red"
        Case 13
            GetColorString = "Pink"
        Case 14
            GetColorString = "Yellow"
        Case 15
            GetColorString = "White"

    End Select
End Function

Public Sub MenuLoop()
Dim FrameTime As Long
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim rec As RECT
Dim TickDelay As Long
Dim rec_pos As RECT, srcRect As D3DRECT
Dim Alpha As Long, degrees As Integer, Size As Long, moving As Long, bannerx As Long
Dim loginbannery As Long, Createcharposition As Byte
Dim AntiHackTmr As Long
Static PersonagemX, PersonagemY As Long, PersonagemDir As Long, PersonagemBalance As Double, PersonagemBalanceDir As Long
Static MenuPeriod As Byte, TickNewMenu As Long
Static MusicStart As Boolean
Dim Nuvens As Byte
    If UZ = False Then
    Nuvens = Rand(17, 19)
    If Nuvens = 17 Then Nuvens = 4
    Else
        Nuvens = 20
    End If
    ' If debug mode, handle error then exit out
    TickNewMenu = GetTickCount

    'On Error GoTo ErrorHandler
restartmenuloop:
    ' *** Start GameLoop ***
    Do While Not InGame
        Tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.
        
        If AutoLogin = True And isLogging = False Then InGame = True
        
        If AntiHackTmr < Tick Then
            VerificarAntiHack
            AntiHackTmr = Tick + 1000
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        'Call DrawGDI
        
        'Check for device lost.
        If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
        
        ' don't render
        If frmMain.WindowState <> vbMinimized Then
        
            ' unload any textures we need to unload
            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
                
            Direct3D_Device.BeginScene
            'Inicio ///////////////////////////////////////////////////////
            
            ' Get rec
                With rec
                    .Top = Camera.Top
                    .Bottom = .Top + ScreenY
                    .Left = Camera.Left
                    .Right = .Left + ScreenX
                End With
                    
                ' rec_pos
                With rec_pos
                    .Bottom = ScreenY
                    .Right = ScreenX
                End With
            
            With srcRect
                    .X1 = 0
                    .X2 = frmMain.ScaleWidth
                    .Y1 = 0
                    .Y2 = frmMain.ScaleHeight
                End With
            
            'Mouse position
            If DrawMousePosition = True Then
                RenderText Font_Default, "MouseX: " & CStr(GlobalX) & " MouseY:" & CStr(GlobalY), 0, 0, Yellow, 0
            End If
            
            'Goplay logo
            If MenuPeriod = 0 And TickNewMenu + 1000 < GetTickCount And TickNewMenu + 3900 > GetTickCount Then
                If TickNewMenu + 3000 > GetTickCount Then
                Alpha = 0 + ((GetTickCount - (TickNewMenu + 1000)) / 4)
                If Alpha > 255 Then Alpha = 255
                Else
                Alpha = 255 - ((GetTickCount - (TickNewMenu + 6000)) / 4)
                If Alpha < 0 Then Alpha = 0
                End If
                RenderTexture Tex_NewGUI(NewGui.Loading), 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, frmMain.ScaleWidth, frmMain.ScaleHeight, D3DColorRGBA(255, 255, 255, Alpha)
            End If
            
            'Musica
            If TickNewMenu + 3000 < GetTickCount And MusicStart = False Then
                TickDelay = GetTickCount
                Call PlayMusic("DBZ Cha-la.mid")
                TickNewMenu = TickNewMenu + (GetTickCount - TickDelay)
                MusicStart = True
            End If
            
            'Esfera
            If TickNewMenu + 5000 < GetTickCount And TickNewMenu + 6500 > GetTickCount Then
                Alpha = 0 + ((GetTickCount - (TickNewMenu + 5000)) / 4)
                If Alpha > 255 Then Alpha = 255
                degrees = degrees + 12
                If degrees > 360 Then degrees = 0
                Size = Size + 2
                If Size > 100 Then Size = 100
                RenderTexture Tex_NewGUI(NewGui.Esfera), 400 - (Size / 2), 300 - (Size / 2), 0, 0, Size, Size, 100, 100, D3DColorRGBA(255, 255, 255, Alpha), degrees
            End If
            
            If TickNewMenu + 8500 < GetTickCount And StateMenu = 0 Then StateMenu = MenuType.MENU_LOGIN
            
            'Menu de login
            If TickNewMenu + 7000 < GetTickCount Then
                'Nuvens \\
                moving = moving - 5
                If moving = -800 Then moving = 0
                RenderTexture Tex_NewGUI(Nuvens), moving, -1, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight + 4, 800, 600, D3DColorRGBA(255, 255, 255, 255)
                RenderTexture Tex_NewGUI(Nuvens), moving + 800, -1, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight + 4, 800, 600, D3DColorRGBA(255, 255, 255, 255)
                '/////////
                
                'Logo \\
                Static LogoAlpha As Byte
                If NewCharTick = 0 Then
                    If LogoAlpha < 254 Then LogoAlpha = LogoAlpha + 2
                    RenderTexture Tex_NewGUI(NewGui.Logo), (frmMain.ScaleWidth / 2) - 160, (frmMain.ScaleHeight / 2) - 55, 0, 0, 320, 110, 320, 110, D3DColorRGBA(255, 255, 255, LogoAlpha)
                    RenderText Font_Default, "Versão " & App.Major & "." & App.Minor, (frmMain.ScaleWidth / 2) - (getWidth(Font_Default, "Versão " & App.Major & "." & App.Minor) / 2), (frmMain.ScaleHeight / 2) + 55, Yellow, 255 - LogoAlpha
                    RenderTexture Tex_GUI(43), (frmMain.ScaleWidth / 2) - 150, (frmMain.ScaleHeight / 2) + 140, 0, 0, 300, 60, 300, 60, D3DColorRGBA(255, 255, 255, LogoAlpha)
                Else
                    If (frmMain.ScaleHeight / 2) - 55 - Int((GetTickCount - NewCharTick) / 2) > 25 Then
                        RenderTexture Tex_NewGUI(NewGui.Logo), (frmMain.ScaleWidth / 2) - 160, (frmMain.ScaleHeight / 2) - 55 - Int((GetTickCount - NewCharTick) / 2), 0, 0, 320, 110, 320, 110, D3DColorRGBA(255, 255, 255, 255)
                    Else
                        RenderTexture Tex_NewGUI(NewGui.Logo), (frmMain.ScaleWidth / 2) - 160, 25, 0, 0, 320, 110, 320, 110, D3DColorRGBA(255, 255, 255, 255)
                    End If
                End If
                '///////
                
                'Servers
                If isLogging And NewCharTick = 0 Then
                    Dim i As Long, color As Long
                    RenderText Font_Default, printf("Use as teclas de seta para selecionar o servidor"), (frmMain.ScaleWidth / 2) - (getWidth(Font_Default, printf("Use as teclas de seta para selecionar o servidor")) / 2), 80, White, 0
                    For i = 1 To UBound(Options.Servers)
                        If Options.Servers(i).Ping < 200 Then color = BrightGreen
                        If Options.Servers(i).Ping >= 200 And Options.Servers(i).Ping < 400 Then color = Yellow
                        If Options.Servers(i).Ping >= 400 Then color = BrightRed
                        If SelectedServer = i Then
                            color = White
                            RenderTexture Tex_NewGUI(NewGui.Estrela), (frmMain.ScaleWidth / 2) - (getWidth(Font_Default, (printf("%s Ping %d ms", Options.Servers(i).name & "," & Options.Servers(i).Ping))) / 2) - 32, 88 + (i * 16), 0, 0, 31, 31, 31, 31, D3DColorRGBA(255, 255, 255, 255)
                        End If
                        If Options.Servers(i).Ping < 3000 Then
                            RenderText Font_Default, printf("%s Ping %d ms", Options.Servers(i).name & "," & Options.Servers(i).Ping), (frmMain.ScaleWidth / 2) - (getWidth(Font_Default, (printf("%s Ping %d ms", Options.Servers(i).name & "," & Options.Servers(i).Ping))) / 2), 96 + (i * 16), color, 0
                        Else
                            RenderText Font_Default, printf("%s sem resposta", Options.Servers(i).name), (frmMain.ScaleWidth / 2) - (getWidth(Font_Default, (printf("%s sem resposta", Options.Servers(i).name))) / 2), 96 + (i * 16), color, 0
                        End If
                        
                    Next i
                End If
                
                'Barra de login \\
                If NewCharTick = 0 Then
                    If TickNewMenu + 9000 > GetTickCount Then
                        Flying = Flying + 20
                        If Flying > 900 Then Flying = 900
                        If Flying < (frmMain.ScaleWidth / 2) + 300 And bannerx <> (frmMain.ScaleWidth / 2) + 300 Then
                            bannerx = Flying
                        Else
                            bannerx = (frmMain.ScaleWidth / 2) + 300
                        End If
                        
                        RenderTexture Tex_NewGUI(NewGui.Login), bannerx - 600, 25, 0, 0, 600, 40, 600, 40, D3DColorRGBA(255, 255, 255, 255)
                        'Sayajin voador
                        RenderTexture Tex_NewGUI(NewGui.Flyingbanner), Flying + 32, 45, 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 255)
                        '////////
                    Else
                        RenderTexture Tex_NewGUI(NewGui.Login), (frmMain.ScaleWidth / 2) - 300, 25, 0, 0, 600, 40, 600, 40, D3DColorRGBA(255, 255, 255, 255)
                    End If
                Else
                    
                    If NewCharTick + 1000 > GetTickCount Then RenderTexture Tex_NewGUI(NewGui.Login), (frmMain.ScaleWidth / 2) - 300, PersonagemY - 25, 0, 0, 600, 40, 600, 40, D3DColorRGBA(255, 255, 255, 255)
                End If
                
                If Len(NewGUIWindow(TEXTLOGIN).value) > 3 And NewCharTick = 0 Then RenderTexture Tex_NewGUI(NewGui.Estrela), 313, 29, 0, 0, 31, 31, 31, 31, D3DColorRGBA(255, 255, 255, 255)
                If Len(NewGUIWindow(TEXTPASSWORD).value) > 3 And NewCharTick = 0 Then RenderTexture Tex_NewGUI(NewGui.Estrela), 570, 28, 0, 0, 31, 31, 31, 31, D3DColorRGBA(255, 255, 255, 255)
                If MsgScreen <> "" Then
                    If NewCharTick = 0 Then
                        RenderText Font_Default, MsgScreen, 400 - ((getWidth(Font_Default, (Trim$(MsgScreen))) / 2)), 64, White, 0
                    Else
                        RenderText Font_Default, MsgScreen, 400 - ((getWidth(Font_Default, (Trim$(MsgScreen))) / 2)), 150, White, 0
                    End If
                End If
            End If
            
            If TickNewMenu + 9000 < GetTickCount Then
                'Barra de direitos autorais \\
                If NewCharTick = 0 Then
                Flying = Flying - 20
                If Flying < -100 Then Flying = -100
                If Flying > (frmMain.ScaleWidth / 2) - 300 Then
                    If bannerx <> (frmMain.ScaleWidth / 2) - 300 Then
                        bannerx = Flying
                    End If
                Else
                    bannerx = (frmMain.ScaleWidth / 2) - 300
                End If
                
                'Sayajin voador
                RenderTexture Tex_NewGUI(NewGui.Flyingbanner2), Flying - 48, 500, 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 255)
                '////////
                End If
                
                RenderTexture Tex_NewGUI(NewGui.Banner), bannerx, 500, 0, 0, 600, 40, 600, 40, D3DColorRGBA(255, 255, 255, 255)
                
            End If
            
            If NewCharTick + 1000 < GetTickCount And NewCharTick <> 0 Then
                If Flying = -100 Then Flying = 900
                Flying = Flying - 20
                If Flying < -50 Then Flying = -50
                
                'Personagem
                RenderTexture Tex_Character(Class(NewCharClasse).MaleSprite(newCharSprite)), Flying - 32, 300, (32 * 20), 64, 32, 64, 32, 64, D3DColorRGBA(255, 255, 255, 255)
                RenderTexture Tex_Hair(0).TexHair(newCharHair), Flying - 32, 300, (32 * 20), 64, 32, 64, 32, 64, D3DColorRGBA(255, 255, 255, 255)
            
                If Flying > (frmMain.ScaleWidth / 2) - 113 Then
                    RenderTexture Tex_NewGUI(NewGui.NewChar), Flying, 300 - 118, 0, 0, 226, 236, 226, 236, D3DColorRGBA(255, 255, 255, 255)
                Else
                    RenderTexture Tex_NewGUI(NewGui.NewChar), (frmMain.ScaleWidth / 2) - 113, 300 - 118, 0, 0, 226, 236, 226, 236, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
            
            Call DrawNewGui
            
            If NewGUIWindow(TEXTCHARNAME).value <> "" Then
                RenderText Font_Default, NewGUIWindow(TEXTCHARNAME).value, PersonagemX + 14 - ((getWidth(Font_Default, (Trim$(NewGUIWindow(TEXTCHARNAME).value))) / 2)), PersonagemY - 8, Yellow, 0
            End If
            
            'Personagem
            If TickNewMenu + 11000 < GetTickCount Then
                PersonagemDir = 0
                Createcharposition = 19
                If InGameTick = 0 Then
                    If GlobalX < PersonagemX Then
                    PersonagemX = PersonagemX - Int((PersonagemX - GlobalX) / 10)
                    PersonagemDir = 1
                    If Int((PersonagemX - GlobalX) / 10) = 0 Then PersonagemDir = 0
                    Else
                    PersonagemX = PersonagemX + Int((GlobalX - PersonagemX) / 10)
                    PersonagemDir = 2
                    If Int((GlobalX - PersonagemX) / 10) = 0 Then PersonagemDir = 0
                    End If
                    
                    If NewCharTick <> 0 And NewCharTick + 3000 > GetTickCount Then GlobalY = -100
                    If GlobalY < PersonagemY Then
                    PersonagemY = PersonagemY - Int((PersonagemY - GlobalY) / 10)
                    Else
                    PersonagemY = PersonagemY + Int((GlobalY - PersonagemY) / 10)
                    End If
                    
                    If GlobalY - 9 = PersonagemY And GlobalX - 9 = PersonagemX Then PersonagemDir = 0
                    
                    If PersonagemY + 64 > 600 Then PersonagemY = 599 - 64
                    If PersonagemX + 32 > 800 Then PersonagemX = 799 - 32
                    
                    If PersonagemDir = 0 Then
                        If PersonagemBalanceDir = 0 Then
                            PersonagemBalance = PersonagemBalance + 0.3
                            If PersonagemBalance > 2 Then PersonagemBalanceDir = 1
                        Else
                            PersonagemBalance = PersonagemBalance - 0.3
                            If PersonagemBalance < -2 Then PersonagemBalanceDir = 0
                        End If
                    Else
                        PersonagemBalance = 0
                    End If
                    
                    If NewCharTick <> 0 And NewCharTick + 1000 > GetTickCount Then PersonagemDir = 5
                Else
                
                    If InGameTick + 2500 > GetTickCount Then
                        If NewCharTick = 0 Then
                            PersonagemDir = 3
                        Else
                            Createcharposition = 10
                            PersonagemDir = 0
                        End If
                    Else
                        If NewCharTick = 0 Then
                            PersonagemDir = 4
                        Else
                            Createcharposition = 11
                            PersonagemDir = 0
                        End If
                    End If
                    
                    If InGameTick + 3000 < GetTickCount Then
                        InGame = True
                    End If
                
                End If
                
                If NewCharTick = 0 Or NewCharTick + 3000 > GetTickCount Then
                    RenderTexture Tex_NewGUI(NewGui.Personagem), PersonagemX, PersonagemY + PersonagemBalance, 0, (64 * PersonagemDir), 32, 64, 32, 64, D3DColorRGBA(255, 255, 255, 255)
                Else
                    RenderTexture Tex_Character(Class(NewCharClasse).MaleSprite(newCharSprite)), PersonagemX, PersonagemY + PersonagemBalance, (32 * Createcharposition), (64 * PersonagemDir), 32, 64, 32, 64, D3DColorRGBA(255, 255, 255, 255)
                    RenderTexture Tex_Hair(0).TexHair(newCharHair), PersonagemX, PersonagemY + PersonagemBalance, (32 * Createcharposition), (64 * PersonagemDir), 32, 64, 32, 64, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
            
            'Fade
            If TickNewMenu + 6000 < GetTickCount Then
                If TickNewMenu + 7000 > GetTickCount Then
                Alpha = 0 + ((GetTickCount - (TickNewMenu + 6000)) / 2)
                If Alpha > 255 Then Alpha = 255
                Else
                Alpha = 255 - ((GetTickCount - (TickNewMenu + 7000)) / 2)
                If Alpha < 0 Then Alpha = 0
                End If
                RenderTexture Tex_NewGUI(NewGui.Fader), 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, Alpha)
            End If
            
            'Fim //////////////////////////////////////////////////////////
            Direct3D_Device.EndScene
            
            If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
                HandleDeviceLost
                Exit Sub
            Else
                Direct3D_Device.Present srcRect, ByVal 0, 0, ByVal 0
                DrawGDI
            End If
        End If
    
        DoEvents
        
        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < Tick + 20
                DoEvents
                Sleep 1
            Loop
        End If
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
    Loop
    
    If InGameTick <> 0 Then
        'Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
        DoInGame
    End If

    ' Error handler
    Exit Sub
errorhandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        GoTo restartmenuloop
    ElseIf Options.Debug >= 1 Then
        HandleError "MenuLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
    End If
End Sub

Sub ProcessWeather()
Dim i As Long
Static WeatherSound As Long
    If CurrentWeather > 0 Then
        If WeatherSound < GetTickCount Then
            If CurrentWeather = WEATHER_TYPE_RAIN Then
                If CurrentWeatherIntensity < 50 Then
                    Call PlaySound("som Chuva Fraca.mp3", -1, -1)
                    WeatherSound = GetTickCount + 20000
                End If
                If CurrentWeatherIntensity >= 50 And CurrentWeatherIntensity < 80 Then
                    Call PlaySound("som Chuva Media Vento.mp3", -1, -1)
                    WeatherSound = GetTickCount + 21000
                End If
                If CurrentWeatherIntensity >= 80 Then
                    Call PlaySound("som Chuva Forte.mp3", -1, -1)
                    WeatherSound = GetTickCount + 6000
                End If
            End If
            If CurrentWeather = WEATHER_TYPE_STORM Then
                Call PlaySound("som Chuva Tempestade.mp3", -1, -1)
                WeatherSound = GetTickCount + 24000
            End If
        End If
        i = Rand(1, 101 - CurrentWeatherIntensity)
        If i = 1 Then
            'Add a new particle
            For i = 1 To MAX_WEATHER_PARTICLES
                If WeatherParticle(i).InUse = False Then
                        If Rand(1, 2) = 1 Then
                            WeatherParticle(i).InUse = True
                            WeatherParticle(i).Type = CurrentWeather
                            WeatherParticle(i).Velocity = Rand(8, 14)
                            WeatherParticle(i).Size = Rand(16, 32)
                            WeatherParticle(i).X = (TileView.Left * 32) - 32
                            WeatherParticle(i).Y = (TileView.Top * 32) + Rand(-32, frmMain.ScaleHeight)
                        Else
                            WeatherParticle(i).InUse = True
                            WeatherParticle(i).Type = CurrentWeather
                            WeatherParticle(i).Velocity = Rand(10, 15)
                            WeatherParticle(i).Size = Rand(16, 32)
                            WeatherParticle(i).X = (TileView.Left * 32) + Rand(-32, frmMain.ScaleWidth)
                            WeatherParticle(i).Y = (TileView.Top * 32) - 32
                        End If
                    Exit For
                End If
                If Splash(i).Tick + 1000 < GetTickCount Then
                    Splash(i).X = Rand((TileView.Left * 32), (TileView.Right * 32))
                    Splash(i).Y = Rand((TileView.Top * 32), (TileView.Bottom * 32))
                    Splash(i).Tick = GetTickCount
                End If
            Next
        End If
    End If
    
    If CurrentWeather = WEATHER_TYPE_STORM Then
        i = Rand(1, 400 - CurrentWeatherIntensity)
        If i = 1 Then
            'Draw Thunder
            DrawThunder = Rand(15, 22)
            PlaySound Sound_Thunder, -1, -1
        End If
    End If
    
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).X > TileView.Right * 32 Or WeatherParticle(i).Y > TileView.Bottom * 32 Then
                WeatherParticle(i).InUse = False
            Else
                WeatherParticle(i).X = WeatherParticle(i).X + WeatherParticle(i).Velocity
                WeatherParticle(i).Y = WeatherParticle(i).Y + WeatherParticle(i).Velocity
            End If
        End If
    Next
End Sub

Public Sub AddChatBubble(ByVal Target As Long, ByVal TargetType As Byte, ByVal Msg As String, ByVal colour As Long)
Dim i As Long, Index As Long

    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    
    ' default to new bubble
    Index = chatBubbleIndex
    
    ' loop through and see if that player/npc already has a chat bubble
    For i = 1 To MAX_BYTE
        If chatBubble(i).TargetType = TargetType Then
            If chatBubble(i).Target = Target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                Index = i
                Exit For
            End If
        End If
    Next
    
    ' set the bubble up
    With chatBubble(Index)
        .Target = Target
        .TargetType = TargetType
        .Msg = Msg
        .colour = colour
        .Timer = GetTickCount
        .active = True
        .Alpha = 255
    End With
End Sub

Public Function IsBankItem(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    IsBankItem = 0
    
    For i = 1 To MAX_BANK
        If Not emptySlot Then
            If GetBankItemNum(i) <= 0 And GetBankItemNum(i) > MAX_ITEMS Then skipThis = True
        End If
        
        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_BANK).Y + BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_BANK).X + BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    
                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBankItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Public Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
Dim i As Long, Top As Long, Left As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    IsShopItem = 0

    For i = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(i).Item > 0 And Shop(InShop).TradeItem(i).Item <= MAX_ITEMS Then
            Top = GUIWindow(GUI_SHOP).Y + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            Left = GUIWindow(GUI_SHOP).X + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))

            If X >= Left And X <= Left + 32 Then
                If Y >= Top And Y <= Top + 32 Then
                    IsShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsShopItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsEqItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    IsEqItem = 0

    For i = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, i) > 0 And GetPlayerEquipment(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .Top = GUIWindow(GUI_CHARACTER).Y + EqTop
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_CHARACTER).X + EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsEqItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsInvItem(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    IsInvItem = 0

    For i = 1 To MAX_INV
        
        If Not emptySlot Then
            If GetPlayerInvItemNum(MyIndex, i) <= 0 Or GetPlayerInvItemNum(MyIndex, i) > MAX_ITEMS Then skipThis = True
        End If

        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_INVENTORY).Y + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_INVENTORY).X + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsPlayerEvoluteSpell(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    IsPlayerEvoluteSpell = 0
    
    Dim z As Long
    z = 0
    For i = 1 To MAX_SPELLS

        If Not ShowSpell(i) Then skipThis = True

        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_SPELLS).Y + 160 + SpellTop + ((SpellOffsetY + 32) * ((z) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((z) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsPlayerEvoluteSpell = i
                    Exit Function
                End If
            End If
            z = z + 1
        End If
        
        skipThis = False
    Next


    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS

        If Not emptySlot Then
            If PlayerSpells(i) <= 0 And PlayerSpells(i) > MAX_PLAYER_SPELLS Then skipThis = True
        End If

        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_SPELLS).Y + SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean, Optional ByVal emptySlot As Boolean = False) As Long
    Dim tempRec As RECT, skipThis As Boolean
    Dim i As Long
    Dim IsTradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    IsTradeItem = 0

    For i = 1 To MAX_INV
    
        If Yours Then
            IsTradeNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        Else
            IsTradeNum = TradeTheirOffer(i).num
        End If
        
        If Not emptySlot Then
            If IsTradeNum <= 0 Or IsTradeNum > MAX_ITEMS Then skipThis = True
        End If
        
        If Not skipThis Then
            If Yours Then
             With tempRec
                .Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_TRADE).X + 29 + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsTradeItem = i
                    Exit Function
                End If
            End If
            Else
            
            tempRec.Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop - 2 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
            tempRec.Bottom = tempRec.Top + PIC_Y
            tempRec.Left = GUIWindow(GUI_TRADE).X + 257 + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
            tempRec.Right = tempRec.Left + PIC_X
            
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsTradeItem = i
                    Exit Function
                End If
            End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTradeItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Public Function IsHotbarSlot(ByVal X As Single, ByVal Y As Single) As Long
Dim Top As Long, Left As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    IsHotbarSlot = 0

    For i = 1 To MAX_HOTBAR
        Top = GUIWindow(GUI_HOTBAR).Y + HotbarTop
        Left = GUIWindow(GUI_HOTBAR).X + HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
        If X >= Left And X <= Left + PIC_X Then
            If Y >= Top And Y <= Top + PIC_Y Then
                IsHotbarSlot = i
                Exit Function
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CensorWord(ByVal sString As String) As String
    CensorWord = String(Len(sString), "*")
End Function
Public Sub setOptionsState()
    ' music
    If Options.Music = 1 Then
        Buttons(26).State = 2
        Buttons(27).State = 0
    Else
        Buttons(26).State = 0
        Buttons(27).State = 2
    End If
    
    ' sound
    If Options.Sound = 1 Then
        Buttons(28).State = 2
        Buttons(29).State = 0
    Else
        Buttons(28).State = 0
        Buttons(29).State = 2
    End If
    
    ' debug
    If Options.Debug >= 1 Then
        Buttons(30).State = 2
        Buttons(31).State = 0
    Else
        Buttons(30).State = 0
        Buttons(31).State = 2
    End If
    
    ' ambiente
    If Options.Ambiente = 1 Then
        Buttons(46).State = 2
        Buttons(47).State = 0
    Else
        Buttons(46).State = 0
        Buttons(47).State = 2
    End If
    
    ' tela
    If Options.Tela Then
        Buttons(48).State = 2
        Buttons(49).State = 0
    Else
        Buttons(48).State = 0
        Buttons(49).State = 2
    End If
    
    ' clima
    If Options.Clima = 1 Then
        Buttons(50).State = 2
        Buttons(51).State = 0
    Else
        Buttons(50).State = 0
        Buttons(51).State = 2
    End If
    
    ' neblina
    If Options.Neblina = 1 Then
        Buttons(52).State = 2
        Buttons(53).State = 0
    Else
        Buttons(52).State = 0
        Buttons(53).State = 2
    End If
End Sub

Public Sub ScrollChatBox(ByVal Direction As Byte)
    ' do a quick exit if we don't have enough text to scroll
    If totalChatLines < 8 Then
        ChatScroll = 8
        UpdateChatArray
        Exit Sub
    End If
    ' actually scroll
    If Direction = 0 Then ' up
        ChatScroll = ChatScroll + 1
    Else ' down
        ChatScroll = ChatScroll - 1
    End If
    ' scrolling down
    If ChatScroll < 8 Then ChatScroll = 8
    ' scrolling up
    If ChatScroll > totalChatLines Then ChatScroll = totalChatLines
    ' update the array
    UpdateChatArray
End Sub

Public Sub SetBarWidth(ByRef MaxWidth As Long, ByRef Width As Long)
Dim barDifference As Long
    If MaxWidth < Width Then
        ' find out the amount to increase per loop
        barDifference = ((Width - MaxWidth) / 100) * 10
        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        Width = Width - barDifference
    ElseIf MaxWidth > Width Then
        ' find out the amount to increase per loop
        barDifference = ((MaxWidth - Width) / 100) * 10
        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        Width = Width + barDifference
    End If
End Sub

Public Sub CreateProjectile(ByVal AttackerIndex As Integer, ByVal TargetIndex As Integer, ByVal TargetType As Integer, ByVal Graphic As Long, ByVal Rotate As Long, ByVal RotateSpeed As Byte, ByVal NPCAttack As Byte)
Dim ProjectileIndex As Integer

    If AttackerIndex = 0 Then Exit Sub
    If TargetIndex = 0 Then Exit Sub
    If AttackerIndex > MAX_PLAYERS Then Exit Sub
    If TargetIndex > MAX_PLAYERS Then Exit Sub

    'Get the next open projectile slot
    Do
        ProjectileIndex = ProjectileIndex + 1
        
        'Update LastProjectile if we go over the size of the current array
        If ProjectileIndex > LastProjectile Then
            LastProjectile = ProjectileIndex
            ReDim Preserve ProjectileList(1 To LastProjectile)
            Exit Do
        End If
        
    Loop While ProjectileList(ProjectileIndex).Graphic > 0
    
    With ProjectileList(ProjectileIndex)
    
        ' ****** Initial Rotation Value ******
        .Rotate = Rotate
        
        ' ****** Set Values ******
        .Graphic = Graphic
        .RotateSpeed = RotateSpeed
        
        If NPCAttack = 0 Then
            ' ****** Get Target Type ******
            Select Case TargetType
                Case TARGET_TYPE_PLAYER
                    .tx = Player(TargetIndex).X * PIC_X
                    .ty = Player(TargetIndex).Y * PIC_Y
                    .X = Player(AttackerIndex).X * PIC_X
                    .Y = Player(AttackerIndex).Y * PIC_Y
                Case TARGET_TYPE_NPC
                    .tx = MapNpc(TargetIndex).X * PIC_X
                    .ty = MapNpc(TargetIndex).Y * PIC_Y
                    .X = GetPlayerX(AttackerIndex) * PIC_X
                    .Y = GetPlayerY(AttackerIndex) * PIC_Y
            End Select
        Else
            .tx = Player(AttackerIndex).X * PIC_X
            .ty = Player(AttackerIndex).Y * PIC_Y
            .X = MapNpc(TargetIndex).X * PIC_X
            .Y = MapNpc(TargetIndex).Y * PIC_Y
        End If
        
    End With
    
End Sub

Public Sub ClearProjectile(ByVal ProjectileIndex As Integer)
 
    'Clear the selected index
    ProjectileList(ProjectileIndex).Graphic = 0
    ProjectileList(ProjectileIndex).X = 0
    ProjectileList(ProjectileIndex).Y = 0
    ProjectileList(ProjectileIndex).tx = 0
    ProjectileList(ProjectileIndex).ty = 0
    ProjectileList(ProjectileIndex).Rotate = 0
    ProjectileList(ProjectileIndex).RotateSpeed = 0
 
    'Update LastProjectile
    If ProjectileIndex = LastProjectile Then
        Do Until ProjectileList(ProjectileIndex).Graphic > 1
            'Move down one projectile
            LastProjectile = LastProjectile - 1
            If LastProjectile = 0 Then Exit Do
        Loop
        If ProjectileIndex <> LastProjectile Then
            'We still have projectiles, resize the array to end at the last used slot
            If LastProjectile > 0 Then
                ReDim Preserve ProjectileList(1 To LastProjectile)
            Else
                Erase ProjectileList
            End If
        End If
    End If
 
End Sub

Public Function GetComparisonOperatorName(ByVal opr As ComparisonOperator) As String
    Select Case opr
        Case GEQUAL
            GetComparisonOperatorName = ">="
            Exit Function
        Case LEQUAL
            GetComparisonOperatorName = "<="
            Exit Function
        Case GREATER
            GetComparisonOperatorName = ">"
            Exit Function
        Case LESS
            GetComparisonOperatorName = "<"
            Exit Function
        Case EQUAL
            GetComparisonOperatorName = "="
            Exit Function
        Case NOTEQUAL
            GetComparisonOperatorName = "><"
            Exit Function
    End Select
    GetComparisonOperatorName = "Unknown"
End Function

Public Function GetEventTypeName(ByVal EventIndex As Long, SubIndex As Long) As String
Dim evtType As EventType
evtType = Events(EventIndex).SubEvents(SubIndex).Type
    Select Case evtType
        Case Evt_Message
            GetEventTypeName = "@Show Message: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_Menu
            GetEventTypeName = "@Show Choices"
            Exit Function
        Case Evt_Quit
            GetEventTypeName = "@Exit Event"
            Exit Function
        Case Evt_OpenShop
            If Events(EventIndex).SubEvents(SubIndex).data(1) > 0 Then
                GetEventTypeName = "@Open Shop: " & Events(EventIndex).SubEvents(SubIndex).data(1) & "-" & Trim$(Shop(Events(EventIndex).SubEvents(SubIndex).data(1)).name)
            Else
                GetEventTypeName = "@Open Shop: " & Events(EventIndex).SubEvents(SubIndex).data(1) & "- None "
            End If
            Exit Function
        Case Evt_OpenBank
            GetEventTypeName = "@Open Bank"
            Exit Function
        Case Evt_GiveItem
            GetEventTypeName = "@Change Item"
            Exit Function
        Case Evt_ChangeLevel
            GetEventTypeName = "@Change Level"
            Exit Function
        Case Evt_PlayAnimation
            If Events(EventIndex).SubEvents(SubIndex).data(1) > 0 Then
                GetEventTypeName = "@Play Animation: " & Events(EventIndex).SubEvents(SubIndex).data(1) & "." & Trim$(Animation(Events(EventIndex).SubEvents(SubIndex).data(1)).name) & " {" & Events(EventIndex).SubEvents(SubIndex).data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).data(3) & "}"
            Else
                GetEventTypeName = "@Play Animation: None {" & Events(EventIndex).SubEvents(SubIndex).data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).data(3) & "}"
            End If
            Exit Function
        Case Evt_Warp
            GetEventTypeName = "@Warp to: " & Events(EventIndex).SubEvents(SubIndex).data(1) & " {" & Events(EventIndex).SubEvents(SubIndex).data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).data(3) & "}"
            Exit Function
        Case Evt_GOTO
            GetEventTypeName = "@GoTo: " & Events(EventIndex).SubEvents(SubIndex).data(1)
            Exit Function
        Case Evt_Switch
            If Events(EventIndex).SubEvents(SubIndex).data(2) = 1 Then
                GetEventTypeName = "@Change Switch: " & Events(EventIndex).SubEvents(SubIndex).data(1) + 1 & "." & Switches(Events(EventIndex).SubEvents(SubIndex).data(1) + 1) & " = True"
            Else
                GetEventTypeName = "@Change Switch: " & Events(EventIndex).SubEvents(SubIndex).data(1) + 1 & "." & Switches(Events(EventIndex).SubEvents(SubIndex).data(1) + 1) & " = False"
            End If
            Exit Function
        Case Evt_Variable
            GetEventTypeName = "@Change Variable: "
            Exit Function
        Case Evt_AddText
            Select Case Events(EventIndex).SubEvents(SubIndex).data(2)
                Case 0: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).data(1)) & ", Player}"
                Case 1: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).data(1)) & ", Map}"
                Case 2: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).data(1)) & ", Global}"
            End Select
            Exit Function
        Case Evt_Chatbubble
            GetEventTypeName = "@Show chatbubble"
            Exit Function
        Case Evt_Branch
            GetEventTypeName = "@Conditional branch"
            Exit Function
        Case Evt_ChangeSkill
            GetEventTypeName = "@Change Spells"
            Exit Function
        Case Evt_ChangeSprite
            GetEventTypeName = "@Change Sprite: " & Events(EventIndex).SubEvents(SubIndex).data(1)
            Exit Function
        Case Evt_ChangePK
            Select Case Events(EventIndex).SubEvents(SubIndex).data(1)
                Case 0: GetEventTypeName = "@Change PK: NO"
                Case 1: GetEventTypeName = "@Change PK: YES"
            End Select
            Exit Function
        Case Evt_SpawnNPC
            If Events(EventIndex).SubEvents(SubIndex).data(1) > 0 Then
                GetEventTypeName = "@Spawn NPC: " & Trim$(Npc(Map.Npc(Events(EventIndex).SubEvents(SubIndex).data(1))).name)
            Else
                GetEventTypeName = "@Spawn NPC: None"
            End If
            Exit Function
        Case Evt_ChangeClass
            If Events(EventIndex).SubEvents(SubIndex).data(1) > 0 Then
                GetEventTypeName = "@Change Class: " & Trim$(Class(Events(EventIndex).SubEvents(SubIndex).data(1)).name)
            Else
                GetEventTypeName = "@Change Class: None"
            End If
            Exit Function
        Case Evt_ChangeSex
            Select Case Events(EventIndex).SubEvents(SubIndex).data(1)
                Case 0: GetEventTypeName = "@Change Sex: MALE"
                Case 1: GetEventTypeName = "@Change Sex: FEMALE"
            End Select
            Exit Function
        Case Evt_ChangeExp
            GetEventTypeName = "@Change Exp"
            Exit Function
        Case Evt_SpecialEffect
            GetEventTypeName = "@Special Effect"
            Exit Function
        Case Evt_PlaySound
            GetEventTypeName = "@Play Sound: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_PlayBGM
            GetEventTypeName = "@Play BGM: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_StopSound
            GetEventTypeName = "@Stop Sound"
            Exit Function
        Case Evt_FadeoutBGM
            GetEventTypeName = "@Fadeout BGM"
            Exit Function
        Case Evt_SetAccess
            GetEventTypeName = "@Set Access: " & Events(EventIndex).SubEvents(SubIndex).data(1)
            Exit Function
        Case Evt_CustomScript
            GetEventTypeName = "@Custom Script: " & Events(EventIndex).SubEvents(SubIndex).data(1)
            Exit Function
        Case Evt_OpenEvent
            Select Case Events(EventIndex).SubEvents(SubIndex).data(3)
                Case 0: GetEventTypeName = "@Open Event: {" & Events(EventIndex).SubEvents(SubIndex).data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).data(2) & "}"
                Case 1: GetEventTypeName = "@Close Event: {" & Events(EventIndex).SubEvents(SubIndex).data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).data(2) & "}"
            End Select
            Exit Function
    End Select
    GetEventTypeName = "Unknown"
End Function

' *****************
' ** Event Logic **
' *****************
Public Sub Events_SetSubEventType(ByVal EIndex As Long, ByVal SIndex As Long, ByVal EType As EventType)
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    If SIndex < LBound(Events(EIndex).SubEvents) Or SIndex > UBound(Events(EIndex).SubEvents) Then Exit Sub
    
    'We are ok, allocate
    With Events(EIndex).SubEvents(SIndex)
        .Type = EType
        Select Case .Type
            Case Evt_Message
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .data(1 To 1)
            Case Evt_Menu
                If Not .HasText Then ReDim .Text(1 To 2)
                If UBound(.Text) < 2 Then ReDim Preserve .Text(1 To 2)
                If Not .HasData Then ReDim .data(1 To 1)
                .HasText = True
                .HasData = True
            Case Evt_OpenShop
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_GOTO
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_GiveItem
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 3)
            Case Evt_PlayAnimation
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 3)
            Case Evt_Warp
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 3)
            Case Evt_Switch
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 2)
            Case Evt_Variable
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 4)
            Case Evt_AddText
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .data(1 To 2)
            Case Evt_Chatbubble
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .data(1 To 2)
            Case Evt_Branch
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 6)
            Case Evt_ChangeSkill
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 2)
            Case Evt_ChangeLevel
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 2)
            Case Evt_ChangeSprite
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_ChangePK
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_SpawnNPC
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_ChangeClass
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_ChangeSex
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_ChangeExp
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 2)
            Case Evt_SpecialEffect
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 5)
            Case Evt_PlaySound
                .HasText = True
                .HasData = False
                Erase .data
                ReDim Preserve .Text(1 To 1)
            Case Evt_PlayBGM
                .HasText = True
                .HasData = False
                Erase .data
                ReDim Preserve .Text(1 To 1)
            Case Evt_SetAccess
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_CustomScript
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 1)
            Case Evt_OpenEvent
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .data(1 To 4)
            Case Else
                .HasText = False
                .HasData = False
                Erase .Text
                Erase .data
        End Select
    End With
End Sub

Public Sub CastEffect(ByVal EffectNum As Long, X As Long, Y As Long)
    'X = X * PIC_X
    'Y = Y * PIC_Y
    Select Case Effect(EffectNum).Type
        Case EFFECT_TYPE_HEAL
        X = X * PIC_X + (PIC_X / 2)
        Y = Y * PIC_Y - (PIC_Y / 2)
        Call Heal_Begin(EffectNum, X, Y)
        
        Case EFFECT_TYPE_PROTECTION
        X = X * PIC_X + (PIC_X / 2)
        Y = Y * PIC_Y + (PIC_Y / 2)
        Call Protection_Begin(EffectNum, X, Y)
        
        Case EFFECT_TYPE_STRENGTHEN
        X = X * PIC_X + (PIC_X / 2)
        Y = Y * PIC_Y + (PIC_Y / 2)
        Call Strengthen_Begin(EffectNum, X, Y)
        
        Case EFFECT_TYPE_SUMMON
        X = X * PIC_X + (PIC_X / 2)
        Y = Y * PIC_Y + (PIC_Y / 2)
        Call Summon_Begin(EffectNum, X, Y)
    End Select
End Sub

Sub ClearEffect(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Effect(Index)), LenB(Effect(Index)))
    Effect(Index).name = vbNullString
    Effect(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearEffect", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearEffects()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_EFFECTS
        Call ClearEffect(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearEffects", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub FindNearestTarget()
Dim i As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, xDif As Long, yDif As Long
Dim bestX As Long, bestY As Long, bestIndex As Long

    X2 = GetPlayerX(MyIndex)
    Y2 = GetPlayerY(MyIndex)
    
    bestX = 255
    bestY = 255
    
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).num > 0 Then
            X = MapNpc(i).X
            Y = MapNpc(i).Y
            ' find the difference - x
            If X < X2 Then
                xDif = X2 - X
            ElseIf X > X2 Then
                xDif = X - X2
            Else
                xDif = 0
            End If
            ' find the difference - y
            If Y < Y2 Then
                yDif = Y2 - Y
            ElseIf Y > Y2 Then
                yDif = Y - Y2
            Else
                yDif = 0
            End If
            ' best so far?
            If (xDif + yDif) < (bestX + bestY) Then
                bestX = xDif
                bestY = yDif
                bestIndex = i
            End If
        End If
    Next
    
    ' target the best
    If bestIndex > 0 And bestIndex <> myTarget Then PlayerTarget bestIndex, TARGET_TYPE_NPC
End Sub

Public Sub FindTarget()
Dim i As Long, X As Long, Y As Long

    ' check players
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            X = (GetPlayerX(i) * 32) + TempPlayer(i).XOffSet + 32
            Y = (GetPlayerY(i) * 32) + TempPlayer(i).YOffSet + 32
            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget i, TARGET_TYPE_PLAYER
                    Exit Sub
                End If
            End If
        End If
    Next
    
    ' check npcs
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).num > 0 Then
            X = (MapNpc(i).X * 32) + TempMapNpc(i).XOffSet + 32
            Y = (MapNpc(i).Y * 32) + TempMapNpc(i).YOffSet + 32
            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget i, TARGET_TYPE_NPC
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

Sub DoAnimation(Animation As Long, X As Long, Y As Long, LockType As Byte, lockindex As Long, ByVal Dir As Long, Optional ReturnAnim As Long = 0)
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = Animation
        .X = X
        .Y = Y
        .LockType = LockType
        .lockindex = lockindex
        .Used(0) = False
        .Used(1) = True
        .Dir = Dir
        .ReturnAnim = ReturnAnim
    End With
    
    PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation, AnimInstance(AnimationIndex).CastAnim
End Sub

Sub FazerBuraco(X As Long, Y As Long, Size As Long)
Dim i As Long
    
    For i = 1 To 10
        If Buracos(i).InUse = True Then
            If Buracos(i).X = X And Buracos(i).Y = Y And Buracos(i).Map = GetPlayerMap(MyIndex) Then
                If Buracos(i).Size <= Size Then
                    Buracos(i).Alpha = 255
                    Buracos(i).Size = Size
                    Buracos(i).IntervalTick = GetTickCount + 5000
                End If
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To 10
        If Buracos(i).InUse = False Then
            Buracos(i).InUse = True
            Buracos(i).Alpha = 255
            Buracos(i).Size = Size
            Buracos(i).X = X
            Buracos(i).Y = Y
            Buracos(i).IntervalTick = GetTickCount + 5000
            Buracos(i).Map = GetPlayerMap(MyIndex)
            Exit Sub
        End If
    Next i
End Sub

Sub DoInGame()
    InGame = True
    Call SendDevSuite
    Call GameInit
    Call GameLoop
End Sub

Public Function IsPlayerQuest(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim i As Long
Dim ActualQuest As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    IsPlayerQuest = 0

    For i = 1 To MAX_QUESTS

        If Not Player(MyIndex).QuestState(i).State = 1 Then
            skipThis = True
        End If

        If Not skipThis Then
            ActualQuest = ActualQuest + 1
            With tempRec
                .Top = GUIWindow(GUI_QUESTS).Y + SpellTop + ((SpellOffsetY + 32) * ((ActualQuest - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_QUESTS).X + SpellLeft + ((SpellOffsetX + 32) * (((ActualQuest - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsPlayerQuest = i
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Sub CheckAFK()
    On Error Resume Next
    Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).LastMove + AFKTime < GetTickCount Or Map.Tile(Player(i).X, Player(i).Y).Type = TILE_TYPE_SHOP Or Map.Tile(Player(i).X, Player(i).Y).Type = TILE_TYPE_EVENT Then
                If TempPlayer(i).AFK = 0 Then TempPlayer(i).AFK = 1
            End If
        End If
    Next i
End Sub


Function GetPlayerStatNextLevel(Index, ByVal Stat As Stats)
    GetPlayerStatNextLevel = StatNextLevel(Stat)
End Function

Function GetPlayerStatPrevLevel(Index, ByVal Stat As Stats)
    GetPlayerStatPrevLevel = StatLastLevel(Stat)
End Function

Function WALK_SPEED(ByVal Index As Long) As Double
    WALK_SPEED = 4
End Function

Function RUN_SPEED(ByVal Index As Long) As Double
    RUN_SPEED = 8
    If TempPlayer(Index).Fly > 0 Then RUN_SPEED = RUN_SPEED + 2
End Function

Function NomeEspecie(ByVal Especie As Byte) As String
    Select Case Especie
        Case 0: NomeEspecie = "???"
        Case 1: NomeEspecie = "Humanóides"
        Case 2: NomeEspecie = "Insetóides"
        Case 3: NomeEspecie = "Ferais"
        Case 4: NomeEspecie = "Amaldiçoados"
        Case 5: NomeEspecie = "Robóticos"
    End Select
End Function

Function GravityValue(ByVal Gravity As Long) As Long
    If Gravity < 10 Then Gravity = 10
    If Gravity > 1500 Then Gravity = 1500
    If Gravity > Int(10 + (GetPlayerLevel(MyIndex) / 95) * 1500) Then Gravity = Int(10 + (GetPlayerLevel(MyIndex) / 95) * 1500)
    GravityValue = Int(((Gravity / 1500) * 50000) + 5000)
End Function

Function MaxGravity() As Long
    MaxGravity = Int(10 + (GetPlayerLevel(MyIndex) / 95) * 1500)
End Function

Function InTutorial() As Boolean
    
    InTutorial = False
    If InGame Then
    If Player(MyIndex).InTutorial = 0 Then InTutorial = True
    End If
End Function

Sub SetTarget(ByVal NPCName As String)
    Dim i As Long
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).num > 0 Then
            If LCase(Trim$(Npc(MapNpc(i).num).name)) = LCase(NPCName) Then
                myTarget = i
                myTargetType = TARGET_TYPE_NPC
            End If
        End If
    Next i
End Sub

Sub HandleTutorialClick()
    If HideContinue Then Exit Sub
    Static LastClick As Long
    If LastClick + TutorialWaitTime < GetTickCount Then
        If Player(MyIndex).InTutorial = 0 Then
            TutorialProgress = 0
            TutorialStep = TutorialStep + 1
            PlaySound Sound_ButtonHover, -1, -1
            If TutorialStep = 2 Then
                GUIWindow(GUI_INVENTORY).visible = True
                TutorialShowIcon = 1
            End If
            If TutorialStep = 3 Then
                GUIWindow(GUI_INVENTORY).visible = False
                GUIWindow(GUI_SPELLS).visible = True
                TutorialShowIcon = 2
            End If
            If TutorialStep = 4 Then
                GUIWindow(GUI_SPELLS).visible = False
                GUIWindow(GUI_CHARACTER).visible = True
                TutorialShowIcon = 3
            End If
            If TutorialStep = 6 Then
                GUIWindow(GUI_CHARACTER).visible = False
                TutorialShowIcon = 0
                chatOn = True
            End If
            If TutorialStep = 18 Then
                Player(MyIndex).X = 18
                Player(MyIndex).Y = 14
                SetTarget "Rei Vegeta"
            End If
            If TutorialStep = 20 Then
                SetTarget "Vegeta"
            End If
            If TutorialStep = 21 Then
                SetTarget "Bardock"
            End If
            If TutorialStep = 22 Then
                Player(MyIndex).X = 3
                Player(MyIndex).Y = 8
                SetTarget "Registro de mérito"
            End If
            If TutorialStep = 23 Then
                SetTarget "Missões"
            End If
            If TutorialStep = 24 Then
                Player(MyIndex).X = 23
                Player(MyIndex).Y = 24
                SetTarget "Extratores"
            End If
            If TutorialStep = 25 Then
                Player(MyIndex).X = 3
                Player(MyIndex).Y = 21
            End If
            If TutorialStep = 26 Then
                Player(MyIndex).X = 15
                Player(MyIndex).Y = 15
            End If
            If TutorialStep = 45 Then
                SendCompleteTutorial
                Player(MyIndex).InTutorial = 1
                GUIWindow(GUI_INVENTORY).visible = True
            End If
        End If
        LastClick = GetTickCount
    Else
        Call PlaySound("Drop.wav", -1, -1)
    End If
End Sub
Public Function GetPlayerGuildIndex(ByVal Index As Long) As Long
    Dim GuildNum As Long
    Dim i As Long
    
    GuildNum = Player(Index).Guild
    
    If GuildNum > 0 Then
        For i = 1 To 10
            If LCase(Trim$(Guild(GuildNum).Member(i).name)) = LCase(Trim$(GetPlayerName(Index))) Then
                GetPlayerGuildIndex = i
                Exit Function
            End If
        Next i
    End If
End Function
Public Function GetPlayerGuildRank(ByVal Index As Long) As Byte
    Dim GuildNum As Long
    Dim GuildIndex As Long
    GuildNum = Player(Index).Guild
    If GuildNum > 0 Then
        GuildIndex = GetPlayerGuildIndex(Index)
        If GuildIndex > 0 Then
            GetPlayerGuildRank = Guild(GuildNum).Member(GuildIndex).Rank
        End If
    End If
End Function
Public Function RankName(ByVal Rank As Long) As String
    Select Case Rank
        Case 0: RankName = "Membro"
        Case 1: RankName = "Capitão"
        Case 2: RankName = "Major"
        Case 3: RankName = "Mestre"
    End Select
End Function
Function MaxGuildMembers(Level As Long) As Byte
    MaxGuildMembers = 3 + Int(Level / 5)
    If MaxGuildMembers > 10 Then MaxGuildMembers = 10
End Function
Public Sub ShowGuildPanel()
    Dim GuildNum As Long
    Dim MyRank As Long
    
    GuildNum = Player(MyIndex).Guild
    If GuildNum = 0 Then
        AddText "Você não está em nenhuma guild!", BrightRed
        Exit Sub
    End If
    
    frmMain.picGuildAdmin.visible = True
    
    Dim PlayerIndex As Long
    Dim MemberCount As Long
    Dim i As Long
    For i = 1 To 10
        If LCase(Trim$(Guild(GuildNum).Member(i).name)) = LCase(Trim$(GetPlayerName(MyIndex))) Then
            PlayerIndex = i
        End If
        If Guild(GuildNum).Member(i).Level > 0 Then
            MemberCount = MemberCount + 1
        End If
    Next i
    
    MyRank = Guild(GuildNum).Member(PlayerIndex).Rank
    
    'Draw Icon
    For i = 0 To 24
        frmMain.picIcon(i).BackColor = QBColor(Guild(GuildNum).IconColor(i + 1))
    Next i
    
    frmMain.lblGuildName.Caption = Trim$(Guild(GuildNum).name)
    frmMain.lblGuildMOTD.Caption = Trim$(Guild(GuildNum).MOTD)
    frmMain.lblGuildLevel.Caption = "Guild Level: " & Guild(GuildNum).Level
    frmMain.lblGuildMembers.Caption = "Membros: " & MemberCount & "/" & MaxGuildMembers(Guild(GuildNum).Level)
    frmMain.picGuildExpBar.Width = frmMain.picGuildExpBarMold.Width * (Guild(GuildNum).EXP / Guild(GuildNum).TNL)
    frmMain.chkBlock.value = Guild(GuildNum).UpBlock
    
    For i = 0 To 2
        frmMain.frmActions(i).visible = False
    Next i
    If MyRank > 0 Then
        frmMain.frmActions(MyRank - 1).visible = True
    End If
    If MyRank = 3 Then frmMain.cmdMOTD.visible = True
    
    frmMain.lblRed.Caption = "Especiaria vermelha: " & Guild(GuildNum).Red
    frmMain.lblBlue.Caption = "Especiaria azul: " & Guild(GuildNum).Blue
    frmMain.lblYellow.Caption = "Especiaria amarela: " & Guild(GuildNum).Yellow
    frmMain.lblGold.Caption = "Moedas Z: " & Guild(GuildNum).Gold
    
    frmMain.lstMembers.Clear
    For i = 1 To 10
        If Guild(GuildNum).Member(i).Level > 0 Then
            frmMain.lstMembers.AddItem "[" & UCase(RankName(Guild(GuildNum).Member(i).Rank)) & "] " & Trim$(Guild(GuildNum).Member(i).name)
        End If
    Next i
End Sub

Public Function Semana() As Integer
  Semana = Format(Now, "w") - 2
  If Semana < 0 Then Semana = 7 - Abs(Semana)
End Function

Sub OpenArenaDialog()
    frmMain.picArena.visible = True
    frmMain.cmbType.ListIndex = 0
    frmMain.cmbPlayers.ListIndex = 0
End Sub

Function InOwnPlanet() As Boolean
    If MAX_PLAYER_PLANETS > 0 Then
        Dim i As Long
        For i = 1 To MAX_PLAYER_PLANETS
            If Trim$(LCase(PlayerPlanet(i).PlanetData.Owner)) = Trim$(LCase(GetPlayerName(MyIndex))) Then
                If PlayerPlanet(i).PlanetData.Map = GetPlayerMap(MyIndex) Then
                    InOwnPlanet = True
                    Exit Function
                End If
            End If
        Next i
    End If
End Function

Function GetRemaining(Remaining As Long) As String
    Dim sString As String
    Dim AddHours As Long
    sString = (Remaining Mod 60) & "m"
    If Int(Remaining / 60) > 0 Then
        sString = Int(Int(Remaining / 60) Mod 24) & "h " & sString
    End If
    If Int(Remaining / 1440) > 0 Then sString = Int(Remaining / 1440) & "d " & sString
    GetRemaining = sString
End Function

Sub ShowFabrica()
    frmMain.picFabrica.visible = True
    frmMain.lstFila.Clear
End Sub

Sub PopConquista(ByVal ConquistaNum As Long)
    PopConquistaNum = ConquistaNum
    PopConquistaTick = GetTickCount
End Sub

Function CanShow(ByVal Index As Long) As Boolean
    If Map.Moral <> 3 Or Index = MyIndex Then
        CanShow = True
        Exit Function
    Else
        Dim i As Long
        For i = 1 To Party.MemberCount
            If Party.Member(i) = Index Then
                CanShow = True
                Exit Function
            End If
        Next i
    End If
End Function
