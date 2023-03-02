Attribute VB_Name = "modGameLogic"
Option Explicit
Public paperdollTestin As Boolean

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

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' *** Start GameLoop ***
    Do While InGame
        Tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.
        
        ' * Check surface timers *
        If surftmr < Tick Then
            For i = 1 To NumTextures
                UnsetTexture (i)
            Next
            surftmr = GetTickCount + 75000
        End If
        
        If Tremor > GetTickCount Then TremorX = Rand(-2, 2)
    
        If tmr10000 < Tick Then
            ' check ping
            Call GetPing
            tmr10000 = Tick + 10000
            If GetPlayerX(MyIndex) <= Map.MaxX And GetPlayerY(MyIndex) <= Map.MaxY Then
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_RESOURCE Then
                    If Resource(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Data1).ResourceType = 3 Then
                        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                            If Item(GetPlayerEquipment(MyIndex, Weapon)).Data3 = 2 Then
                                If FishingTime = 0 Then FishingTime = 1
                            End If
                        End If
                    End If
                Else
                    FishingTime = 0
                End If
            End If
        End If

        If tmr25 < Tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything
            
            
            On Error GoTo Jump
            If GetPlayerX(MyIndex) <= Map.MaxX And GetPlayerY(MyIndex) <= Map.MaxY Then
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_RESOURCE Then
                    If FishingTime = 1 And Resource(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Data1).ResourceType = 3 Then
                        If Rand(1, 500) = 1 Then
                            FishTime = GetTickCount + Rand(400, 750)
                        End If
                    End If
                End If
            End If
Jump:

            If GetForegroundWindow() = frmMain.hWnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if we need to end the CD icon
            If NumSpellIcons > 0 Then
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
            If TempPlayer(MyIndex).SpellBuffer > 0 Then
                If TempPlayer(MyIndex).SpellBufferTimer + (Spell(PlayerSpells(TempPlayer(MyIndex).SpellBuffer)).CastTime * 1000) < Tick Then
                    TempPlayer(MyIndex).SpellBuffer = 0
                    TempPlayer(MyIndex).SpellBufferTimer = 0
                    TempPlayer(MyIndex).SpellBufferNum = 0
                End If
            End If

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
                    DrawAnimatedInvItems
                    tmr100 = Tick + 100
                End If
            End If
            
            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next
            
            tmr25 = Tick + 25
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
                If MapNpc(i).Num > 0 Then
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
        If targettmr < Tick Then
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
        If bartmr < Tick Then
            SetBarWidth BarWidth_GuiHP_Max, BarWidth_GuiHP
            SetBarWidth BarWidth_GuiSP_Max, BarWidth_GuiSP
            SetBarWidth BarWidth_GuiEXP_Max, BarWidth_GuiEXP
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
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
        
        ' ****** Parallax X ******
        If ParallaxX = -800 Then
            ParallaxX = 0
        Else
            ParallaxX = ParallaxX - 1
        End If
        
        ' ****** Parallax Y ******
        If ParallaxY = 0 Then
            ParallaxY = -600
        Else
            ParallaxY = ParallaxY + 1
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
            
            tmr500 = Tick + 500
        End If
        
        ProcessWeather
        
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
        If rendertmr < Tick Then
            Call Render_Graphics
            Call UpdateSounds
            rendertmr = Tick + 25
        End If
        
        DoEvents

        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < Tick + Options.FPS
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
        StopMusic
        PlayMusic Options.MenuMusic
    Else
        ' Shutdown the game
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if player is walking, and if so process moving them over
    Select Case TempPlayer(Index).Moving
        Case MOVING_WALKING: MovementSpeed = ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
        Case MOVING_RUNNING: MovementSpeed = ((ElapsedTime / 1000) * (RUN_SPEED * SIZE_X))
        Case Else: Exit Sub
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            TempPlayer(Index).YOffset = TempPlayer(Index).YOffset - MovementSpeed
            If TempPlayer(Index).YOffset < 0 Then TempPlayer(Index).YOffset = 0
        Case DIR_DOWN
            TempPlayer(Index).YOffset = TempPlayer(Index).YOffset + MovementSpeed
            If TempPlayer(Index).YOffset > 0 Then TempPlayer(Index).YOffset = 0
        Case DIR_LEFT
            TempPlayer(Index).xOffset = TempPlayer(Index).xOffset - MovementSpeed
            If TempPlayer(Index).xOffset < 0 Then TempPlayer(Index).xOffset = 0
        Case DIR_RIGHT
            TempPlayer(Index).xOffset = TempPlayer(Index).xOffset + MovementSpeed
            If TempPlayer(Index).xOffset > 0 Then TempPlayer(Index).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If TempPlayer(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (TempPlayer(Index).xOffset >= 0) And (TempPlayer(Index).YOffset >= 0) Then
                TempPlayer(Index).Moving = 0
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
            If (TempPlayer(Index).xOffset <= 0) And (TempPlayer(Index).YOffset <= 0) Then
                TempPlayer(Index).Moving = 0
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).Num = 0 Then Exit Sub

    ' Check if NPC is walking, and if so process moving them over
    If TempMapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                TempMapNpc(MapNpcNum).YOffset = TempMapNpc(MapNpcNum).YOffset - ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).Num).speed * SIZE_X))
                If TempMapNpc(MapNpcNum).YOffset < 0 Then TempMapNpc(MapNpcNum).YOffset = 0
                
            Case DIR_DOWN
                TempMapNpc(MapNpcNum).YOffset = TempMapNpc(MapNpcNum).YOffset + ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).Num).speed * SIZE_X))
                If TempMapNpc(MapNpcNum).YOffset > 0 Then TempMapNpc(MapNpcNum).YOffset = 0
                
            Case DIR_LEFT
                TempMapNpc(MapNpcNum).xOffset = TempMapNpc(MapNpcNum).xOffset - ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).Num).speed * SIZE_X))
                If TempMapNpc(MapNpcNum).xOffset < 0 Then TempMapNpc(MapNpcNum).xOffset = 0
                
            Case DIR_RIGHT
                TempMapNpc(MapNpcNum).xOffset = TempMapNpc(MapNpcNum).xOffset + ((ElapsedTime / 1000) * (Npc(MapNpc(MapNpcNum).Num).speed * SIZE_X))
                If TempMapNpc(MapNpcNum).xOffset > 0 Then TempMapNpc(MapNpcNum).xOffset = 0
                
        End Select
    
        ' Check if completed walking over to the next tile
        If TempMapNpc(MapNpcNum).Moving > 0 Then
            If MapNpc(MapNpcNum).Dir = DIR_RIGHT Or MapNpc(MapNpcNum).Dir = DIR_DOWN Then
                If (TempMapNpc(MapNpcNum).xOffset >= 0) And (TempMapNpc(MapNpcNum).YOffset >= 0) Then
                    TempMapNpc(MapNpcNum).Moving = 0
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
                If (TempMapNpc(MapNpcNum).xOffset <= 0) And (TempMapNpc(MapNpcNum).YOffset <= 0) Then
                    TempMapNpc(MapNpcNum).Moving = 0
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
Dim Buffer As New clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer

    If GetTickCount > TempPlayer(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            TempPlayer(MyIndex).MapGetTimer = GetTickCount
            Buffer.WriteLong CMapGetItem
            SendData Buffer.ToArray()
        End If
    End If

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAttack()
Dim Buffer As clsBuffer
Dim AttackSpeed As Long, X As Long, Y As Long, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ControlDown Then
    
        If TempPlayer(MyIndex).SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If TempPlayer(MyIndex).StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            AttackSpeed = Item(GetPlayerEquipment(MyIndex, Weapon)).speed - (GetPlayerStat(MyIndex, Agility) * 10)
        Else
            AttackSpeed = 1000 - (GetPlayerStat(MyIndex, Agility) * 10)
        End If

        If AttackSpeed < 200 Then AttackSpeed = 200
        
        If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_RESOURCE Then Exit Sub
        
        If TempPlayer(MyIndex).AttackTimer + AttackSpeed < GetTickCount Then
            If TempPlayer(MyIndex).Attacking = 0 Then

                With TempPlayer(MyIndex)
                    '.Attacking = 1
                    '.AttackTimer = GetTickCount
                    '.AttackAnim = Rand(0, 1)
                End With
                
                Set Buffer = New clsBuffer
                    Buffer.WriteLong CAttack
                    SendData Buffer.ToArray()
                Set Buffer = Nothing

            End If
        End If
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If DirUp Or DirDown Or DirLeft Or DirRight Then
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If TempPlayer(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    If TempPlayer(MyIndex).KamehamehaLast + 300 > GetTickCount Or TempPlayer(MyIndex).SpiritBombLast + 300 > GetTickCount Then
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
    
    If GUIWindow(GUI_NEWS).visible = True Then
        GUIWindow(GUI_NEWS).visible = False
    End If

    d = GetPlayerDir(MyIndex)

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

Function CheckDirection(ByVal direction As Byte) As Boolean
Dim X As Long
Dim Y As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    CheckDirection = False
    If CollisionDisabled = True Or TempPlayer(MyIndex).Fly = 1 Then Exit Function
    ' check directional blocking
    If isDirBlocked(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case direction
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

    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        If Resource(Map.Tile(X, Y).Data1).ResourceType <> 3 Then
            CheckDirection = True
            Exit Function
        End If
    End If
    
    If Map.Tile(X, Y).Type = TILE_TYPE_EVENT Then
        If Map.Tile(X, Y).Data1 > 0 Then
            If Events(Map.Tile(X, Y).Data1).WalkThrought = NO Then
                If Player(MyIndex).EventOpen(Map.Tile(X, Y).Data1) = NO Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Check to see if a player is already on that tile
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If GetPlayerX(i) = X Then
                If GetPlayerY(i) = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next i

    ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If IsTryingToMove Then
        If CanMove Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                TempPlayer(MyIndex).Moving = MOVING_RUNNING
            Else
                TempPlayer(MyIndex).Moving = MOVING_WALKING
            End If

            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    TempPlayer(MyIndex).YOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    TempPlayer(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    TempPlayer(MyIndex).xOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    TempPlayer(MyIndex).xOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select

            If TempPlayer(MyIndex).xOffset = 0 Then
                If TempPlayer(MyIndex).YOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP And Not TempPlayer(MyIndex).Fly = 1 Then
                        GettingMap = True
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    If Item(PlayerInv(InventoryItemSelected).Num).Type = ITEM_TYPE_SCOUTER Then
        PlaySound "ligandoscouter.mp3", -1, -1
        ScouterOn = Not ScouterOn
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
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellSlot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(MyIndex).SpellBuffer = spellSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellSlot) > 0 Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CForgetSpell
        Buffer.WriteLong spellSlot
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal spellSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If SpellCD(spellSlot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    
    'Spell Cast
    If TempPlayer(MyIndex).SpellBuffer > 0 Then Exit Sub
    If TempPlayer(MyIndex).StunDuration > 0 Then Exit Sub
    
    If PlayerSpells(spellSlot) = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellSlot)).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(PlayerSpells(spellSlot)).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellSlot) > 0 Then
        If GetTickCount > TempPlayer(MyIndex).AttackTimer + 1000 Then
            If TempPlayer(MyIndex).Moving = 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong CCast
                Buffer.WriteLong spellSlot
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                'TempPlayer(MyIndex).SpellBuffer = spellSlot
                'TempPlayer(MyIndex).SpellBufferTimer = GetTickCount
                'TempPlayer(MyIndex).SpellBufferNum = PlayerSpells(spellSlot)
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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

Public Sub CreateActionMsg(ByVal message As String, ByVal color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .message = message
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsg(Index).message = vbNullString
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
Dim layer As Long
Dim FrameCount As Long
Dim lockindex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub
    
    For layer = 0 To 1
        If AnimInstance(Index).Used(layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(layer) = 0 Then AnimInstance(Index).frameIndex(layer) = 1
            If AnimInstance(Index).LoopIndex(layer) = 0 Then AnimInstance(Index).LoopIndex(layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).Timer(layer) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimInstance(Index).frameIndex(layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(layer) = AnimInstance(Index).LoopIndex(layer) + 1
                    If AnimInstance(Index).LoopIndex(layer) > Animation(AnimInstance(Index).Animation).LoopCount(layer) Then
                        AnimInstance(Index).Used(layer) = False
                    Else
                        AnimInstance(Index).frameIndex(layer) = 1
                    End If
                Else
                    AnimInstance(Index).frameIndex(layer) = AnimInstance(Index).frameIndex(layer) + 1
                End If
                AnimInstance(Index).Timer(layer) = GetTickCount
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    GetBankItemNum = Bank.Item(bankslot).Num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal itemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).Num = itemNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GetBankItemValue = Bank.Item(bankslot).Value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).Value = ItemValue
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim soundName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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

    ' play the sound
    PlaySound soundName, X, Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal Data1 As Long = 0)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' exit out if we've already got a dialogue open
    If dialogueIndex > 0 Then Exit Sub
    
    ' set global dialogue index
    dialogueIndex = diIndex
    
    ' set the global dialogue data
    dialogueData1 = Data1

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

Public Sub dialogueHandler(ByVal Index As Long)
    ' find out which button
    If Index = 1 Then ' okay button
        ' dialogue index
        Select Case dialogueIndex
        
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
        End Select
    ElseIf Index = 3 Then ' no button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendDeclineTradeRequest
            Case DIALOGUE_TYPE_PARTY
                SendDeclineParty
        End Select
    End If
End Sub

Public Function GetColorString(color As Long)
    Select Case color
        Case 0
            GetColorString = "Preto"
        Case 1
            GetColorString = "Azul"
        Case 2
            GetColorString = "Verde"
        Case 3
            GetColorString = "Ciano"
        Case 4
            GetColorString = "Vermelho"
        Case 5
            GetColorString = "Magenta"
        Case 6
            GetColorString = "Marrom"
        Case 7
            GetColorString = "Cinza"
        Case 8
            GetColorString = "Cinza escuro"
        Case 9
            GetColorString = "Azul claro"
        Case 10
            GetColorString = "Verde claro"
        Case 11
            GetColorString = "Ciano claro"
        Case 12
            GetColorString = "Vermelho claro"
        Case 13
            GetColorString = "Rosa"
        Case 14
            GetColorString = "Amarelo"
        Case 15
            GetColorString = "Branco"

    End Select
End Function

Public Sub MenuLoop()
Dim FrameTime As Long
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
restartmenuloop:
    ' *** Start GameLoop ***
    Do While Not InGame
        Tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call DrawGDI
        DoEvents
        
        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < Tick + Options.FPS
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

    ' Error handler
    Exit Sub
errorhandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        GoTo restartmenuloop
    ElseIf Options.Debug = 1 Then
        HandleError "MenuLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
    End If
End Sub

Sub ProcessWeather()
Dim i As Long
    If CurrentWeather > 0 Then
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    
    Dim Z As Long
    Z = 0
    For i = 1 To MAX_SPELLS

        If Not ShowSpell(i) Then skipThis = True

        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_SPELLS).Y + 160 + SpellTop + ((SpellOffsetY + 32) * ((Z) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((Z) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsPlayerEvoluteSpell = i
                    Exit Function
                End If
            End If
            Z = Z + 1
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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

Public Function IsPlayerQuest(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim i As Long
Dim ActualQuest As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsPlayerQuest = 0

    For i = 1 To MAX_QUESTS

        If Not Player(MyIndex).QuestState(i).State = 1 Then
            skipThis = True
        End If

        If Not skipThis Then
            ActualQuest = ActualQuest + 1
            With tempRec
                .Top = GUIWindow(GUI_SPELLS).Y + SpellTop + ((SpellOffsetY + 32) * ((ActualQuest - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((ActualQuest - 1) Mod SpellColumns)))
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

Public Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean, Optional ByVal emptySlot As Boolean = False) As Long
    Dim tempRec As RECT, skipThis As Boolean
    Dim i As Long
    Dim IsTradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsTradeItem = 0

    For i = 1 To MAX_INV
    
        If Yours Then
            IsTradeNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)
        Else
            IsTradeNum = TradeTheirOffer(i).Num
        End If
        
        If Not emptySlot Then
            If IsTradeNum <= 0 Or IsTradeNum > MAX_ITEMS Then skipThis = True
        End If
        
        If Not skipThis Then
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug = 1 Then
        Buttons(30).State = 2
        Buttons(31).State = 0
    Else
        Buttons(30).State = 0
        Buttons(31).State = 2
    End If
End Sub

Public Sub ScrollChatBox(ByVal direction As Byte)
    ' do a quick exit if we don't have enough text to scroll
    If totalChatLines < 8 Then
        ChatScroll = 8
        UpdateChatArray
        Exit Sub
    End If
    ' actually scroll
    If direction = 0 Then ' up
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
            GetEventTypeName = "@Mostrar mensagem: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_Menu
            GetEventTypeName = "@Mostrar escolhas"
            Exit Function
        Case Evt_Quit
            GetEventTypeName = "@Sair do evento"
            Exit Function
        Case Evt_OpenShop
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                GetEventTypeName = "@Abrir Shop: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "-" & Trim$(Shop(Events(EventIndex).SubEvents(SubIndex).Data(1)).Name)
            Else
                GetEventTypeName = "@Abrir Shop: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "- None "
            End If
            Exit Function
        Case Evt_OpenBank
            GetEventTypeName = "@Abrir Banco"
            Exit Function
        Case Evt_GiveItem
            GetEventTypeName = "@Mudar Item"
            Exit Function
        Case Evt_ChangeLevel
            GetEventTypeName = "@Mudar Level"
            Exit Function
        Case Evt_PlayAnimation
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                GetEventTypeName = "@Mostrar Animação: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "." & Trim$(Animation(Events(EventIndex).SubEvents(SubIndex).Data(1)).Name) & " {" & Events(EventIndex).SubEvents(SubIndex).Data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(3) & "}"
            Else
                GetEventTypeName = "@Mostrar Animação: Nenhuma {" & Events(EventIndex).SubEvents(SubIndex).Data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(3) & "}"
            End If
            Exit Function
        Case Evt_Warp
            GetEventTypeName = "@Teleportar: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & " {" & Events(EventIndex).SubEvents(SubIndex).Data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(3) & "}"
            Exit Function
        Case Evt_GOTO
            GetEventTypeName = "@Ir para: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_Switch
            On Error Resume Next
            If Events(EventIndex).SubEvents(SubIndex).Data(2) = 1 Then
                GetEventTypeName = "@Mudar Switch: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "." & Switches(Events(EventIndex).SubEvents(SubIndex).Data(1)) & " = True"
            Else
                GetEventTypeName = "@Mudar Switch: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "." & Switches(Events(EventIndex).SubEvents(SubIndex).Data(1)) & " = False"
            End If
            Exit Function
        Case Evt_Variable
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                GetEventTypeName = "@Mudar Variable: " & Variables(Events(EventIndex).SubEvents(SubIndex).Data(1))
            Else
                GetEventTypeName = "@Mudar Variable: Nenhuma"
            End If
            Exit Function
        Case Evt_AddText
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(2)
                Case 0: GetEventTypeName = "@Adicionar texto: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).Data(1)) & ", Player}"
                Case 1: GetEventTypeName = "@Adicionar texto: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).Data(1)) & ", Map}"
                Case 2: GetEventTypeName = "@Adicionar texto: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).Data(1)) & ", Global}"
            End Select
            Exit Function
        Case Evt_Chatbubble
            GetEventTypeName = "@Mostrar balão de fala"
            Exit Function
        Case Evt_Branch
            GetEventTypeName = "@Condição"
            Exit Function
        Case Evt_ChangeSkill
            GetEventTypeName = "@Mudar magias"
            Exit Function
        Case Evt_ChangeSprite
            GetEventTypeName = "@Mudar Sprite: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_ChangePK
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(1)
                Case 0: GetEventTypeName = "@Mudar PK: NO"
                Case 1: GetEventTypeName = "@Mudar PK: YES"
            End Select
            Exit Function
        Case Evt_SpawnNPC
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                If Map.Npc(Events(EventIndex).SubEvents(SubIndex).Data(1)) <> 0 Then GetEventTypeName = "@Nascer NPC: " & Trim$(Npc(Map.Npc(Events(EventIndex).SubEvents(SubIndex).Data(1))).Name)
            Else
                GetEventTypeName = "@Nascer NPC: Nenhum"
            End If
            Exit Function
        Case Evt_ChangeClass
            On Error Resume Next
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                GetEventTypeName = "@Mudar Classe: " & Trim$(Class(Events(EventIndex).SubEvents(SubIndex).Data(1)).Name)
            Else
                GetEventTypeName = "@Mudar Classe: Nenhuma"
            End If
            Exit Function
        Case Evt_ChangeSex
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(1)
                Case 0: GetEventTypeName = "@Mudar sexo: homem"
                Case 1: GetEventTypeName = "@Mudar sexo: mulher"
            End Select
            Exit Function
        Case Evt_ChangeExp
            GetEventTypeName = "@Mudar Exp"
            Exit Function
        Case Evt_SpecialEffect
            GetEventTypeName = "@Efeito especial"
            Exit Function
        Case Evt_PlaySound
            GetEventTypeName = "@Tocar som: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_PlayBGM
            GetEventTypeName = "@Tocar BGM: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_StopSound
            GetEventTypeName = "@Parar som"
            Exit Function
        Case Evt_FadeoutBGM
            GetEventTypeName = "@Parar BGM"
            Exit Function
        Case Evt_SetAccess
            GetEventTypeName = "@Setar acesso: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_CustomScript
            GetEventTypeName = "@Script especial: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_OpenEvent
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(3)
                Case 0: GetEventTypeName = "@Abrir evento: {" & Events(EventIndex).SubEvents(SubIndex).Data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(2) & "}"
                Case 1: GetEventTypeName = "@Fechar evento: {" & Events(EventIndex).SubEvents(SubIndex).Data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(2) & "}"
            End Select
            Exit Function
        Case Evt_Quest
            If Events(EventIndex).SubEvents(SubIndex).Data(2) = 0 Then GetEventTypeName = "@Alterar Quest: {Não completada}"
            If Events(EventIndex).SubEvents(SubIndex).Data(2) = 1 Then GetEventTypeName = "@Alterar Quest: {Em andamento}"
            If Events(EventIndex).SubEvents(SubIndex).Data(2) = 2 Then GetEventTypeName = "@Alterar Quest: {Completa}"
            Exit Function
    End Select
    GetEventTypeName = "Desconhecido"
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
                ReDim Preserve .Data(1 To 1)
            Case Evt_Menu
                If Not .HasText Then ReDim .Text(1 To 2)
                If UBound(.Text) < 2 Then ReDim Preserve .Text(1 To 2)
                If Not .HasData Then ReDim .Data(1 To 1)
                .HasText = True
                .HasData = True
            Case Evt_OpenShop
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_GOTO
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_GiveItem
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 3)
            Case Evt_PlayAnimation
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 3)
            Case Evt_Warp
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 3)
            Case Evt_Switch
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_Variable
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 4)
            Case Evt_AddText
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .Data(1 To 2)
            Case Evt_Chatbubble
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .Data(1 To 2)
            Case Evt_Branch
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 6)
            Case Evt_ChangeSkill
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_ChangeLevel
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_ChangeSprite
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangePK
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_SpawnNPC
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangeClass
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangeSex
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangeExp
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_SpecialEffect
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 5)
            Case Evt_PlaySound
                .HasText = True
                .HasData = False
                Erase .Data
                ReDim Preserve .Text(1 To 1)
            Case Evt_PlayBGM
                .HasText = True
                .HasData = False
                Erase .Data
                ReDim Preserve .Text(1 To 1)
            Case Evt_SetAccess
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_CustomScript
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_OpenEvent
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 4)
            Case Evt_Quest
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Else
                .HasText = False
                .HasData = False
                Erase .Text
                Erase .Data
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Effect(Index)), LenB(Effect(Index)))
    Effect(Index).Name = vbNullString
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
        If MapNpc(i).Num > 0 Then
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
            X = (GetPlayerX(i) * 32) + TempPlayer(i).xOffset + 32
            Y = (GetPlayerY(i) * 32) + TempPlayer(i).YOffset + 32
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
        If MapNpc(i).Num > 0 Then
            X = (MapNpc(i).X * 32) + TempMapNpc(i).xOffset + 32
            Y = (MapNpc(i).Y * 32) + TempMapNpc(i).YOffset + 32
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

Sub FazerBuraco(X As Long, Y As Long, Size As Long)
Dim i As Long
    
    For i = 1 To 10
        If Buracos(i).InUse = True Then
            If Buracos(i).X = X And Buracos(i).Y = Y Then
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
            Exit Sub
        End If
    Next i
End Sub

Function GetPlayerStatNextLevel(Index, ByVal Stat As Stats)
    GetPlayerStatNextLevel = StatNextLevel(Stat)
End Function

Function GetPlayerStatPrevLevel(Index, ByVal Stat As Stats)
    GetPlayerStatPrevLevel = StatLastLevel(Stat)
End Function
