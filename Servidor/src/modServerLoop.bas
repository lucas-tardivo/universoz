Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, X As Long, Y As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long, tmr60000 As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long
    Dim EsferasTick As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = GetTickCount
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick

        If Tick > tmr25 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If Player(i).VIP = 0 And GetPlayerAccess(i) = 0 And Player(i).GravityHours = 0 Then
                        If TempPlayer(i).LastMove + 1800000 < GetTickCount Then
                            AlertMSG i, printf("Você ficou inativo por 30 minutos e foi kickado!")
                        End If
                        If TempPlayer(i).LastMove + 1500000 < GetTickCount Then
                            If TempPlayer(i).AlertMSG = 0 Then
                                PlayerMsg i, printf("Voce está inativo a 25 minutos, mais 5 minutos e você será kickado"), Yellow
                                TempPlayer(i).AlertMSG = 1
                            End If
                        Else
                            TempPlayer(i).AlertMSG = 0
                        End If
                    End If
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If GetTickCount > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.Target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.Target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i, TargetType.TARGET_TYPE_PLAYER, 0, GetPlayerMap(i)
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < GetTickCount Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For X = 1 To MAX_DOTS
                        HandleDoT_Player i, X
                        HandleHoT_Player i, X
                    Next
                End If
            Next
            
            For i = 1 To MAX_AOEEFFECTS
                If AoEEffect(i).Tick + AoEEffect(i).Duration > GetTickCount Then
                    If (GetTickCount - AoEEffect(i).CastTick) > Spell(AoEEffect(i).SpellNum).AoETick And AoEEffect(i).Map > 0 Then
                        Dim AntiTrava As Long
                        AntiTrava = 0
SelectPlace:
                        X = rand(AoEEffect(i).X - Spell(AoEEffect(i).SpellNum).AoE, AoEEffect(i).X + Spell(AoEEffect(i).SpellNum).AoE)
                        Y = rand(AoEEffect(i).Y - Spell(AoEEffect(i).SpellNum).AoE, AoEEffect(i).Y + Spell(AoEEffect(i).SpellNum).AoE)
                        
                        If X < 0 Or X > Map(AoEEffect(i).Map).MaxX Or Y < 0 Or Y > Map(AoEEffect(i).Map).MaxY Then
                            AntiTrava = AntiTrava + 1
                            If AntiTrava <= 3 Then GoTo SelectPlace
                        Else
                            SendAnimation AoEEffect(i).Map, Spell(AoEEffect(i).SpellNum).SpellAnim, X, Y, rand(0, 3)
                            If AoEEffect(i).CasterType = TARGET_TYPE_PLAYER Then
                                Dim PlayerIndex As Long
                                PlayerIndex = FindPlayer(AoEEffect(i).Caster)
                                If PlayerIndex > 0 Then
                                    CheckAttack PlayerIndex, X, Y, GetPlayerSkillDamage(PlayerIndex, Spell(AoEEffect(i).SpellNum).Vital), AoEEffect(i).SpellNum
                                End If
                            Else
                                If MapNpc(AoEEffect(i).Map).Npc(Val(AoEEffect(i).Caster)).Num > 0 Then CheckAttack Val(AoEEffect(i).Caster), X, Y, Npc(MapNpc(AoEEffect(i).Map).Npc(Val(AoEEffect(i).Caster)).Num).Damage + Spell(AoEEffect(i).SpellNum).Vital, AoEEffect(i).SpellNum, True, AoEEffect(i).Map
                            End If
                        End If
                        
                        AoEEffect(i).CastTick = GetTickCount
                    End If
                End If
            Next i
            
            frmServer.lblCPS.Caption = Format$(GameCPS, "#,###,###,###")
            tmr25 = GetTickCount + 25
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            UpdateMapLogic
            tmr500 = GetTickCount + 500
        End If

        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            HandleArena
            ' Update the form labels, and reset the packets per second
            frmServer.lblPackIn.Caption = Trim(STR(PacketsIn))
            frmServer.lblPackOut.Caption = Trim(STR(PacketsOut))
            PacketsIn = 0
            PacketsOut = 0
            ' Update the Server Online Time
            ServerSeconds = ServerSeconds + 1
            If ServerSeconds > 59 Then
                ServerMinutes = ServerMinutes + 1
                ServerSeconds = 0
                If ServerMinutes > 59 Then
                    ServerMinutes = 0
                    ServerHours = ServerHours + 1
                End If
            End If
            frmServer.lblTime.Caption = Trim(KeepTwoDigit(STR(ServerHours))) & ":" & Trim(KeepTwoDigit(STR(ServerMinutes))) & ":" & Trim(KeepTwoDigit(STR(ServerSeconds)))
            
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = MapRespawn Then
                        If TempPlayer(i).RespawnTick + RespawnTime < GetTickCount Then
                            If TempPlayer(i).inDevSuite = 0 Then Call PlayerWarp(i, START_MAP, START_X, START_Y)
                        End If
                    End If
                    If TempPlayer(i).Trans > 0 Then
                        For X = 1 To Vitals.Vital_Count - 1
                            If TempPlayer(i).Trans > 0 Then
                                If Spell(TempPlayer(i).Trans).TransVital(X) > 0 Then
                                    Dim MPCost As Long
                                    MPCost = Spell(TempPlayer(i).Trans).TransVital(X)
                                    If UZ Then
                                        If TempPlayer(i).HairChange = 5 Then
                                            Dim PlanetNum As Long
                                            PlanetNum = GetPlanetNum(GetPlayerMap(i))
                                            If PlanetNum > 0 And PlanetNum <= MAX_PLANET_BASE Then
                                                If Planets(PlanetNum).MoonData.Pic > 0 Then MPCost = 0
                                            End If
                                        End If
                                    End If
                                    If GetPlayerVital(i, X) - MPCost > 0 Then
                                        SetPlayerVital i, X, GetPlayerVital(i, X) - MPCost
                                    Else
                                        SetPlayerVital i, X, 0
                                        TransPlayer i, 0
                                        If X = Vitals.HP Then
                                            PlayerMsg i, printf("Você não conseguiu manter a habilidade por mais tempo e morreu!"), brightred
                                            If Player(i).IsDead = 0 And TempPlayer(i).inDevSuite = 0 Then Player(i).IsDead = 1
                                            Call SendPlayerData(i)
                                            'OnDeath i
                                        Else
                                            PlayerMsg i, printf("Você não tem mais energia para manter essa habilidade ativa!"), brightred
                                        End If
                                    End If
                                    SendVital i, X
                                End If
                            End If
                        Next X
                    End If
                    'If Map(GetPlayerMap(i)).Tile(GetPlayerX(i), GetPlayerY(i)).Type = TILE_TYPE_RESOURCE Then
                    '    If Resource(Map(GetPlayerMap(i)).Tile(GetPlayerX(i), GetPlayerY(i)).Data1).ResourceType = 3 Then
                    '        If TempPlayer(i).NextFish < GetTickCount Then
                    '            Call SendFishingTime(i)
                    '        End If
                    '    End If
                    'End If
                End If
            Next i
            
            If UZ Then
                For i = 1 To MAX_PLANET_BASE
                    If Planets(i).TimeToExplode > 0 Then
                        If Planets(i).TimeToExplode < GetTickCount Then
                            SendBossMsg Planets(i).Map, "O planeta " & Trim$(Planets(i).Name) & " acabou de explodir!", Yellow
                            ExplodePlanet i
                        End If
                    End If
                Next i
                Dim n As Long
                For i = 1 To UBound(MatchData)
                    If MatchData(i).Active >= 1 And Not frmServer.chkTrava.Value = 1 Then
                        If MatchData(i).Points < Planets(MatchData(i).Planet).PointsToConquest Or Planets(MatchData(i).Planet).Type >= 2 Then
                            If MatchData(i).WaveTick > GetTickCount Then
                                If MatchData(i).SpawnTick < GetTickCount And PlayerAlive(i) Then
                                    SpawnEnemy i
                                    MatchData(i).SpawnTick = GetTickCount + 1000
                                End If
                            Else
                                If MatchData(i).WaveTick + Planets(MatchData(i).Planet).WaveCooldown < GetTickCount Then
                                    MatchData(i).WaveNum = MatchData(i).WaveNum + 1
                                    RollWave i
                                    If Not EliteWave(i) Then
                                        If MatchData(i).WaveTick < 0 Then 'Manter onda anterior por efeito especial
                                            MatchData(i).WaveTick = MatchData(i).WaveTick * -1 'Invertemos
                                        Else
                                            MatchData(i).WaveTick = 0
                                        End If
                                        If Int(MatchData(i).WaveNum / 3) < MAX_MAP_NPCS - (Planets(MatchData(i).Planet).WaveDuration / 1000) Then
                                            MatchData(i).WaveTick = GetTickCount + MatchData(i).WaveTick + Planets(MatchData(i).Planet).WaveDuration + (Int(MatchData(i).WaveNum / 4) * 1000)
                                        Else
                                            MatchData(i).WaveTick = GetTickCount + MatchData(i).WaveTick + (MAX_MAP_NPCS * 1000)
                                        End If
                                        If NpcCount(Planets(MatchData(i).Planet).Map) = 0 And Planets(MatchData(i).Planet).Type = 0 Then
                                            MatchData(i).Stars = MatchData(i).Stars + 1
                                            SendMatchData i
                                        End If
                                        ClearWaveItems Planets(MatchData(i).Planet).Map
                                        If Planets(MatchData(i).Planet).Type = 0 Then SendBossMsg Planets(MatchData(i).Planet).Map, "Preparem-se para a onda número " & MatchData(i).WaveNum, Yellow
                                    Else
                                        MatchData(i).WaveTick = GetTickCount + 1900
                                        If Planets(MatchData(i).Planet).Type = 0 Then SendBossMsg Planets(MatchData(i).Planet).Map, "Tropas especiais foram acionadas! Prepare-se!", Yellow
                                    End If
                                End If
                            End If
                        Else
                            If MatchData(i).Active = 1 Then
                                If MatchData(i).Indexes(1) > 0 Then
                                Planets(MatchData(i).Planet).Owner = Trim$(GetPlayerName(MatchData(i).Indexes(1)))
                                Planets(MatchData(i).Planet).State = 2
                                Planets(MatchData(i).Planet).TimeToExplode = GetTickCount + 900000
                                Planets(MatchData(i).Planet).Preco = (Planets(MatchData(i).Planet).Preco / 100) * (100 + MatchData(i).PriceBonus)
                                SendBossMsg Planets(MatchData(i).Planet).Map, "Parabéns! Você conquistou este planeta!", Yellow
                                MatchData(i).WaveTick = GetTickCount + 5000
                                MatchData(i).Winner = MatchData(i).Indexes(1)
                                MatchData(i).Active = 2
                                SendPlanetToAll MatchData(i).Planet
                                TempPlayer(MatchData(i).Winner).PlanetCaptured = MatchData(i).Planet
                                
                                For n = 1 To MAX_MAP_NPCS
                                    DespawnNPC Planets(MatchData(i).Planet).Map, n
                                Next n
                                For n = 1 To UBound(MatchData(i).Indexes)
                                    If IsPlaying(MatchData(i).Indexes(n)) Then
                                        GiveConquistaReward MatchData(i).Indexes(n), UBound(MatchData(i).Indexes), MatchData(i).HighLevel, 100
                                    End If
                                Next n
                                End If
                            Else
                                If MatchData(i).WaveTick < GetTickCount Then
                                    MatchData(i).Active = 0
                                    SendMatchData i
                                    For n = 1 To UBound(MatchData(i).Indexes)
                                        If IsPlaying(MatchData(i).Indexes(n)) Then
                                            TempPlayer(MatchData(i).Indexes(n)).MatchIndex = 0
                                            Viajar MatchData(i).Indexes(n), ViagemMap, Planets(MatchData(i).Planet).X, Planets(MatchData(i).Planet).Y
                                        End If
                                    Next n
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
            
            Call UpdateTransportes
            Call UpdateProvacoes
            
            tmr1000 = GetTickCount + 1000
        End If
        
        If UZ Then
            If EventGlobalTick + EventGlobalInterval < GetTickCount Then
                Call StartGlobalEvent(rand(1, frmServer.cmbEvent.ListCount))
                EventGlobalTick = GetTickCount
            End If
            If EventGlobalTick + 300000 < GetTickCount And EventGlobalType > 0 Then
                Call DesactivateEvent
            End If
        End If

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If

        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            SetupBonuses
            
            'Anunciar preços
            If UZ Then
                Dim Base(1 To 3) As Long
                Base(1) = 800
                Base(2) = 1600
                Base(3) = 3400
                For i = 1 To 3
                    If GetEspeciariaPrice(i) >= Base(i) * 1.5 Then
                        GlobalMsg "[ANUNCIO] A " & Trim$(Item(79 + i).Name) & " está super valorizada com uma alta de " & Int(((GetEspeciariaPrice(i) / Base(i) * 100) - 100)) & "%! Deem prioridade para esta especiaria!", Yellow
                    End If
                Next i
            End If
            
            LastUpdateMapSpawnItems = GetTickCount + 300000
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 300000
        End If
        
        If Tick > tmr60000 Then
            'update esotericas time
            UpdatePlayersEsotericas
            
            If DatePart("w", Now) >= 1 And DatePart("w", Now) <= 5 Then
                If DatePart("h", Now) >= 20 And DatePart("h", Now) <= 22 Then
                    If Not TesouroStarted Then
                        If TotalOnlinePlayers >= 5 Then PlanetaDoTesouro
                    End If
                Else
                    If TesouroStarted Then
                        DestroyTesouro
                    End If
                End If
            Else
                If DatePart("h", Now) >= 22 Or DatePart("h", Now) <= 2 Then
                    If Not TesouroStarted Then
                        If TotalOnlinePlayers >= 5 Then PlanetaDoTesouro
                    End If
                Else
                    If TesouroStarted Then
                        DestroyTesouro
                    End If
                End If
            End If
            
            For i = 1 To Player_HighIndex
                If Player(i).GravityHours > 0 Then
                    Call HandleGravity(i)
                End If
            Next i
            
            If ShenlongTick <> 0 Then
                If ShenlongTick + 600000 < GetTickCount Then
                    If ShenlongActive = 1 Then
                        ShenlongActive = 0
                        ShenlongMap = 0
                        SendShenlong 0, 0, 0
                    End If
                End If
            End If
            
            If LastCustomPlanetsRespawn + CUSTOM_PLANETS_RESPAWN < GetTickCount Then
                CreateCustomPlanets
                LastCustomPlanetsRespawn = GetTickCount
            End If
            
            Call UpdatePlanets
            
            'If MapNpc(30).Npc(7).Num > 0 Then
            '    If MapNpc(30).Npc(7).Vital(Vitals.HP) = Npc(MapNpc(30).Npc(7).Num).HP Then
            '        MapNpc(30).Npc(7).Num = 0
            '        UpdateMapBlock 30, MapNpc(30).Npc(7).x, MapNpc(30).Npc(7).y, False
            '        ' send death to the map
            '        Dim buffer As clsBuffer
            '        Set buffer = New clsBuffer
            '        buffer.WriteLong SNpcDead
            '        buffer.WriteLong 7
            '        SendDataToMap 30, buffer.ToArray()
            '        Set buffer = Nothing
            '    End If
            'End If
            
            tmr60000 = GetTickCount + 60000
        End If
        
        If EsferasTick < GetTickCount Then
            SendDragonballs
            EsferasTick = GetTickCount + 10800000
        End If

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Private Sub UpdateMapSpawnItems()
    Dim X As Long
    Dim Y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For Y = 1 To MAX_MAPS

        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(Y) Then

            ' Clear out unnecessary junk
            For X = 1 To MAX_MAP_ITEMS
                If MapItem(Y, X).Num > 0 Then
                    If Not Item(MapItem(Y, X).Num).Type = ITEM_TYPE_DRAGONBALL Then
                        Call ClearMapItem(X, Y)
                    End If
                End If
            Next

            ' Spawn the items
            Call SpawnMapItems(Y)
            Call SendMapItemsToAll(Y)
        End If

        DoEvents
    Next

End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, X As Long, MapNum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long
    Dim Target As Long, TargetType As Byte, didwalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim targetX As Long, targetY As Long, target_verify As Boolean

    For MapNum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(MapNum, i).Num > 0 Then
                If MapItem(MapNum, i).PlayerName <> vbNullString Then
                    ' make item public?
                    If MapItem(MapNum, i).playerTimer < GetTickCount Then
                        ' make it public
                        MapItem(MapNum, i).PlayerName = vbNullString
                        MapItem(MapNum, i).playerTimer = 0
                        ' send updates to everyone
                        SendMapItemsToAll MapNum
                    End If
                    ' despawn item?
                    If MapItem(MapNum, i).Num > 0 Then
                    If MapItem(MapNum, i).canDespawn And Not Item(MapItem(MapNum, i).Num).Type = ITEM_TYPE_DRAGONBALL Then
                        If MapItem(MapNum, i).despawnTimer < GetTickCount Then
                            ' despawn it
                            ClearMapItem i, MapNum
                            ' send updates to everyone
                            SendMapItemsToAll MapNum
                        End If
                    End If
                    End If
                End If
            End If
        Next
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(MapNum).Npc(i).Num > 0 Then
                For X = 1 To MAX_DOTS
                    HandleDoT_Npc MapNum, i, X
                    HandleHoT_Npc MapNum, i, X
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(MapNum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(MapNum).Resource_Count
                Resource_index = Map(MapNum).Tile(ResourceCache(MapNum).ResourceData(i).X, ResourceCache(MapNum).ResourceData(i).Y).data1

                If Resource_index > 0 Then
                    If ResourceCache(MapNum).ResourceData(i).ResourceState = 1 And (ResourceCache(MapNum).ResourceData(i).cur_health < 1 Or Resource(Resource_index).ResourceType = 4) Then  ' dead or fucked up
                        Dim ResourceTimer As Long
                        ResourceTimer = ResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000)
                        If Resource(Resource_index).ResourceType = 4 Then
                            If ResourceCache(MapNum).ResourceData(i).cur_health > 1 Then
                                ResourceTimer = ResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * (10 * (100 - ResourceCache(MapNum).ResourceData(i).cur_health)))
                            End If
                        End If
                        If ResourceTimer < GetTickCount Then
                            ResourceCache(MapNum).ResourceData(i).ResourceTimer = GetTickCount
                            ResourceCache(MapNum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            
                            SendResourceCacheToMap MapNum, i
                            If Resource(Resource_index).ResourceType = 4 Then
                                Dim PlanetNum As Long
                                PlanetNum = GetPlanetNum(MapNum)
                                Dim PlayerIndex As Long
                                PlayerIndex = FindPlayer(Trim$(Planets(PlanetNum).Owner))
                                If PlayerIndex > 0 Then
                                    PlayerMsg PlayerIndex, "Um extrator no planeta " & Trim$(Planets(PlanetNum).Name) & " está completo!", Yellow
                                    SendPlaySound PlayerIndex, "Success1.mp3"
                                End If
                            Else
                                ResourceCache(MapNum).ResourceData(i).cur_health = Resource(Resource_index).health
                            End If
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(MapNum) = YES Then
            TickCount = GetTickCount
            
            For X = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(MapNum).Npc(X).Num
                
                If Map(MapNum).Npc(X) > 0 And MapNpc(MapNum).Npc(X).Num > 0 Then
                    If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_TREINO Or Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_TREINOHOUSE Then
                        GoTo DontWalk
                    End If
                End If

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(X) > 0 And MapNpc(MapNum).Npc(X).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(MapNum).TempNpc(X).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum And MapNpc(MapNum).Npc(X).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        n = Npc(NpcNum).Range
                                        DistanceX = MapNpc(MapNum).Npc(X).X - GetPlayerX(i)
                                        DistanceY = MapNpc(MapNum).Npc(X).Y - GetPlayerY(i)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                'If Len(Trim$(Npc(npcnum).AttackSay)) > 0 Then
                                                '    Call SendChatBubble(MapNum, X, TARGET_TYPE_NPC, Trim$(Npc(npcnum).AttackSay), DarkBrown)
                                                'End If
                                                MapNpc(MapNum).Npc(X).TargetType = 1 ' player
                                                MapNpc(MapNum).Npc(X).Target = i
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(X) > 0 And MapNpc(MapNum).Npc(X).Num > 0 Then
                    If MapNpc(MapNum).TempNpc(X).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(MapNum).TempNpc(X).StunTimer + (MapNpc(MapNum).TempNpc(X).StunDuration * 1000) Then
                            MapNpc(MapNum).TempNpc(X).StunDuration = 0
                            MapNpc(MapNum).TempNpc(X).StunTimer = 0
                        End If
                    Else
                            
                        Target = MapNpc(MapNum).Npc(X).Target
                        TargetType = MapNpc(MapNum).Npc(X).TargetType
                        If Npc(MapNpc(MapNum).Npc(X).Num).Speed > 0 Then
                            If MapNpc(MapNum).Npc(X).WalkingTick + (1000 / Npc(MapNpc(MapNum).Npc(X).Num).Speed) > GetTickCount Then Exit Sub
                        End If
                        
                        ' Check to see if its time for the npc to walk
                        If Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER And Npc(NpcNum).Speed > 0 Then
                        
                            If TargetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If Target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                        didwalk = False
                                        target_verify = True
                                        targetY = GetPlayerY(Target)
                                        targetX = GetPlayerX(Target)
                                        
                                        If TempPlayer(Target).Target = 0 Then 'Mirar de volta
                                            TempPlayer(Target).Target = X
                                            TempPlayer(Target).TargetType = TARGET_TYPE_NPC
                                            SendTarget Target
                                        End If
                                    Else
                                        MapNpc(MapNum).Npc(X).TargetType = 0 ' clear
                                        MapNpc(MapNum).Npc(X).Target = 0
                                    End If
                                End If
                            
                            ElseIf TargetType = 2 Then 'npc
                                
                                If Target > 0 Then
                                    
                                    If MapNpc(MapNum).Npc(Target).Num > 0 Then
                                        didwalk = False
                                        target_verify = True
                                        targetY = MapNpc(MapNum).Npc(Target).Y
                                        targetX = MapNpc(MapNum).Npc(Target).X
                                    Else
                                        MapNpc(MapNum).Npc(X).TargetType = 0 ' clear
                                        MapNpc(MapNum).Npc(X).Target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                'Gonna make the npcs smarter.. Implementing a pathfinding algorithm.. we shall see what happens.
                                If IsOneBlockAway(targetX, targetY, CLng(MapNpc(MapNum).Npc(X).X), CLng(MapNpc(MapNum).Npc(X).Y)) = False Then
                                    If PathfindingType = 1 Then
                                        i = Int(Rnd * 5)
            
                                        ' Lets move the npc
                                        Select Case i
                                            Case 0
                                                 ' Up Left

                                                    If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then
            
                                                        If MapNpc(MapNum).Npc(X).X > targetX Then
            
                                                            If CanNpcMove(MapNum, X, DIR_UP_LEFT) Then
            
                                                                Call NpcMove(MapNum, X, DIR_UP_LEFT, MOVING_WALKING)
            
                                                                didwalk = True
            
                                                            End If
            
                                                        End If
            
                                                    End If
            
                                                   
            
                                                    ' Up right
            
                                                    If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then
            
                                                        If MapNpc(MapNum).Npc(X).X < targetX Then
            
                                                            If CanNpcMove(MapNum, X, DIR_UP_RIGHT) Then
            
                                                                Call NpcMove(MapNum, X, DIR_UP_RIGHT, MOVING_WALKING)
            
                                                                didwalk = True
            
                                                            End If
            
                                                        End If
            
                                                    End If
            
                                                   
            
                                                    ' Down Left
            
                                                    If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then
            
                                                        If MapNpc(MapNum).Npc(X).X > targetX Then
            
                                                            If CanNpcMove(MapNum, X, DIR_DOWN_LEFT) Then
            
                                                                Call NpcMove(MapNum, X, DIR_DOWN_LEFT, MOVING_WALKING)
            
                                                                didwalk = True
            
                                                            End If
            
                                                        End If
            
                                                    End If
            
                                                   
            
                                                    ' Down Right
            
                                                    If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then
            
                                                        If MapNpc(MapNum).Npc(X).X < targetX Then
            
                                                            If CanNpcMove(MapNum, X, DIR_DOWN_RIGHT) Then
            
                                                                Call NpcMove(MapNum, X, DIR_DOWN_RIGHT, MOVING_WALKING)
            
                                                                didwalk = True
            
                                                            End If
            
                                                        End If
            
                                                    End If
            
                                                ' Up
                                                If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_UP) Then
                                                        Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If MapNpc(MapNum).Npc(X).X > targetX And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If MapNpc(MapNum).Npc(X).X < targetX And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 1
                                                ' Up Left

                                                    If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then
            
                                                        If MapNpc(MapNum).Npc(X).X > targetX Then
            
                                                            If CanNpcMove(MapNum, X, DIR_UP_LEFT) Then
            
                                                                Call NpcMove(MapNum, X, DIR_UP_LEFT, MOVING_WALKING)
            
                                                                didwalk = True
            
                                                            End If
            
                                                        End If
            
                                                    End If
            
                                                   
            
                                                    ' Up right
            
                                                    If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then
            
                                                        If MapNpc(MapNum).Npc(X).X < targetX Then
            
                                                            If CanNpcMove(MapNum, X, DIR_UP_RIGHT) Then
            
                                                                Call NpcMove(MapNum, X, DIR_UP_RIGHT, MOVING_WALKING)
            
                                                                didwalk = True
            
                                                            End If
            
                                                        End If
            
                                                    End If
            
                                                   
            
                                                    ' Down Left
            
                                                    If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then
            
                                                        If MapNpc(MapNum).Npc(X).X > targetX Then
            
                                                            If CanNpcMove(MapNum, X, DIR_DOWN_LEFT) Then
            
                                                                Call NpcMove(MapNum, X, DIR_DOWN_LEFT, MOVING_WALKING)
            
                                                                didwalk = True
            
                                                            End If
            
                                                        End If
            
                                                    End If
            
                                                   
            
                                                    ' Down Right
            
                                                    If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then
            
                                                        If MapNpc(MapNum).Npc(X).X < targetX Then
            
                                                            If CanNpcMove(MapNum, X, DIR_DOWN_RIGHT) Then
            
                                                                Call NpcMove(MapNum, X, DIR_DOWN_RIGHT, MOVING_WALKING)
            
                                                                didwalk = True
            
                                                            End If
            
                                                        End If
            
                                                    End If
            
                                                ' Right
                                                If MapNpc(MapNum).Npc(X).X < targetX And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If MapNpc(MapNum).Npc(X).X > targetX And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_UP) Then
                                                        Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 2
                                            
                                                        ' Up Left

                                        If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then

                                            If MapNpc(MapNum).Npc(X).X > targetX Then

                                                If CanNpcMove(MapNum, X, DIR_UP_LEFT) Then

                                                    Call NpcMove(MapNum, X, DIR_UP_LEFT, MOVING_WALKING)

                                                    didwalk = True

                                                End If

                                            End If

                                        End If

                                       

                                        ' Up right

                                        If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then

                                            If MapNpc(MapNum).Npc(X).X < targetX Then

                                                If CanNpcMove(MapNum, X, DIR_UP_RIGHT) Then

                                                    Call NpcMove(MapNum, X, DIR_UP_RIGHT, MOVING_WALKING)

                                                    didwalk = True

                                                End If

                                            End If

                                        End If

                                       

                                        ' Down Left

                                        If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then

                                            If MapNpc(MapNum).Npc(X).X > targetX Then

                                                If CanNpcMove(MapNum, X, DIR_DOWN_LEFT) Then

                                                    Call NpcMove(MapNum, X, DIR_DOWN_LEFT, MOVING_WALKING)

                                                    didwalk = True

                                                End If

                                            End If

                                        End If

                                       

                                        ' Down Right

                                        If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then

                                            If MapNpc(MapNum).Npc(X).X < targetX Then

                                                If CanNpcMove(MapNum, X, DIR_DOWN_RIGHT) Then

                                                    Call NpcMove(MapNum, X, DIR_DOWN_RIGHT, MOVING_WALKING)

                                                    didwalk = True

                                                End If

                                            End If

                                        End If
            
                                                ' Down
                                                If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_UP) Then
                                                        Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If MapNpc(MapNum).Npc(X).X < targetX And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If MapNpc(MapNum).Npc(X).X > targetX And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 3
                                            
                                                    ' Up Left

                                        If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then

                                            If MapNpc(MapNum).Npc(X).X > targetX Then

                                                If CanNpcMove(MapNum, X, DIR_UP_LEFT) Then

                                                    Call NpcMove(MapNum, X, DIR_UP_LEFT, MOVING_WALKING)

                                                    didwalk = True

                                                End If

                                            End If

                                        End If

                                       

                                        ' Up right

                                        If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then

                                            If MapNpc(MapNum).Npc(X).X < targetX Then

                                                If CanNpcMove(MapNum, X, DIR_UP_RIGHT) Then

                                                    Call NpcMove(MapNum, X, DIR_UP_RIGHT, MOVING_WALKING)

                                                    didwalk = True

                                                End If

                                            End If

                                        End If

                                       

                                        ' Down Left

                                        If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then

                                            If MapNpc(MapNum).Npc(X).X > targetX Then

                                                If CanNpcMove(MapNum, X, DIR_DOWN_LEFT) Then

                                                    Call NpcMove(MapNum, X, DIR_DOWN_LEFT, MOVING_WALKING)

                                                    didwalk = True

                                                End If

                                            End If

                                        End If

                                       

                                        ' Down Right

                                        If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then

                                            If MapNpc(MapNum).Npc(X).X < targetX Then

                                                If CanNpcMove(MapNum, X, DIR_DOWN_RIGHT) Then

                                                    Call NpcMove(MapNum, X, DIR_DOWN_RIGHT, MOVING_WALKING)

                                                    didwalk = True

                                                End If

                                            End If

                                        End If
            
                                                ' Left
                                                If MapNpc(MapNum).Npc(X).X > targetX And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_LEFT) Then
                                                        Call NpcMove(MapNum, X, DIR_LEFT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If MapNpc(MapNum).Npc(X).X < targetX And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_RIGHT) Then
                                                        Call NpcMove(MapNum, X, DIR_RIGHT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If MapNpc(MapNum).Npc(X).Y > targetY And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_UP) Then
                                                        Call NpcMove(MapNum, X, DIR_UP, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If MapNpc(MapNum).Npc(X).Y < targetY And Not didwalk Then
                                                    If CanNpcMove(MapNum, X, DIR_DOWN) Then
                                                        Call NpcMove(MapNum, X, DIR_DOWN, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                        End Select
            
                                        ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                        If Not didwalk Then
                                            If MapNpc(MapNum).Npc(X).X - 1 = targetX And MapNpc(MapNum).Npc(X).Y = targetY Then
                                                If MapNpc(MapNum).Npc(X).Dir <> DIR_LEFT Then
                                                    Call NpcDir(MapNum, X, DIR_LEFT)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If MapNpc(MapNum).Npc(X).X + 1 = targetX And MapNpc(MapNum).Npc(X).Y = targetY Then
                                                If MapNpc(MapNum).Npc(X).Dir <> DIR_RIGHT Then
                                                    Call NpcDir(MapNum, X, DIR_RIGHT)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If MapNpc(MapNum).Npc(X).X = targetX And MapNpc(MapNum).Npc(X).Y - 1 = targetY Then
                                                If MapNpc(MapNum).Npc(X).Dir <> DIR_UP Then
                                                    Call NpcDir(MapNum, X, DIR_UP)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If MapNpc(MapNum).Npc(X).X = targetX And MapNpc(MapNum).Npc(X).Y + 1 = targetY Then
                                                If MapNpc(MapNum).Npc(X).Dir <> DIR_DOWN Then
                                                    Call NpcDir(MapNum, X, DIR_DOWN)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            ' We could not move so Target must be behind something, walk randomly.
                                            If Not didwalk Then
                                                i = Int(Rnd * 2)
            
                                                If i = 1 Then
                                                    i = Int(Rnd * 4)
            
                                                    If CanNpcMove(MapNum, X, i) Then
                                                        Call NpcMove(MapNum, X, i, MOVING_WALKING)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        i = FindNpcPath(MapNum, X, targetX, targetY)
                                        If i < 4 Then 'Returned an answer. Move the NPC
                                            If CanNpcMove(MapNum, X, i) Then
                                                NpcMove MapNum, X, i, MOVING_WALKING
                                            End If
                                        Else 'No good path found. Move randomly
                                            i = Int(Rnd * 4)
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
                
                                                If CanNpcMove(MapNum, X, i) Then
                                                    Call NpcMove(MapNum, X, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    Call NpcDir(MapNum, X, GetNpcDir(targetX, targetY, CLng(MapNpc(MapNum).Npc(X).X), CLng(MapNpc(MapNum).Npc(X).Y)))
                                End If
                            Else
                                i = Int(Rnd * 4)
    
                                If i = 1 Then
                                    i = Int(Rnd * 4)
    
                                    If CanNpcMove(MapNum, X, i) Then
                                        Call NpcMove(MapNum, X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                            MapNpc(MapNum).Npc(X).WalkingTick = GetTickCount
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(X) > 0 And MapNpc(MapNum).Npc(X).Num > 0 Then
                    Target = MapNpc(MapNum).Npc(X).Target
                    TargetType = MapNpc(MapNum).Npc(X).TargetType

                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                    
                        If TargetType = 1 Then ' player
                        
                            If UZ Then
                            Dim CantAttack As Boolean
                                PlanetNum = PlayerMapIndex(MapNum)
                                If PlanetNum > 0 Then
                                    If Trim$(LCase(PlayerPlanet(PlanetNum).PlanetData.Owner)) = Trim$(LCase(GetPlayerName(Target))) Then CantAttack = True
                                End If
                            End If
                            
                            If Not CantAttack Then
                                ' Is the target playing and on the same map?
                                If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                    TryNpcAttackPlayer X, Target
                                Else
                                    ' Player left map or game, set target to 0
                                    MapNpc(MapNum).Npc(X).Target = 0
                                    MapNpc(MapNum).Npc(X).TargetType = 0 ' clear
                                End If
                            End If
                        Else
                            ' lol no npc combat :(
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(MapNum).TempNpc(X).stopRegen Then
                    If MapNpc(MapNum).Npc(X).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(MapNum).Npc(X).Vital(Vitals.HP) > 0 Then
                            MapNpc(MapNum).Npc(X).Vital(Vitals.HP) = MapNpc(MapNum).Npc(X).Vital(Vitals.HP) + GetNpcVitalRegen(NpcNum, Vitals.HP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(MapNum).Npc(X).Vital(Vitals.HP) > GetNpcMaxVital(MapNum, X, Vitals.HP) Then
                                MapNpc(MapNum).Npc(X).Vital(Vitals.HP) = GetNpcMaxVital(MapNum, X, Vitals.HP)
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(MapNum).Npc(X).Num = 0 And Map(MapNum).Npc(X) > 0 Then
                    If TickCount > MapNpc(MapNum).TempNpc(X).SpawnWait + (Npc(Map(MapNum).Npc(X)).SpawnSecs * 1000) Then
                        If Not isProvação(MapNum) Then
                            Call SpawnNpc(X, MapNum)
                        End If
                    End If
                End If
                
DontWalk:
            Next
        Else
            If UZ And PlanetStarted Then
                If MapNum >= PlanetStart And MapNum < PlanetStart + MAX_PLANET_BASE Then
                    If Planets(GetPlanetNum(MapNum)).Map = MapNum Then
                        If Planets(GetPlanetNum(MapNum)).State = 1 Then Planets(GetPlanetNum(MapNum)).State = 0
                    End If
                End If
            End If
        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

End Sub



Private Sub UpdatePlayerVitals()
Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not TempPlayer(i).stopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, Vitals.HP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
    
                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, Vitals.MP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...", ChatSystem)

        For i = 1 To Player_HighIndex

            If IsPlaying(i) Then
                Call SavePlayer(i)
                Call SaveBank(i)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", brightblue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.", ChatSystem)
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", brightred)
        Call DestroyServer
    End If

End Sub

Function FindNpcPath(MapNum As Long, MapNPCNum As Long, targetX As Long, targetY As Long) As Long
Dim tim As Long, sX As Long, sY As Long, Pos() As Long, reachable As Boolean, X As Long, Y As Long, j As Long, LastSum As Long, Sum As Long, FX As Long, FY As Long, i As Long
Dim path() As Vector, LastX As Long, LastY As Long, did As Boolean

'Initialization phase
tim = 0
sX = MapNpc(MapNum).Npc(MapNPCNum).X
sY = MapNpc(MapNum).Npc(MapNPCNum).Y
FX = targetX
FY = targetY

ReDim Pos(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
Pos = MapBlocks(MapNum).Blocks

Pos(sX, sY) = 100 + tim
Pos(FX, FY) = 2

'reset reachable
reachable = False

'Do while reachable is false... if its set true in progress, we jump out
'If the path is decided unreachable in process, we will use exit sub. Not proper,
'but faster ;-)
Do While reachable = False
    'we loop through all squares
    For j = 0 To Map(MapNum).MaxY
        For i = 0 To Map(MapNum).MaxX
            'If j = 10 And i = 0 Then MsgBox "hi!"
            'If they are to be extended, the pointer TIM is on them
            If Pos(i, j) = 100 + tim Then
            'The part is to be extended, so do it
                'We have to make sure that there is a pos(i+1,j) BEFORE we actually use it,
                'because then we get error... If the square is on side, we dont test for this one!
                If i < Map(MapNum).MaxX Then
                    'If there isnt a wall, or any other... thing
                    If Pos(i + 1, j) = 0 Then
                        'Expand it, and make its pos equal to tim+1, so the next time we make this loop,
                        'It will exapand that square too! This is crucial part of the program
                        Pos(i + 1, j) = 100 + tim + 1
                    ElseIf Pos(i + 1, j) = 2 Then
                        'If the position is no 0 but its 2 (FINISH) then Reachable = true!!! We found end
                        reachable = True
                    End If
                End If
            
                'This is the same as the last one, as i said a lot of copy paste work and editing that
                'This is simply another side that we have to test for... so instead of i+1 we have i-1
                'Its actually pretty same then... I wont comment it therefore, because its only repeating
                'same thing with minor changes to check sides
                If i > 0 Then
                    If Pos((i - 1), j) = 0 Then
                        Pos(i - 1, j) = 100 + tim + 1
                    ElseIf Pos(i - 1, j) = 2 Then
                        reachable = True
                    End If
                End If
            
                If j < Map(MapNum).MaxY Then
                    If Pos(i, j + 1) = 0 Then
                        Pos(i, j + 1) = 100 + tim + 1
                    ElseIf Pos(i, j + 1) = 2 Then
                        reachable = True
                    End If
                End If
            
                If j > 0 Then
                    If Pos(i, j - 1) = 0 Then
                        Pos(i, j - 1) = 100 + tim + 1
                    ElseIf Pos(i, j - 1) = 2 Then
                        reachable = True
                    End If
                End If
            End If
            DoEvents
        Next i
    Next j
    
    'If the reachable is STILL false, then
    If reachable = False Then
        'reset sum
        Sum = 0
        For j = 0 To Map(MapNum).MaxY
            For i = 0 To Map(MapNum).MaxX
            'we add up ALL the squares
            Sum = Sum + Pos(i, j)
            Next i
        Next j
        
        'Now if the sum is euqal to the last sum, its not reachable, if it isnt, then we store
        'sum to lastsum
        If Sum = LastSum Then
            FindNpcPath = 4
            Exit Function
        Else
            LastSum = Sum
        End If
    End If
    
    'we increase the pointer to point to the next squares to be expanded
    tim = tim + 1
Loop

'We work backwards to find the way...
LastX = FX
LastY = FY

ReDim path(tim + 1)

'The following code may be a little bit confusing but ill try my best to explain it.
'We are working backwards to find ONE of the shortest ways back to Start.
'So we repeat the loop until the LastX and LastY arent in start. Look in the code to see
'how LastX and LasY change
Do While LastX <> sX Or LastY <> sY
    'We decrease tim by one, and then we are finding any adjacent square to the final one, that
    'has that value. So lets say the tim would be 5, because it takes 5 steps to get to the target.
    'Now everytime we decrease that, so we make it 4, and we look for any adjacent square that has
    'that value. When we find it, we just color it yellow as for the solution
    tim = tim - 1
    'reset did to false
    did = False
    
    'If we arent on edge
    If LastX < Map(MapNum).MaxX Then
        'check the square on the right of the solution. Is it a tim-1 one? or just a blank one
        If Pos(LastX + 1, LastY) = 100 + tim Then
            'if it, then make it yellow, and change did to true
            LastX = LastX + 1
            did = True
        End If
    End If
    
    'This will then only work if the previous part didnt execute, and did is still false. THen
    'we want to check another square, the on left. Is it a tim-1 one ?
    If did = False Then
        If LastX > 0 Then
            If Pos(LastX - 1, LastY) = 100 + tim Then
                LastX = LastX - 1
                did = True
            End If
        End If
    End If
    
    'We check the one below it
    If did = False Then
        If LastY < Map(MapNum).MaxY Then
            If Pos(LastX, LastY + 1) = 100 + tim Then
                LastY = LastY + 1
                did = True
            End If
        End If
    End If
    
    'And above it. One of these have to be it, since we have found the solution, we know that already
    'there is a way back.
    If did = False Then
        If LastY > 0 Then
            If Pos(LastX, LastY - 1) = 100 + tim Then
                LastY = LastY - 1
            End If
        End If
    End If
    
    path(tim).X = LastX
    path(tim).Y = LastY
    
    'Now we loop back and decrease tim, and look for the next square with lower value
    DoEvents
Loop

'Ok we got a path. Now, lets look at the first step and see what direction we should take.
If path(1).X > LastX Then
    FindNpcPath = DIR_RIGHT
ElseIf path(1).Y > LastY Then
    FindNpcPath = DIR_DOWN
ElseIf path(1).Y < LastY Then
    FindNpcPath = DIR_UP
ElseIf path(1).X < LastX Then
    FindNpcPath = DIR_LEFT
End If

End Function

Public Sub CacheMapBlocks(MapNum As Long)
Dim X As Long, Y As Long
    ReDim MapBlocks(MapNum).Blocks(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            If NpcTileIsOpen(MapNum, X, Y) = False Then
                MapBlocks(MapNum).Blocks(X, Y) = 9
            End If
        Next
    Next
End Sub

Public Sub UpdateMapBlock(MapNum, X, Y, blocked As Boolean)
    On Error Resume Next
    If blocked Then
        MapBlocks(MapNum).Blocks(X, Y) = 9
    Else
        MapBlocks(MapNum).Blocks(X, Y) = 0
    End If
End Sub

Function IsOneBlockAway(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Boolean
    If x1 = x2 Then
        If y1 = y2 - 1 Or y1 = y2 + 1 Then
            IsOneBlockAway = True
        Else
            IsOneBlockAway = False
        End If
    ElseIf y1 = y2 Then
        If x1 = x2 - 1 Or x1 = x2 + 1 Then
            IsOneBlockAway = True
        Else
            IsOneBlockAway = False
        End If
    Else
        IsOneBlockAway = False
    End If
End Function

Function GetNpcDir(X As Long, Y As Long, x1 As Long, y1 As Long) As Long
    Dim i As Long, Distance As Long
    
    i = DIR_RIGHT
    
    If X - x1 > 0 Then
        If X - x1 > Distance Then
            i = DIR_RIGHT
            Distance = X - x1
        End If
    ElseIf X - x1 < 0 Then
        If ((X - x1) * -1) > Distance Then
            i = DIR_LEFT
            Distance = ((X - x1) * -1)
        End If
    End If
    
    If Y - y1 > 0 Then
        If Y - y1 > Distance Then
            i = DIR_DOWN
            Distance = Y - y1
        End If
    ElseIf Y - y1 < 0 Then
        If ((Y - y1) * -1) > Distance Then
            i = DIR_UP
            Distance = ((Y - y1) * -1)
        End If
    End If
    
    GetNpcDir = i
        
End Function
Sub UpdatePlayersEsotericas()
Dim i As Long
        
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).EsoBonus > 0 Then
                Player(i).EsoTime = Player(i).EsoTime - 1
                If Player(i).EsoTime <= 0 Then
                    Player(i).EsoTime = 0
                    Player(i).EsoBonus = 0
                    Player(i).EsoNum = 0
                    PlayerMsg i, printf("Sua esoterica terminou!"), brightred
                End If
                SendPlayerData i
            End If
        End If
    Next i
    
End Sub
Function NeedWalk(ByVal MapNum As Long, MapNPCNum As Long) As Boolean
    NeedWalk = True
    If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Ranged = 1 Then
        If isInRange(Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Range, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, GetPlayerX(MapNpc(MapNum).Npc(MapNPCNum).Target), GetPlayerY(MapNpc(MapNum).Npc(MapNPCNum).Target)) Then
            NeedWalk = False
        End If
    End If
End Function
Sub UpdateTransportes()
    Dim i As Long, n As Long
    If UZ Then Exit Sub
    If IsTransporteEmpty() Then Exit Sub
    
    For i = 1 To UBound(Transporte)
        If Transporte(i).State = 0 Then
            If Transporte(i).IntervalTravel + Transporte(i).Tick < GetTickCount Then
                Transporte(i).State = 1
                Transporte(i).Tick = GetTickCount
                
                'Levar ao destino
                If PlayersOnMap(Transporte(i).TravelMap) Then
                    For n = 1 To Player_HighIndex
                        If Player(n).Map = Transporte(i).TravelMap And TempPlayer(n).inDevSuite = 0 Then
                            Call PlayerWarp(n, Transporte(i).DestinyMap, Transporte(i).DestinyX, Transporte(i).DestinyY)
                            If i = 1 Then Call SendPlaySound(n, "airplaneleave.mp3")
                        End If
                    Next n
                End If
                
                If Transporte(i).AlterMap > 0 Then
                    Call SwitchValue(Transporte(i).Map, Transporte(i).AlterMap)
                    Call SwitchValue(Transporte(i).DestinyMap, Transporte(i).AlterDestinyMap)
                    Call SwitchValue(Transporte(i).DestinyX, Transporte(i).AlterDestinyX)
                    Call SwitchValue(Transporte(i).DestinyY, Transporte(i).AlterDestinyY)
                End If
                
                Call SendTransporteCome(Val(i))
            End If
        End If
        If Transporte(i).State = 1 Then
            If Transporte(i).IntervalWait + Transporte(i).Tick < GetTickCount Then
                Transporte(i).State = 0
                Transporte(i).Tick = GetTickCount
                Call SendTransporteCome(Val(i), 2)
                
                'Levar a viagem
                If PlayersOnMap(Transporte(i).LoadMap) Then
                    For n = 1 To Player_HighIndex
                        If Player(n).Map = Transporte(i).LoadMap And TempPlayer(n).inDevSuite = 0 Then
                            Call PlayerWarp(n, Transporte(i).TravelMap, Player(n).X, Player(n).Y)
                            If Transporte(i).Sound <> "" Then Call SendPlaySound(n, Transporte(i).Sound)
                        End If
                    Next n
                End If
            End If
        End If
    Next i
End Sub

Sub UpdateProvacoes()
    Dim i As Long, n As Long
    For i = 1 To ProvaçãoCount
        If PlayersOnMap(Provação(i).Map) > 0 Then
            'Hora de sair cambada!
            If Provação(i).ActualTick + 600000 < GetTickCount Then
                For n = 1 To Player_HighIndex
                    If GetPlayerMap(n) = Provação(i).Map And TempPlayer(n).inDevSuite = 0 Then
                        PlayerMsg n, printf("O tempo acabou e você não conseguiu completar sua provação!"), brightred
                        PlayerWarp n, START_MAP, START_X, START_Y
                        Call SendProvacaoState(n, 0)
                    End If
                Next n
                Exit Sub
            Else
                If Provação(i).ProvaçãoIndex > 0 Then
                    If Not ProvaçãoEnded(i) Then
                        If canSpawnWave(i) Then
                            Call SpawnWave(i)
                        End If
                    Else
                        Call CompleteProvação(i)
                    End If
                End If
            End If
        End If
    Next i
End Sub

Sub SpawnWaveNPC(ByVal MapNum As Long, ByVal MapNPCNum As Long, NpcNum As Long, Optional TargetIndex As Long = 0, Optional Point As Long = 0)
    MapNpc(MapNum).Npc(MapNPCNum).Num = NpcNum
    MapNpc(MapNum).Npc(MapNPCNum).Points = Point
    Map(MapNum).Npc(MapNPCNum) = NpcNum
    SpawnNpc MapNPCNum, MapNum, True
    MapNpc(MapNum).Npc(MapNPCNum).Target = TargetIndex
    If TargetIndex > 0 Then
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = TARGET_TYPE_PLAYER
    Else
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = TARGET_TYPE_NONE
    End If
End Sub

Sub SpawnWave(ByVal ProvNum As Byte)
    'Limpar
    Dim i As Long, WaveNum As Long
    Provação(ProvNum).ActualWave = Provação(ProvNum).ActualWave + 1
    WaveNum = Provação(ProvNum).ActualWave
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(Provação(ProvNum).Map).Npc(i).Num > 0 Then MapNpc(Provação(ProvNum).Map).Npc(i).Num = 0
        If Provação(ProvNum).Wave(WaveNum).Enemy(i).Num > 0 Then
            MapNpc(Provação(ProvNum).Map).Npc(i).Num = Provação(ProvNum).Wave(WaveNum).Enemy(i).Num
            Map(Provação(ProvNum).Map).Npc(i) = Provação(ProvNum).Wave(WaveNum).Enemy(i).Num
            MapNpc(Provação(ProvNum).Map).Npc(i).Level = Provação(ProvNum).MinLevel
            SpawnNpc i, Provação(ProvNum).Map, True
            MapNpc(Provação(ProvNum).Map).Npc(i).Target = Provação(ProvNum).ProvaçãoIndex
            MapNpc(Provação(ProvNum).Map).Npc(i).TargetType = TARGET_TYPE_PLAYER
            
            Dim X As Long, Y As Long
            Dim XLimit As Long, YLimit As Long
            Dim SelectedIndex As Long
            SelectedIndex = Provação(ProvNum).ProvaçãoIndex
            X = GetPlayerX(SelectedIndex) - 10
            If X < 0 Then X = 0
            Y = GetPlayerY(SelectedIndex) - 10
            If Y < 0 Then Y = 0
            XLimit = GetPlayerX(SelectedIndex) + 10
            If XLimit > Map(Provação(ProvNum).Map).MaxX Then XLimit = Map(Provação(ProvNum).Map).MaxX
            YLimit = GetPlayerY(SelectedIndex) + 10
            If YLimit > Map(Provação(ProvNum).Map).MaxY Then YLimit = Map(Provação(ProvNum).Map).MaxY
            
SelectPosition:
            MapNpc(Provação(ProvNum).Map).Npc(i).X = rand(X, XLimit)
            MapNpc(Provação(ProvNum).Map).Npc(i).Y = rand(Y, YLimit)
            If Map(Provação(ProvNum).Map).Tile(MapNpc(Provação(ProvNum).Map).Npc(i).X, MapNpc(Provação(ProvNum).Map).Npc(i).Y).Type <> TileType.TILE_TYPE_WALKABLE Then GoTo SelectPosition
            SendMapNpcXY i, Provação(ProvNum).Map
            
            SendAnimation Provação(ProvNum).Map, SpawnAnim, MapNpc(Provação(ProvNum).Map).Npc(i).X, MapNpc(Provação(ProvNum).Map).Npc(i).Y, MapNpc(Provação(ProvNum).Map).Npc(i).Dir
            
            'SpawnWaveNPC Provação(ProvNum).Map, i, Provação(ProvNum).Wave(WaveNum).Enemy(i).Num, Provação(ProvNum).ProvaçãoIndex
        End If
    Next i
End Sub

Sub CompleteProvação(ByVal ProvNum As Byte)
    Dim Index As Long
    Index = Provação(ProvNum).ProvaçãoIndex
    Call GlobalMsg(GetPlayerName(Index) & " completou sua provação! Vamos parabenizá-lo!", brightgreen)
    If Provação(ProvNum).TradeItem > 0 Then Call TakeInvItem(Index, Provação(ProvNum).TradeItem, 1)
    Call GiveInvItem(Index, Provação(ProvNum).RewardItem, 1, True)
    Call TakeInvItem(Index, 114, 1)
    Player(Index).NumServices = 0
    Call GivePlayerEXP(Index, Provação(ProvNum).RewardXP)
    Call PlayerWarp(Index, START_MAP, START_X, START_Y)
    Call SendProvacaoState(Index, 0)
    SendPlayerData Index
End Sub

Function canSpawnWave(ByVal ProvNum As Byte) As Boolean
    If isWaveCleared(ProvNum) Then
        If Provação(ProvNum).ActualWave + 1 <= UBound(Provação(ProvNum).Wave) Then
            If Provação(ProvNum).ActualTick + Provação(ProvNum).Wave(Provação(ProvNum).ActualWave + 1).WaveTimer < GetTickCount Then
                canSpawnWave = True
                Exit Function
            End If
        End If
    End If
    canSpawnWave = False
End Function

Function ProvaçãoEnded(ByVal ProvNum As Byte) As Boolean
    If isWaveCleared(ProvNum) Then
        If Provação(ProvNum).ActualWave + 1 > UBound(Provação(ProvNum).Wave) Then
            ProvaçãoEnded = True
            Exit Function
        End If
    End If
    ProvaçãoEnded = False
End Function

Function isWaveCleared(ByVal ProvNum As Byte) As Boolean
    Dim i As Long
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(Provação(ProvNum).Map).Npc(i).Num > 0 Then
            isWaveCleared = False
            Exit Function
        End If
    Next i
    isWaveCleared = True
End Function

Function isProvação(ByVal MapNum As Long) As Boolean
    Dim i As Long
    For i = 1 To ProvaçãoCount
        If Provação(i).Map = MapNum Then
            isProvação = True
            Exit Function
        End If
    Next i
    If UZ Then
        If MapNum >= PlanetStart And MapNum <= PlanetStart + MAX_PLANET_BASE Then
            isProvação = True
            Exit Function
        End If
    End If
    isProvação = False
End Function

Function getProvação(ByVal MapNum As Long) As Long
    Dim i As Long
    For i = 1 To ProvaçãoCount
        If Provação(i).Map = MapNum Then
            getProvação = i
            Exit Function
        End If
    Next i
    getProvação = 0
End Function
