Attribute VB_Name = "modPlayer"
Option Explicit

Public Sub DoRedims()
    ReDim Player(1 To MAX_PLAYERS) As PlayerRec
    ReDim Bank(1 To MAX_PLAYERS) As BankRec
    ReDim TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
End Sub

Sub HandleUseChar(ByVal Index As Long)
    If Not IsPlaying(Index) Then
        Call JoinGame(Index)
        Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".", ChatPlayer)
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal Index As Long)
    Dim i As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True
    TempPlayer(Index).MatchIndex = 0
    TempPlayer(Index).Speed = 4
    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(4) = GetPlayerLevel(Index)
    
    ' send the login ok
    SendLoginOk Index
    
    'Mission
    HandleMission Index
    HandleLastLogin Index
    SendPlayerDailyQuest Index
    SendPlanets Index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    TempPlayer(Index).LastMove = GetTickCount
    
    If UZ Then
        If Player(Index).GravityHours > 0 Then
            HandleGravity Index
        Else
            If GetPlayerMap(Index) <> MapRespawn Then
                Call SetPlayerSprite(Index, GetPlayerNormalSprite(Index))
                If Player(Index).PlanetNum = 0 Then
                    Call PlayerWarp(Index, 2, Map(2).MaxX / 2, Map(2).MaxY / 2)
                Else
                    If Player(Index).PlanetNum <= GetMaxPlayerPlanets Then
                        Call PlayerWarp(Index, PlayerPlanet(Player(Index).PlanetNum).PlanetData.Map, Map(PlayerPlanet(Player(Index).PlanetNum).PlanetData.Map).MaxX / 2, Map(PlayerPlanet(Player(Index).PlanetNum).PlanetData.Map).MaxY / 2)
                        SendMap Index, PlayerPlanet(Player(Index).PlanetNum).PlanetData.Map
                        SendPlayerPlanets Index
                    End If
                End If
            Else
                TempPlayer(Index).RespawnTick = GetTickCount
            End If
        End If
    End If
    
    If Player(Index).Guild > 0 Then
        'Verificar se não foi expulso da guild enquanto offline
        If GetPlayerGuildIndex(Index) = 0 Then
            SendDialogue Index, "Notificação de Guild", "Você foi expulso da sua guild " & Trim$(Guild(Player(Index).Guild).Name)
            Player(Index).Guild = 0
        End If
    End If
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call CheckVip(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendAnimations(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendResources(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    Call SendPlayerSpells(Index)
    Call SendHotbar(Index)
    Call SendEffects(Index)
    Call SendQuests(Index)
    Call SendGuilds(Index)
    Call SendPlayerQuests(Index)
    Call SendConquistas(Index)
    Call SendPlayerConquistas(Index)
    
    If ShenlongActive = 1 Then SendShenlong ShenlongMap, 1, 0, , ShenlongX, ShenlongY
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, i)
    Next
    
    For i = 1 To MAX_EVENTS
        Call Events_SendEventData(Index, i)
        Call SendMapKey(Index, Player(Index).EventOpen(i), i)
    Next
    
    SendEXP Index
    Call SendStats(Index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(Index) > ADMIN_MONITOR Then
        'Call GlobalMsg(printf("%s está online!", GetPlayerName(Index)), white)
    End If
    
    ' Send welcome messages
    Call SendWelcome(Index)

    ' Send Resource cache
    If GetPlayerMap(Index) > 0 Then
        'For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            SendResourceCacheTo Index
        'Next
    End If
    
    ' Send the flag so they know they can start doing stuff
    SendInGame Index
    'Call SendNews(Index, 1)
End Sub

Sub LeftGame(ByVal Index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False
        If TempPlayer(Index).MatchIndex > 0 Then RemoveFromMatchData Index
        If TempPlayer(Index).PlanetService > 0 Then PlanetInService(TempPlayer(Index).PlanetService) = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If
        
        'Provações
        Call CheckProvacoesMap(Index)
        RemoveTemporaryItems Index
        
        ' cancel any trade they're in
        If TempPlayer(Index).InTrade > 0 Then
            tradeTarget = TempPlayer(Index).InTrade
            PlayerMsg tradeTarget, printf("%s negou a troca.", Trim$(GetPlayerName(Index))), brightred
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave Index

        ' save and clear data.
        Call SavePlayer(Index)
        Call SaveBank(Index)
        Call ClearBank(Index)

        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
        '    Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " desconectou!", White)
        End If

        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & Options.Game_Name & ".", ChatPlayer)
        Call SendLeftGame(Index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(Index)
End Sub

Function GetPlayerProtection(ByVal Index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(Index, Armor)
    Helm = GetPlayerEquipment(Index, helmet)
    GetPlayerProtection = (GetPlayerStat(Index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(Index, Weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Strength) \ 2) + (GetPlayerLevel(Index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(Index, shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Endurance) \ 2) + (GetPlayerLevel(Index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If X > Map(MapNum).MaxX Then X = Map(MapNum).MaxX
    If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    
    ' if same map then just send their co-ordinates
    If MapNum = GetPlayerMap(Index) Then
        SendPlayerXYToMap Index
    End If
    
    ' clear target
    TempPlayer(Index).Target = 0
    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
    If UZ Then
        If TempPlayer(Index).MatchIndex > 0 Then
            If Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Map <> MapNum Then
                RemoveFromMatchData Index
            End If
        End If
    End If
    SendTarget Index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If OldMap <> MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If
    
    UpdateMapBlock OldMap, GetPlayerX(Index), GetPlayerY(Index), False
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)
    UpdateMapBlock MapNum, X, Y, True
    
    If UZ Then
        Dim PlanetNum As Long
        PlanetNum = PlayerMapIndex(MapNum)
        If PlanetNum > 0 Then
            SendSaibamans Index, PlanetNum
        End If
    End If
    
    ' send player's equipment to new map
    SendMapEquipment Index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(MapNum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    SendMapEquipmentTo i, Index
                    If isAFK(i) Then
                        SendPlayerAFK i, 1
                    End If
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If OldMap > 0 Then
        If GetTotalMapPlayers(OldMap) = 0 Then
            PlayersOnMap(OldMap) = NO
    
            ' Regenerate all NPCs' health
            For i = 1 To MAX_MAP_NPCS
    
                If MapNpc(OldMap).Npc(i).Num > 0 Then
                    MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetNpcMaxVital(OldMap, i, Vitals.HP)
                End If
    
            Next
    
        End If
    End If
    
    If Not IsTransporteEmpty() Then
        For i = 1 To UBound(Transporte)
            If Transporte(i).Map = MapNum Then
                If Transporte(i).State = 1 Then
                    SendTransporteComeTo Val(i), Index
                End If
            End If
        Next i
    End If
    
    If Player(Index).Map > 0 Then
    If Map(Player(Index).Map).Fly = 1 And TempPlayer(Index).Fly = 1 Then
        TempPlayer(Index).Fly = 0
        Call SendPlayerFly(Index)
    End If
    End If
    
    If UZ Then
        If MapNum = MapRespawn Then
            TempPlayer(Index).RespawnTick = GetTickCount
        End If
        If (OldMap = ViagemMap Or OldMap = 53 Or OldMap = 54) And (MapNum <> ViagemMap And MapNum <> 53 And MapNum <> 54) Then
            Call SetPlayerSprite(Index, GetPlayerNormalSprite(Index))
        End If
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    TempPlayer(Index).GettingMap = YES
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong MapNum
    Buffer.WriteLong Map(MapNum).Revision
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, MapNum As Long, i As Long
    Dim X As Long, Y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, Amount As Long, begineventprocessing As Boolean
    Dim CheckForResource As Boolean

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If
    
    CheckForResource = True

    Call SetPlayerDir(Index, Dir)
    Moved = NO
    MapNum = GetPlayerMap(Index)
    
    Dim DirMoveY As Long, DirMoveX As Long
    
    Select Case Dir
        Case DIR_UP: DirMoveY = -1
        Case DIR_DOWN: DirMoveY = 1
        Case DIR_RIGHT: DirMoveX = 1
        Case DIR_LEFT: DirMoveX = -1
    End Select
    
    On Error Resume Next
    
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + DirMoveX, GetPlayerY(Index) + DirMoveY).Type = TILE_TYPE_RESOURCE Then
        If Resource(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + DirMoveX, GetPlayerY(Index) + DirMoveY).data1).ResourceType = 3 Then
            CheckForResource = False
            If GetPlayerEquipment(Index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(Index, Weapon)).data3 = Resource(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + DirMoveX, GetPlayerY(Index) + DirMoveY).data1).ToolRequired Then
                    Call SendFishingTime(Index)
                End If
            End If
        End If
    End If
    
    Select Case Dir
        Case DIR_UP_LEFT

            ' Check to make sure not outside of boundries

            If GetPlayerY(Index) > 0 Or GetPlayerX(Index) > 0 Then



                ' Check to make sure that the tile is walkable

                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_UP + 1) And Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_LEFT + 1) Then

                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) - 1).Type <> tile_type_blocked Or TempPlayer(Index).Fly = 1 Then

                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) - 1).Type <> TILE_TYPE_RESOURCE Or TempPlayer(Index).Fly = 1 Then

   

                            ' Check to see if the tile is a key and if it is check if its opened

                            'If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index) - 1) = YES) Then

                                Call SetPlayerY(Index, GetPlayerY(Index) - 1)

                                Call SetPlayerX(Index, GetPlayerX(Index) - 1)

                                SendPlayerMove Index, movement, sendToSelf

                                Moved = YES

                            'End If

                        End If

                    End If

                End If



            Else



                ' Check to see if we can move them to the another map

                If Map(GetPlayerMap(Index)).Up > 0 And Map(GetPlayerMap(Index)).Left > 0 Then

                    NewMapY = Map(Map(GetPlayerMap(Index)).Up).MaxY

                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), NewMapY)

                    Moved = YES

                    ' clear their target

                    TempPlayer(Index).Target = 0

                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE

                    SendTarget Index

                End If

            End If

           

            Case DIR_UP_RIGHT

            ' Check to make sure not outside of boundries

            If GetPlayerY(Index) > 0 Or GetPlayerX(Index) < Map(MapNum).MaxX Then



                ' Check to make sure that the tile is walkable

                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_UP + 1) And Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_RIGHT + 1) Then

                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) - 1).Type <> tile_type_blocked Or TempPlayer(Index).Fly = 1 Then

                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) - 1).Type <> TILE_TYPE_RESOURCE Or TempPlayer(Index).Fly = 1 Then

   

                            ' Check to see if the tile is a key and if it is check if its opened

                            'If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index) - 1) = YES) Then

                                Call SetPlayerY(Index, GetPlayerY(Index) - 1)

                                Call SetPlayerX(Index, GetPlayerX(Index) + 1)

                                SendPlayerMove Index, movement, sendToSelf

                                Moved = YES

                            'End If

                        End If

                    End If

                End If



            Else



                ' Check to see if we can move them to the another map

                If Map(GetPlayerMap(Index)).Up > 0 And Map(GetPlayerMap(Index)).Right > 0 Then

                    NewMapY = Map(Map(GetPlayerMap(Index)).Up).MaxY

                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), NewMapY)

                    Moved = YES

                    ' clear their target

                    TempPlayer(Index).Target = 0

                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE

                    SendTarget Index

                End If

            End If

           

            Case DIR_DOWN_LEFT

            ' Check to make sure not outside of boundries

            If GetPlayerY(Index) < Map(MapNum).MaxY Or GetPlayerX(Index) > 0 Then



                ' Check to make sure that the tile is walkable

                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_DOWN + 1) And Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_LEFT + 1) Then

                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) + 1).Type <> tile_type_blocked Or TempPlayer(Index).Fly = 1 Then

                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index) + 1).Type <> TILE_TYPE_RESOURCE Or TempPlayer(Index).Fly = 1 Then

   

                            ' Check to see if the tile is a key and if it is check if its opened

                            'If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index) + 1) = YES) Then

                                Call SetPlayerY(Index, GetPlayerY(Index) + 1)

                                Call SetPlayerX(Index, GetPlayerX(Index) - 1)

                                SendPlayerMove Index, movement, sendToSelf

                                Moved = YES

                            'End If

                        End If

                    End If

                End If



            Else



                ' Check to see if we can move them to the another map

                If Map(GetPlayerMap(Index)).Down > 0 And Map(GetPlayerMap(Index)).Left > 0 Then

                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)

                    Moved = YES

                    ' clear their target

                    TempPlayer(Index).Target = 0

                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE

                    SendTarget Index

                End If

            End If

           

            Case DIR_DOWN_RIGHT

            ' Check to make sure not outside of boundries

            If GetPlayerY(Index) < Map(MapNum).MaxY Or GetPlayerX(Index) < Map(MapNum).MaxX Then



                ' Check to make sure that the tile is walkable

                If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_DOWN + 1) And Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_RIGHT + 1) Then

                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) + 1).Type <> tile_type_blocked Or TempPlayer(Index).Fly = 1 Then

                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index) + 1).Type <> TILE_TYPE_RESOURCE Or TempPlayer(Index).Fly = 1 Then

   

                            ' Check to see if the tile is a key and if it is check if its opened

                            'If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index) + 1) = YES) Then

                                Call SetPlayerY(Index, GetPlayerY(Index) + 1)

                                Call SetPlayerX(Index, GetPlayerX(Index) + 1)

                                SendPlayerMove Index, movement, sendToSelf

                                Moved = YES

                            'End If

                        End If

                    End If

                End If



            Else



                ' Check to see if we can move them to the another map

                If Map(GetPlayerMap(Index)).Down > 0 And Map(GetPlayerMap(Index)).Right > 0 Then

                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)

                    Moved = YES

                    ' clear their target

                    TempPlayer(Index).Target = 0

                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE

                    SendTarget Index

                End If

            End If
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then
                If TempPlayer(Index).inDevSuite = YES Or TempPlayer(Index).Fly = 1 Then
                    Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                    SendPlayerMove Index, movement, sendToSelf
                    Moved = YES
                Else
                    ' Check to make sure that the tile is walkable
                    If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_UP + 1) Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> tile_type_blocked Then
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_RESOURCE Or CheckForResource = False Then
        
                                ' Check to see if the tile is a event and if it is check if its opened
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_EVENT Then
                                    Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                    SendPlayerMove Index, movement, sendToSelf
                                    Moved = YES
                                Else
                                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).data1 > 0 Then
                                        If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).data1) = YES) Then
                                            Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                                            SendPlayerMove Index, movement, sendToSelf
                                            Moved = YES
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(Index)).Up).MaxY
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                    SendTarget Index
                End If
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < Map(MapNum).MaxY Then
                If TempPlayer(Index).inDevSuite = YES Or TempPlayer(Index).Fly = 1 Then
                    Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                    SendPlayerMove Index, movement, sendToSelf
                    Moved = YES
                Else
                    ' Check to make sure that the tile is walkable
                    If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_DOWN + 1) Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> tile_type_blocked Then
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_RESOURCE Or CheckForResource = False Then
        
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_EVENT Then
                                    Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                    SendPlayerMove Index, movement, sendToSelf
                                    Moved = YES
                                Else
                                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).data1 > 0 Then
                                        If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).data1) = YES) Then
                                            Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                                            SendPlayerMove Index, movement, sendToSelf
                                            Moved = YES
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Down > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                    SendTarget Index
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) > 0 Then
                If TempPlayer(Index).inDevSuite = YES Or TempPlayer(Index).Fly = 1 Then
                    Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                    SendPlayerMove Index, movement, sendToSelf
                    Moved = YES
                Else
                    ' Check to make sure that the tile is walkable
                    If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_LEFT + 1) Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> tile_type_blocked Then
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_RESOURCE Or CheckForResource = False Then
        
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_EVENT Then
                                    Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                    SendPlayerMove Index, movement, sendToSelf
                                    Moved = YES
                                Else
                                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).data1 > 0 Then
                                        If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).data1) = YES) Then
                                            Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                                            SendPlayerMove Index, movement, sendToSelf
                                            Moved = YES
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(Index)).Left).MaxX
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, NewMapX, GetPlayerY(Index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                    SendTarget Index
                End If
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < Map(MapNum).MaxX Then
                If TempPlayer(Index).inDevSuite = YES Or TempPlayer(Index).Fly = 1 Then
                    Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                    SendPlayerMove Index, movement, sendToSelf
                    Moved = YES
                Else
                    ' Check to make sure that the tile is walkable
                    If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, DIR_RIGHT + 1) Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> tile_type_blocked Then
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_RESOURCE Or CheckForResource = False Then
        
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_EVENT Then
                                    Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                    SendPlayerMove Index, movement, sendToSelf
                                    Moved = YES
                                Else
                                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).data1 > 0 Then
                                        If Events(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).data1).WalkThrought = YES Or (Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).data1) = YES) Then
                                            Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                                            SendPlayerMove Index, movement, sendToSelf
                                            Moved = YES
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Right > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(Index).Target = 0
                    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                    SendTarget Index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = tile_type_warp And Not TempPlayer(Index).Fly = 1 Then
            MapNum = .data1
            X = .data2
            Y = .data3
            Call PlayerWarp(Index, MapNum, X, Y)
            Moved = YES
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            X = .data1
            If X > 0 Then ' shop exists?
                If Len(Trim$(Shop(X).Name)) > 0 Then ' name exists?
                    SendOpenShop Index, X
                    TempPlayer(Index).InShop = X ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank Index
            TempPlayer(Index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .data1
            Amount = .data2
            If Not GetPlayerVital(Index, VitalType) = GetPlayerMaxVital(Index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = brightgreen
                Else
                    Colour = brightblue
                End If
                SendActionMsg GetPlayerMap(Index), "+" & Amount, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
                SetPlayerVital Index, VitalType, GetPlayerVital(Index, VitalType) + Amount
                PlayerMsg Index, "You feel rejuvinating forces flowing through your boy.", brightgreen
                Call SendVital(Index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            Amount = .data1
            SendActionMsg GetPlayerMap(Index), "-" & Amount, brightred, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32, 1
            If GetPlayerVital(Index, HP) - Amount <= 0 Then
                KillPlayer Index
                PlayerMsg Index, "You're killed by a trap.", brightred
            Else
                SetPlayerVital Index, HP, GetPlayerVital(Index, HP) - Amount
                PlayerMsg Index, "You're injured by a trap.", brightred
                Call SendVital(Index, HP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
            End If
            Moved = YES
        End If
        
                ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            ForcePlayerMove Index, MOVING_WALKING, .data1
            Moved = YES
        End If
        
        'Event
        If .Type = TILE_TYPE_EVENT Then
            If .data1 > 0 Then
                If Events(.data1).Trigger = 0 Then
                    InitEvent Index, .data1
                End If
            End If
            Moved = YES
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
    Else
        If movement = 2 Then
            Call SetPlayerVital(Index, MP, GetPlayerVital(Index, MP) - 1)
            Call SendVital(Index, MP)
        End If
        PlayerMapGetItem Index, True
    End If
    
    X = GetPlayerX(Index)
    Y = GetPlayerY(Index)
    
    If UZ Then
        If GetPlayerMap(Index) = ViagemMap Or GetPlayerMap(Index) = 53 Or GetPlayerMap(Index) = 54 Then
            For i = 1 To MAX_PLANETS
                If Planets(i).X = X Then
                    If Planets(i).Y = Y Then
                        'If i <= MAX_PLANETS Then
                            If Planets(i).State <> 1 And InLevel(Index, Planets(i).Level) Then
                                EnterPlanet Index, i
                            End If
                        'Else
                        '    PlayerWarp Index, TesouroMap, 12, 9
                        'End If
                        Exit Sub
                    End If
                End If
            Next i
        End If
        If GetPlayerMap(Index) = VirgoMap Then
            For i = 1 To UBound(PlayerPlanet)
                If PlayerPlanet(i).PlanetData.X = X Then
                    If PlayerPlanet(i).PlanetData.Y = Y Then
                        If Trim$(LCase(PlayerPlanet(i).PlanetData.Owner)) = Trim$(LCase(GetPlayerName(Index))) Or InPartyWith(Index, Trim$(LCase(PlayerPlanet(i).PlanetData.Owner))) Then
                            PlayerWarp Index, PlayerPlanet(i).PlanetData.Map, 7, 14
                            Call SetPlayerSprite(Index, GetPlayerNormalSprite(Index))
                            Exit Sub
                        Else
                            PlayerMsg Index, "Este planeta não é seu! Por enquanto não liberamos invasões!", brightred
                            Exit Sub
                        End If
                    End If
                End If
            Next i
        End If
    End If

End Sub

Sub ForcePlayerMove(ByVal Index As Long, ByVal movement As Long, ByVal Direction As Long)

    If Direction < DIR_UP Or Direction > DIR_DOWN_RIGHT Then Exit Sub

    If movement < 1 Or movement > 2 Then Exit Sub

   

    Select Case Direction

        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub

        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub

        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub

        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub

        Case DIR_UP_LEFT

            If GetPlayerY(Index) = 0 And GetPlayerX(Index) = 0 Then Exit Sub

        Case DIR_UP_RIGHT

            If GetPlayerY(Index) = 0 And GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub

        Case DIR_DOWN_LEFT

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY And GetPlayerX(Index) = 0 Then Exit Sub

        Case DIR_DOWN_RIGHT

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY And GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub

    End Select

   

    PlayerMove Index, Direction, movement, True

End Sub

Sub CheckEquippedItems(ByVal Index As Long)
    Dim slot As Long
    Dim ItemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(Index, i)

        If ItemNum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment Index, 0, i
                Case Equipment.Armor

                    If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment Index, 0, i
                Case Equipment.helmet

                    If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment Index, 0, i
                Case Equipment.shield

                    If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment Index, 0, i
            End Select

        Else
            SetPlayerEquipment Index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable > 0 Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Function
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable > 0 Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function
Function HasItems(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable > 0 Then
                HasItems = GetPlayerInvItemValue(Index, i)
            Else
                HasItems = HasItems + 1
            End If
        End If

    Next

End Function

Function FindItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            FindItem = i
            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable > 0 Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ItemNum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, invSlot)

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable > 0 Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(Index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(Index, invSlot, GetPlayerInvItemValue(Index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(Index, invSlot, 0)
        Call SetPlayerInvItemValue(Index, invSlot, 0)
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, Optional ByVal sendupdate As Boolean = True) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(Index, ItemNum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        If sendupdate Then Call SendInventoryUpdate(Index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(Index, printf("Inventário cheio."), brightred)
        GiveInvItem = False
    End If

End Function
Public Sub SetPlayerItems(ByVal Index As Long, ByVal itemID As Long, ByVal itemCount As Long)
    Dim i As Long
    Dim given As Long
    If Item(itemID).Type = ITEM_TYPE_CURRENCY Then
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = itemID Then
                Call SetPlayerInvItemValue(Index, i, itemCount)
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
        Next i
    End If
    
    For i = 1 To MAX_INV
        If given >= itemCount Then Exit Sub
        If GetPlayerInvItemNum(Index, i) = 0 Then
            Call SetPlayerInvItemNum(Index, i, itemID)
            given = given + 1
            If Item(itemID).Type = ITEM_TYPE_CURRENCY Then
                Call SetPlayerInvItemValue(Index, i, itemCount)
                given = itemCount
            End If
            Call SendInventoryUpdate(Index, i)
        End If
    Next i
End Sub
Public Sub GivePlayerItems(ByVal Index As Long, ByVal itemID As Long, ByVal itemCount As Long)
    Dim i As Long
    Dim given As Long
    If Item(itemID).Type = ITEM_TYPE_CURRENCY Then
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = itemID Then
                Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + itemCount)
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
        Next i
    End If
    
    For i = 1 To MAX_INV
        If given >= itemCount Then Exit Sub
        If GetPlayerInvItemNum(Index, i) = 0 Then
            Call SetPlayerInvItemNum(Index, i, itemID)
            given = given + 1
            If Item(itemID).Type = ITEM_TYPE_CURRENCY Then
                Call SetPlayerInvItemValue(Index, i, itemCount)
                given = itemCount
            End If
            Call SendInventoryUpdate(Index, i)
        End If
    Next i
End Sub
Public Sub TakePlayerItems(ByVal Index As Long, ByVal itemID As Long, ByVal itemCount As Long)
Dim i As Long
    If HasItems(Index, itemID) >= itemCount Then
        If Item(itemID).Type = ITEM_TYPE_CURRENCY Then
            TakeInvItem Index, itemID, itemCount
        Else
            For i = 1 To MAX_INV
                If GetPlayerInvItemNum(Index, i) = itemID Then
                    SetPlayerInvItemNum Index, i, 0
                    SetPlayerInvItemValue Index, i, 0
                    SendInventoryUpdate Index, i
                End If
            Next
        End If
    Else
        PlayerMsg Index, printf("Você precisa [%d] de [%s]", itemCount & "," & Trim$(Item(itemID).Name)), AlertColor
    End If
End Sub
Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long, SpellRealName As String
    Dim SpellLength As Byte

    'Evolutions
    SpellRealName = Trim$(Spell(SpellNum).Name)
    SpellLength = Len(SpellRealName)
    
    If Len(SpellRealName) = 0 Then Exit Function
    
    If Mid(SpellRealName, SpellLength - 1, 1) = "+" Then
        SpellLength = SpellLength - 3
        SpellRealName = Mid(SpellRealName, 1, SpellLength)
    End If

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) > 0 Then
            If GetPlayerSpell(Index, i) = SpellNum Then
                HasSpell = True
                Exit Function
            End If
            
            Dim SpellName As String
            SpellName = Mid(Trim$(Spell(Player(Index).Spell(i)).Name), 1, SpellLength)
        
            If SpellName = SpellRealName Then
                HasSpell = True
                Exit Function
            End If
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal Index As Long, Optional AutoGet As Boolean = False)
    Dim i As Long
    Dim n As Long
    Dim MapNum As Long
    Dim Msg As String

    If Not IsPlaying(Index) Then Exit Sub
    MapNum = GetPlayerMap(Index)

    For i = MAX_MAP_ITEMS To 1 Step -1
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(Index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(MapNum, i).X = GetPlayerX(Index)) Then
                    If (MapItem(MapNum, i).Y = GetPlayerY(Index)) Then
                        If Item(MapItem(MapNum, i).Num).Type = ItemType.ITEM_TYPE_CONSUME Then
                            If UZ And AutoGet Then
                                'Consumir instantaneamente
                                If Item(MapItem(MapNum, i).Num).AddHP > 0 Then
                                    HealPlayer Index, GetPlayerMaxVital(Index, HP) * (Item(MapItem(MapNum, i).Num).AddHP / 100)
                                    SendActionMsg GetPlayerMap(Index), "+" & Item(MapItem(MapNum, i).Num).AddHP & "% HP", brightgreen, 1, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                                    If IsDaily(Index, GetGlobes) Then UpdateDaily Index
                                End If
                                If Item(MapItem(MapNum, i).Num).AddMP > 0 Then
                                    HealPlayerMP Index, GetPlayerMaxVital(Index, MP) * (Item(MapItem(MapNum, i).Num).AddMP / 100)
                                    SendActionMsg GetPlayerMap(Index), "+" & Item(MapItem(MapNum, i).Num).AddMP & "% KI", brightblue, 1, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                                    If IsDaily(Index, GetGlobes) Then UpdateDaily Index
                                End If
                                If Item(MapItem(MapNum, i).Num).AddEXP > 0 Then
                                    If Item(MapItem(MapNum, i).Num).AddEXP = 1 Then 'Next Wave
                                        If TempPlayer(Index).MatchIndex > 0 Then
                                            If Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Type = 0 Then
                                                MatchData(TempPlayer(Index).MatchIndex).WaveTick = -(MatchData(TempPlayer(Index).MatchIndex).WaveTick - GetTickCount)
                                                SendActionMsg GetPlayerMap(Index), "Próxima onda ativada!", brightred, ActionMsgType.ACTIONMSG_STATIC, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                                            End If
                                        End If
                                    End If
                                End If
                                
                                If Item(MapItem(MapNum, i).Num).Animation > 0 Then Call SendAnimation(GetPlayerMap(Index), Item(MapItem(MapNum, i).Num).Animation, 0, 0, GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index)
                                If Item(MapItem(MapNum, i).Num).Effect > 0 Then Call SendEffect(GetPlayerMap(Index), Item(MapItem(MapNum, i).Num).Effect, GetPlayerX(Index), GetPlayerY(Index))
                                
                                ' send the sound
                                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, MapItem(MapNum, i).Num
                                
                                ClearMapItem i, MapNum
                                Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), 0, 0)
                                If Not AutoGet Then Exit Sub
                            End If
                        Else
                            If Not AutoGet Then
                                ' Find open slot
                                n = FindOpenInvSlot(Index, MapItem(MapNum, i).Num)
            
                                ' Open slot available?
                                If n <> 0 Then
                                    ' Set item in players inventor
                                    Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, i).Num)
                                    
                                    If UZ Then
                                        If Item(MapItem(MapNum, i).Num).Type = ITEM_TYPE_DRAGONBALL Then
                                            DragonballInfo(Item(MapItem(MapNum, i).Num).Dragonball) = 0
                                            UpdateDragonballList
                                        End If
                                    End If
            
                                    If Item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, n)).Stackable > 0 Then
                                        Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, i).Value)
                                        Msg = MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                                    Else
                                        Call SetPlayerInvItemValue(Index, n, 0)
                                        Msg = Trim$(Item(GetPlayerInvItemNum(Index, n)).Name)
                                    End If
            
                                    ' Erase item from the map
                                    ClearMapItem i, MapNum
                                    
                                    Call SendInventoryUpdate(Index, n)
                                    Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), 0, 0)
                                    SendActionMsg GetPlayerMap(Index), Msg, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                    If TempPlayer(Index).MatchIndex > 0 Then
                                    If Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Type = 4 Then
                                        If GetPlayerInvItemNum(Index, n) = TesouroItem Then
                                            MatchData(TempPlayer(Index).MatchIndex).Points = GetPlayerInvItemValue(Index, n)
                                            SendMatchData TempPlayer(Index).MatchIndex
                                            If Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).PointsToConquest <= MatchData(TempPlayer(Index).MatchIndex).Points Then
                                                TakeInvItem Index, TesouroItem, GetPlayerInvItemValue(Index, n)
                                                GiveInvItem Index, MoedaZ, Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Preco
                                                SendBossMsg GetPlayerMap(Index), "Parabéns! Você completou o saque! Este planeta explodirá em 10 segundos.", Yellow, Index
                                                PlayerMsg Index, "-- Recompensa pela missão --", brightgreen
                                                PlayerMsg Index, "Você recebeu: " & Int(Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Preco) & "z", brightgreen
                                                GivePlayerEXP Index, Int(ExperienceBase(Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Level))
                                                PlayerMsg Index, "Você recebeu: " & Int(ExperienceBase(Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Level)) & " exp", brightgreen
                                                GivePlayerVIPExp Index, Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Level
                                                If Player(Index).Guild > 0 Then GiveGuildExp Index, Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Level, Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Level
                                                Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Owner = GetPlayerName(Index)
                                                Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).State = 2
                                                Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).TimeToExplode = GetTickCount + 10000
                                                MatchData(TempPlayer(Index).MatchIndex).Active = 0
                                                MatchData(TempPlayer(Index).MatchIndex).WaveTick = GetTickCount + 5000
                                                
                                                For n = 1 To MAX_MAP_NPCS
                                                    DespawnNPC Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Map, n
                                                Next n
                                                
                                                If TempPlayer(Index).PlanetService = MatchData(TempPlayer(Index).MatchIndex).Planet Then
                                                    CompleteService Index
                                                End If
                                            End If
                                        End If
                                    End If
                                    End If
                                    Exit For
                                Else
                                    Call PlayerMsg(Index, printf("Inventário cheio."), brightred)
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal Index As Long, ByVal MapItemNum As Long)
Dim MapNum As Long

    MapNum = GetPlayerMap(Index)
    
    ' no lock or locked to player?
    If MapItem(MapNum, MapItemNum).PlayerName = vbNullString Or MapItem(MapNum, MapItemNum).PlayerName = Trim$(GetPlayerName(Index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If UZ Then
        If GetPlayerMap(Index) = ViagemMap Then Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Or TempPlayer(Index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
            Dim CanDrop As Boolean
            If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_NAVE Then
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(Index, i) > 0 Then
                    If Item(GetPlayerInvItemNum(Index, i)).Type = ITEM_TYPE_NAVE And i <> InvNum Then CanDrop = True
                    End If
                Next i
            Else
                CanDrop = True
            End If
            If CanDrop = False Then
                PlayerMsg Index, "Você não pode dropar sua única nave!", brightred
                Exit Sub
            End If
            i = FindOpenMapItemSlot(GetPlayerMap(Index))

            If i <> 0 Then
                MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), i).Y = GetPlayerY(Index)
                MapItem(GetPlayerMap(Index), i).PlayerName = Trim$(GetPlayerName(Index))
                MapItem(GetPlayerMap(Index), i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(Index), i).canDespawn = True
                MapItem(GetPlayerMap(Index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME

                If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, InvNum)).Stackable > 0 Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                        MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, InvNum)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, InvNum, 0)
                        Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Else
                        MapItem(GetPlayerMap(Index), i).Value = Amount
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), i).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(Index, InvNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, Amount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), Trim$(GetPlayerName(Index)), MapItem(GetPlayerMap(Index), i).canDespawn)
            Else
                Call PlayerMsg(Index, printf("Existem muitos itens no chão."), brightred)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    Do While GetPlayerExp(Index) >= GetPlayerNextLevel(Index)
        expRollover = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
        
        ' can level up?
        If Not SetPlayerLevel(Index, GetPlayerLevel(Index) + 1) Then
            Exit Sub
        End If
        
        Player(Index).PDL = Player(Index).PDL + LevelUpBonus
        Player(Index).Points = Player(Index).Points + 3
        Call SetPlayerExp(Index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        UpdateWebRank LCase(GetPlayerLogin(Index)), "level", GetPlayerLevel(Index)
        UpdateGuildLevel Index
        If Player(Index).PlanetNum > 0 Then
            For i = 1 To GetMaxPlayerPlanets
                If Trim$(LCase(PlayerPlanet(i).PlanetData.Owner)) = Trim$(LCase(GetPlayerName(Index))) Then
                    PlayerPlanet(i).PlanetData.Level = GetPlayerLevel(Index)
                    Exit For
                End If
            Next i
        End If
        PlayerMsg Index, printf("Você passou %d nível(s)!", Val(level_count)), Yellow
        SendStats Index
        SendEXP Index
        SendActionMsg GetPlayerMap(Index), "Level UP!", brightgreen, 1, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        SendAnimation GetPlayerMap(Index), PlayerLevelUpAnim, GetPlayerX(Index), GetPlayerY(Index), GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index
        SavePlayer Index
        frmServer.lvwInfo.ListItems(Index).SubItems(4) = GetPlayerLevel(Index)
    End If
End Sub

Sub CheckPlayerGodLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    If Player(Index).IsGod = 0 Then Exit Sub
    
    Do While Player(Index).GodExp >= GetPlayerGodNextLevel(Index)
        expRollover = Player(Index).GodExp - GetPlayerGodNextLevel(Index)
        
        ' can level up?
        If Player(Index).GodLevel + 1 >= MAX_LEVELS Then
            Player(Index).GodExp = 0
            Player(Index).GodLevel = 1
            Player(Index).IsGod = Player(Index).IsGod + 1
            PlayerMsg Index, "Parabéns! Você adquiriu uma ascenção divina!", Yellow
            GiveInvItem Index, 209, 1
            SendEXP Index
            Exit Sub
        End If
        
        If Player(Index).GodLevel > MAX_LEVELS - 10 Then
            PlayerMsg Index, "Você está próximo de efetuar uma ascenção! Certifique-se de deixar pelo menos 1 slot livre em seu inventário para receber o item de ascenção!", BrightCyan
        End If
        
        'Player(Index).PDL = Player(Index).PDL + LevelUpBonus
        Player(Index).GodExp = expRollover
        Player(Index).GodLevel = Player(Index).GodLevel + 1
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        'UpdateWebRank LCase(GetPlayerLogin(Index)), "level", GetPlayerLevel(Index)
        'UpdateGuildLevel Index
        PlayerMsg Index, printf("Você passou %d nível(s) divinos!", Val(level_count)), Yellow
        SendEXP Index
        SendActionMsg GetPlayerMap(Index), "Level UP!", brightgreen, 1, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        SendAnimation GetPlayerMap(Index), PlayerLevelUpAnim, GetPlayerX(Index), GetPlayerY(Index), GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index
        SavePlayer Index
        frmServer.lvwInfo.ListItems(Index).SubItems(4) = GetPlayerLevel(Index)
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
    
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
End Function

Function SetPlayerLevel(ByVal Index As Long, ByVal Level As Long) As Boolean
    SetPlayerLevel = False
    If Level >= MAX_LEVELS Then Exit Function
    Player(Index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    On Error Resume Next
    GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
End Function

Function GetPlayerGodNextLevel(ByVal Index As Long) As Long
    On Error Resume Next
    If Player(Index).IsGod > 0 Then
        GetPlayerGodNextLevel = Experience(Player(Index).GodLevel)
    Else
        GetPlayerGodNextLevel = 0
    End If
End Function

Function GetPlayerVipNextLevel(ByVal Index As Long) As Long
    On Error Resume Next
    If Player(Index).VIP = 0 Then Exit Function
    GetPlayerVipNextLevel = Experience(Player(Index).VIP)
End Function

Function GetPlayerLastLevel(ByVal Index As Long) As Long
    If GetPlayerLevel(Index) > 1 Then
        GetPlayerLastLevel = Experience(GetPlayerLevel(Index) - 1)
    Else
        GetPlayerLastLevel = 0
    End If
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Exp = Exp
    If GetPlayerLevel(Index) = MAX_LEVELS And Player(Index).Exp > GetPlayerNextLevel(Index) Then
        Player(Index).Exp = GetPlayerNextLevel(Index)
    End If
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Or Index = 0 Then Exit Function
    If Vital > Vital_Count - 1 Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    If GetPlayerVital(Index, Vital) < 0 Then
        Player(Index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal Index As Long, ByVal stat As Stats, Optional NoTrans As Boolean = False) As Long
    Dim X As Long, i As Long
    If Index > MAX_PLAYERS Then Exit Function
    
    X = Player(Index).stat(stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(Index).Equipment(i) > 0 Then
            If Item(Player(Index).Equipment(i)).Add_Stat(stat) > 0 Then
                X = X + Item(Player(Index).Equipment(i)).Add_Stat(stat)
            End If
        End If
    Next
    
    If TempPlayer(Index).Trans > 0 And NoTrans = False Then
        X = X + Spell(TempPlayer(Index).Trans).Add_Stat(stat)
    End If
    
    GetPlayerStat = X
End Function

Public Function GetPlayerRawStat(ByVal Index As Long, ByVal stat As Stats) As Long
    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(Index).stat(stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal stat As Stats, ByVal Value As Long)
    Player(Index).stat(stat) = Value
End Sub

Public Function GetPlayerStatPoints(ByVal Index As Long, ByVal stat As Stats) As Long
    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerStatPoints = Player(Index).statPoints(stat)
End Function

Public Sub SetPlayerStatPoints(ByVal Index As Long, ByVal stat As Stats, ByVal Value As Long)
    Player(Index).statPoints(stat) = Value
    Call CheckPlayerStatLevelUp(Index, stat)
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).Points
    'Dim TotalPoints, i As Long
    'For i = 1 To 5
    '    TotalPoints = TotalPoints + Player(Index).statPoints(i)
    'Next i
    'GetPlayerPOINTS = Player(Index).PDL - TotalPoints
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal Points As Long)
    If Points <= 0 Then Points = 0
    Player(Index).Points = Points
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Map = MapNum
    End If

End Sub

Function GetPlayerX(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(Index).Inv(invSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(invSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(Index).Inv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(invSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(Index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal SpellNum As Long)
    Player(Index).Spell(spellslot) = SpellNum
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot) = InvNum
End Sub

' ToDo
Sub OnDeath(ByVal Index As Long)
    Dim i As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(Index, Vitals.HP, 0)

    ' Warp player away
    Call SetPlayerDir(Index, DIR_DOWN)
    
    TempPlayer(Index).Fly = 0
    SendPlayerFly Index
    
    If TempPlayer(Index).MatchIndex > 0 Then RemoveFromMatchData Index
    
    If Player(Index).Map = 0 Then Exit Sub
    With Map(GetPlayerMap(Index))
        ' to the bootmap if it is set
        If Not .Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_ARENA Then
            If .BootMap > 0 Then
                PlayerWarp Index, .BootMap, .BootX, .BootY
            Else
                Call PlayerWarp(Index, RESPAWN_MAP, RESPAWN_X, RESPAWN_Y)
            End If
        Else
            Call PlayerWarp(Index, .Tile(GetPlayerX(Index), GetPlayerY(Index)).data1, .Tile(GetPlayerX(Index), GetPlayerY(Index)).data2, .Tile(GetPlayerX(Index), GetPlayerY(Index)).data3)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(Index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(Index).spellBuffer.Spell = 0
    TempPlayer(Index).spellBuffer.Timer = 0
    TempPlayer(Index).spellBuffer.Target = 0
    TempPlayer(Index).spellBuffer.tType = 0
    Player(Index).IsDead = 0
    Call SendClearSpellBuffer(Index)
    
    ' Restore vitals
    Call SetPlayerVital(Index, Vitals.HP, GetPlayerMaxVital(Index, Vitals.HP))
    Call SetPlayerVital(Index, Vitals.MP, GetPlayerMaxVital(Index, Vitals.MP))
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
    End If
    
    Call SendPlayerData(Index)

End Sub

Sub CheckResource(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, Optional InflictDamage As Long = 0)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    Dim n As Long
    
    
    
    'Pesca
    If Not UZ Then
        If Map(Player(Index).Map).Tile(Player(Index).X, Player(Index).Y).Type = TILE_TYPE_RESOURCE Then
            If Resource(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1).ResourceType = 3 Then
                X = GetPlayerX(Index)
                Y = GetPlayerY(Index)
            End If
        End If
    End If
    
    If GetPlayerMap(Index) = 0 Then Exit Sub
    If X < 0 Or X > Map(GetPlayerMap(Index)).MaxX Or Y < 0 Or Y > Map(GetPlayerMap(Index)).MaxY Then Exit Sub
    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(Index)).Tile(X, Y).data1
        
        If Resource(Resource_index).ResourceType = 6 Then Exit Sub 'Decoração

        If UZ Then
        Dim PlanetNum As Long
        PlanetNum = PlayerMapIndex(GetPlayerMap(Index))
        If PlanetNum > 0 Then
            If Trim$(LCase(PlayerPlanet(PlanetNum).PlanetData.Owner)) = Trim$(LCase(GetPlayerName(Index))) Then
                If Resource(Resource_index).ItemReward > 0 Then
                    Dim diff As Long, ExtratorIndex As Long
                    ExtratorIndex = GetExtratorIndex(PlanetNum, X, Y)
                    If ExtratorIndex > 0 Then
                        diff = DateDiff("n", PlayerPlanet(PlanetNum).Extrator(ExtratorIndex).TaskInit, Now)
                        diff = Int(diff / 10) * Resource(Resource_index).RespawnTime
                        If Resource(Resource_index).ItemReward = EspV Then
                            diff = (diff / 100) * PlayerPlanet(PlanetNum).PlanetData.EspeciariaVermelha
                        End If
                        If Resource(Resource_index).ItemReward = EspAz Then
                            diff = (diff / 100) * PlayerPlanet(PlanetNum).PlanetData.EspeciariaAzul
                        End If
                        If Resource(Resource_index).ItemReward = EspAm Then
                            diff = (diff / 100) * PlayerPlanet(PlanetNum).PlanetData.EspeciariaAmarela
                        End If
                        If PlayerPlanet(PlanetNum).Extrator(ExtratorIndex).Acc > 0 Then
                            diff = diff * (PlayerPlanet(PlanetNum).Extrator(ExtratorIndex).Acc + 1)
                            If DateDiff("h", PlayerPlanet(PlanetNum).Extrator(ExtratorIndex).AccStart, Now) > 24 Then
                                PlayerPlanet(PlanetNum).Extrator(ExtratorIndex).Acc = 0
                            End If
                        End If
                        If diff > 0 Then
                            If diff > CapacidadeMaxima(Resource(Resource_index).ResourceLevel) Then diff = CapacidadeMaxima(Resource(Resource_index).ResourceLevel)
                            GiveInvItem Index, Resource(Resource_index).ItemReward, diff
                            SendActionMsg GetPlayerMap(Index), "+" & diff & " " & Trim$(Item(Resource(Resource_index).ItemReward).Name), White, ActionMsgType.ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                            PlayerPlanet(PlanetNum).Extrator(ExtratorIndex).TaskInit = Now
                            SavePlayerPlanet PlanetNum
                        End If
                    End If
                End If
            End If
            If Resource(Resource_index).NucleoLevel > 0 Then
                'Teleportar para o mapa da casa
                Player(Index).PlayerHouseNum = GetPlayerMap(Index)
                PlayerWarp Index, 50, 14, 13
            End If
            Exit Sub
        End If
        End If

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count

            If ResourceCache(GetPlayerMap(Index)).ResourceData(i).X = X Then
                If ResourceCache(GetPlayerMap(Index)).ResourceData(i).Y = Y Then
                    Resource_num = i
                End If
            End If

        Next
        
        If Resource_num > 0 Then
            
            If Resource(Resource_index).ToolRequired > 0 Then
                If GetPlayerEquipment(Index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(Index, Weapon)).data3 = Resource(Resource_index).ToolRequired Then
                    
                    Else
                        PlayerMsg Index, printf("Você não tem a ferramenta correta."), brightred
                        Exit Sub
                    End If
                Else
                    PlayerMsg Index, printf("Você precisa de uma ferramenta para interagir com esse recurso."), brightred
                    Exit Sub
                End If
            End If

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(Index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg Index, printf("Inventário cheio."), brightred
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).X
                        rY = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).Y
                        If GetPlayerEquipment(Index, Weapon) > 0 Then
                            Damage = Item(GetPlayerEquipment(Index, Weapon)).data2
                        Else
                            Damage = 1
                        End If
                        If UZ Then
                            If Resource(Resource_index).ResourceType = 0 Or Resource(Resource_index).ResourceType = 1 Then
                                If InflictDamage = 0 Then
                                    Damage = GetPlayerDamage(Index)
                                Else
                                    Damage = InflictDamage
                                End If
                            End If
                            If Resource(Resource_index).ResourceType = 4 Then Damage = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health
                            If Damage <= 0 Then Damage = 1
                        End If
                        
                        ' check if damage is more than health
                        If Damage > 0 Then
                            ' cut it down!
                            If ResourceCache(Player(Index).Map).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                                If UZ Then
                                    If GetPlayerMap(Index) = 49 Then 'Planeta Trash
                                        If EventGlobalType = 3 Then

                                            Dim Total As Long
                                            Dim ItemNum As Long
                                            Total = rand(5, 8)
                                            For i = 1 To Total
                                                X = rand(ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).X - 2, ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).X + 2)
                                                Y = rand(ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).Y - 2, ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).Y + 2)
                                                ItemNum = 144
                                                Call SpawnItem(ItemNum, 1, GetPlayerMap(Index), X, Y)
                                                SendAnimation GetPlayerMap(Index), Item(ItemNum).Animation, X, Y, 0
                                            Next i
                                        End If
                                    End If
                                    If IsDaily(Index, Destroy) Then UpdateDaily Index
                                    If Resource(Resource_index).ResourceType <> 4 Then
                                        Select Case Resource(Resource_index).ItemReward
                                            Case 1 'Spawn elite
                                                If IsDaily(Index, DestroyArmy) Then UpdateDaily Index
                                            Case 2 'Estrela extra
                                                If IsDaily(Index, DestroyMayor) Then UpdateDaily Index
                                            Case 3 'Próxima onda
                                                If IsDaily(Index, DestroyArt) Then UpdateDaily Index
                                            Case 4 'Aumentar preço em 10%
                                                If IsDaily(Index, DestroyGold) Then UpdateDaily Index
                                        End Select
                                    End If
                                    
                                    If TempPlayer(Index).MatchIndex > 0 Then
                                        If Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Type = 0 Then
                                            MatchData(TempPlayer(Index).MatchIndex).Points = MatchData(TempPlayer(Index).MatchIndex).Points + 2
                                            SendActionMsg GetPlayerMap(Index), "+25% HP", brightgreen, 1, GetPlayerX(Index) * 32, (GetPlayerY(Index) + 1) * 32
                                            HealPlayer Index, GetPlayerMaxVital(Index, HP) * 0.25
                                            MatchData(TempPlayer(Index).MatchIndex).Stars = MatchData(TempPlayer(Index).MatchIndex).Stars + 1
                                            SendMatchData TempPlayer(Index).MatchIndex
                                            If Resource(Resource_index).ItemReward > 0 Then
                                                If Resource(Resource_index).ResourceType <> 4 Then Call HandleScriptedResource(Index, Resource_index)
                                            Else
                                                If IsDaily(Index, DestroyMinor) Then UpdateDaily Index
                                            End If
                                        End If
                                        If Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).Type = 4 Then
                                            Dim DropX As Long, DropY As Long, Drops As Long
                                            Drops = rand(1, 5)
                                            For n = 1 To Drops
                                                DropX = rand(rX - 1, rX + 1)
                                                DropY = rand(rY - 1, rY + 1)
                                                If DropX > 0 And DropX < Map(GetPlayerMap(Index)).MaxX Then
                                                    If DropY > 0 And DropY < Map(GetPlayerMap(Index)).MaxY Then
                                                        If DropX <> rX Or DropY <> rY Then
                                                            If Map(GetPlayerMap(Index)).Tile(DropX, DropY).Type = TileType.TILE_TYPE_WALKABLE Then
                                                                Call SpawnItem(TesouroItem, 1, GetPlayerMap(Index), DropX, DropY)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next n
                                        End If
                                    End If
                                End If
                                ResourceCache(Player(Index).Map).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(Player(Index).Map).ResourceData(Resource_num).cur_health = 0
                                SendResourceCacheToMap GetPlayerMap(Index), Resource_num
                                SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), brightgreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                If Resource(Resource_index).ResourceType = 2 Then GiveInvItem Index, Resource(Resource_index).ItemReward, 1
                                If Resource(Resource_index).ResourceType < 3 Then
                                    PlanetNum = GetPlanetNum(GetPlayerMap(Index))
                                    Dim PlanetType As Long
                                    If PlanetNum > 0 Then PlanetType = Planets(PlanetNum).Type
                                    If TempPlayer(Index).MatchIndex = 0 Or PlanetType <> 0 Then
                                        If PlanetType = 2 Then GiveInvItem Index, Resource(Resource_index).ItemReward, 1
                                        PlanetNum = GetPlanetNum(GetPlayerMap(Index))
                                        If PlanetNum > 0 Then
                                            If Planets(PlanetNum).Type = 2 And TempPlayer(Index).MatchIndex > 0 Then
                                                ResourceCache(Player(Index).Map).ResourceData(Resource_num).cur_health = 0
                                                Dim PlayerItem As Long
                                                PlayerItem = HasItem(Index, Resource(Resource_index).ItemReward)
                                                MatchData(TempPlayer(Index).MatchIndex).Points = PlayerItem
                                                SendMatchData TempPlayer(Index).MatchIndex
                                                If PlayerItem >= Planets(PlanetNum).PointsToConquest Then
                                                    TakeInvItem Index, Resource(Resource_index).ItemReward, PlayerItem
                                                    GiveInvItem Index, MoedaZ, Int(Planets(PlanetNum).Preco * 2)
                                                    SendBossMsg GetPlayerMap(Index), "Parabéns! Você completou a coleta! Este planeta explodirá em 10 segundos.", Yellow, Index
                                                    PlayerMsg Index, "-- Recompensa pela missão --", brightgreen
                                                    PlayerMsg Index, "Você recebeu: " & Int(Planets(PlanetNum).Preco * 2) & "z", brightgreen
                                                    GivePlayerEXP Index, Int(ExperienceBase(Planets(PlanetNum).Level) * 1.5)
                                                    GivePlayerVIPExp Index, Planets(PlanetNum).Level
                                                    If Player(Index).Guild > 0 Then GiveGuildExp Index, Planets(PlanetNum).Level, Planets(PlanetNum).Level
                                                    PlayerMsg Index, "Você recebeu: " & Int(ExperienceBase(Planets(PlanetNum).Level) * 1.5) & " exp", brightgreen
                                                    Planets(PlanetNum).Owner = GetPlayerName(Index)
                                                    Planets(PlanetNum).State = 2
                                                    Planets(PlanetNum).TimeToExplode = GetTickCount + 10000
                                                    MatchData(TempPlayer(Index).MatchIndex).Active = 0
                                                    MatchData(TempPlayer(Index).MatchIndex).WaveTick = GetTickCount + 5000
                                                    For n = 1 To MAX_MAP_NPCS
                                                        DespawnNPC Planets(PlanetNum).Map, n
                                                    Next n
                                                    If TempPlayer(Index).PlanetService = PlanetNum Then
                                                        CompleteService Index
                                                    End If
                                                End If
                                            Else
                                                ResourceCache(Player(Index).Map).ResourceData(Resource_num).cur_health = 1
                                                If PlanetType = 3 And TempPlayer(Index).MatchIndex > 0 Then
                                                    Dim Count As Long
                                                    For n = 1 To ResourceCache(GetPlayerMap(Index)).Resource_Count
                                                        If ResourceCache(GetPlayerMap(Index)).ResourceData(n).ResourceState = 1 And Resource(ResourceCache(GetPlayerMap(Index)).ResourceData(n).ResourceNum).ResourceType <> 6 Then
                                                            Count = Count + 1
                                                        End If
                                                    Next n
                                                    If Count >= ResourceCache(GetPlayerMap(Index)).Resource_Count Then
                                                        GiveInvItem Index, MoedaZ, Int(Planets(PlanetNum).Preco)
                                                        SendBossMsg GetPlayerMap(Index), "Parabéns! Você completou a destruição! Este planeta explodirá em 10 segundos.", Yellow, Index
                                                        PlayerMsg Index, "-- Recompensa pela missão --", brightgreen
                                                        PlayerMsg Index, "Você recebeu: " & Int(Planets(PlanetNum).Preco) & "z", brightgreen
                                                        GivePlayerEXP Index, Int(ExperienceBase(Planets(PlanetNum).Level))
                                                        PlayerMsg Index, "Você recebeu: " & Int(ExperienceBase(Planets(PlanetNum).Level)) & " exp", brightgreen
                                                        GivePlayerVIPExp Index, Planets(PlanetNum).Level
                                                        If Player(Index).Guild > 0 Then GiveGuildExp Index, Planets(PlanetNum).Level, Planets(PlanetNum).Level
                                                        Planets(PlanetNum).Owner = GetPlayerName(Index)
                                                        Planets(PlanetNum).State = 2
                                                        Planets(PlanetNum).TimeToExplode = GetTickCount + 10000
                                                        MatchData(TempPlayer(Index).MatchIndex).Active = 0
                                                        MatchData(TempPlayer(Index).MatchIndex).WaveTick = GetTickCount + 5000
                                                        MatchData(TempPlayer(Index).MatchIndex).Points = (Count / ResourceCache(GetPlayerMap(Index)).Resource_Count) * Planets(PlanetNum).PointsToConquest
                                                        SendMatchData TempPlayer(Index).MatchIndex
                                                        For n = 1 To MAX_MAP_NPCS
                                                            DespawnNPC Planets(PlanetNum).Map, n
                                                        Next n
                                                        If TempPlayer(Index).PlanetService = PlanetNum Then
                                                            CompleteService Index
                                                        End If
                                                    Else
                                                        MatchData(TempPlayer(Index).MatchIndex).Points = (Count / ResourceCache(GetPlayerMap(Index)).Resource_Count) * Planets(PlanetNum).PointsToConquest
                                                        SendMatchData TempPlayer(Index).MatchIndex
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    If Resource(Resource_index).ResourceType = 3 Then
                                        If GetPlayerX(Index) = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).X Then
                                            If GetPlayerY(Index) = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).Y Then
                                                If Fisgada(Index) Then
                                                    Dim Peixe As Integer, MapItemNum As Byte
                                                    Peixe = Resource(Resource_index).ItemReward
                                                        
                                                    MapItemNum = FindOpenMapItemSlot(GetPlayerMap(Index))
                                                    MapItem(GetPlayerMap(Index), MapItemNum).Num = Peixe
                                                    MapItem(GetPlayerMap(Index), MapItemNum).Value = 1
                                                    MapItem(GetPlayerMap(Index), MapItemNum).X = GetPlayerX(Index)
                                                    MapItem(GetPlayerMap(Index), MapItemNum).Y = GetPlayerY(Index)
                                                    Call SpawnItemSlot(MapItemNum, MapItem(GetPlayerMap(Index), MapItemNum).Num, 1, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), Trim$(GetPlayerName(Index)), MapItem(GetPlayerMap(Index), MapItemNum).canDespawn)
                                                        
                                                    SendPlaySound Index, "fish.mp3"
                                                    SendActionMsg GetPlayerMap(Index), Trim$(Item(Peixe).Name) & "!", brightgreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                                    Call SendFishingTime(Index)
                                                Else
                                                    Call SendFishingTime(Index)
                                                End If
                                            End If
                                        End If
                                    End If
                                    SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY, GetPlayerDir(Index)
                                End If
                                If Resource(Resource_index).ResourceType = 4 Then
                                    ExtractorGet Index, Resource_index, Resource_num
                                End If
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage
                                SendActionMsg GetPlayerMap(Index), "-" & Damage, brightred, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY, GetPlayerDir(Index)
                            End If
                            ' send the sound
                            SendMapSound Index, rX, rY, SoundEntity.seResource, Resource_index
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(Index), actionf("Errou!"), brightred, 1, (rX * 32), (rY * 32)
                        End If
                    Else
                        SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), brightred, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                    End If
        End If
    End If
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(Index).Item(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Bank(Index).Item(BankSlot).Num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(Index).Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank(Index).Item(BankSlot).Value = ItemValue
End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal invSlot As Long, ByVal Amount As Long)
Dim BankSlot

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invSlot) Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(Index, GetPlayerInvItemNum(Index, invSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, invSlot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, invSlot)).Stackable > 0 Then
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, invSlot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), Amount)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, invSlot))
                Call SetPlayerBankItemValue(Index, BankSlot, Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(Index, BankSlot) = GetPlayerInvItemNum(Index, invSlot) Then
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) + 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(Index, BankSlot, GetPlayerInvItemNum(Index, invSlot))
                Call SetPlayerBankItemValue(Index, BankSlot, 1)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), 0)
            End If
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index

End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal BankSlot As Long, ByVal Amount As Long)
Dim invSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If Amount < 0 Or Amount > GetPlayerBankItemValue(Index, BankSlot) Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(Index, GetPlayerBankItemNum(Index, BankSlot))
        
    If invSlot > 0 Then
        If Item(GetPlayerBankItemNum(Index, BankSlot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(Index, BankSlot)).Stackable > 0 Then
            Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), Amount)
            Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - Amount)
            If GetPlayerBankItemValue(Index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(Index, BankSlot) > 1 Then
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemValue(Index, BankSlot, GetPlayerBankItemValue(Index, BankSlot) - 1)
            Else
                Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankSlot), 0)
                Call SetPlayerBankItemNum(Index, BankSlot, 0)
                Call SetPlayerBankItemValue(Index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index

End Sub

Public Sub KillPlayer(ByVal Index As Long)
Dim Exp As Long

    ' Calculate exp to give attacker
    'Exp = GetPlayerExp(Index) \ 3

    ' Make sure we dont get less then 0
    'If Exp < 0 Then Exp = 0
    'If Exp > 0 Then
    '    Call SetPlayerExp(Index, GetPlayerExp(Index) - Exp)
    '    SendEXP Index
    '    Call PlayerMsg(Index, printf("Você perdeu %d exp.", Val(Exp)), brightred)
    'End If
    
    Dim i As Long
    If GetTotalMapPlayers(GetPlayerMap(Index)) = 0 Then
        PlayersOnMap(GetPlayerMap(Index)) = NO

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS

            If MapNpc(GetPlayerMap(Index)).Npc(i).Num > 0 Then
                MapNpc(GetPlayerMap(Index)).Npc(i).Vital(Vitals.HP) = GetNpcMaxVital(GetPlayerMap(Index), i, Vitals.HP)
            End If

        Next
    End If
    
    If GetPlanetNum(GetPlayerMap(Index)) > 0 Then
        If TempPlayer(Index).MatchIndex > 0 Then
            GiveConquistaReward Index, UBound(MatchData(TempPlayer(Index).MatchIndex).Indexes), MatchData(TempPlayer(Index).MatchIndex).HighLevel, (MatchData(TempPlayer(Index).MatchIndex).Points / Planets(MatchData(TempPlayer(Index).MatchIndex).Planet).PointsToConquest) * 100
        End If
    End If
    Call CheckItemDrop(Index)
    
    If Player(Index).IsDead = 0 And TempPlayer(Index).inDevSuite = 0 Then Player(Index).IsDead = 1
    Call SendPlayerData(Index)
    'Call OnDeath(Index)
End Sub

Public Sub UseItem(ByVal Index As Long, ByVal InvNum As Long)
Dim n As Long, i As Long, tempItem As Long, X As Long, Y As Long, ItemNum As Long

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If TempPlayer(Index).InTrade > 0 Then Exit Sub
    If (GetPlayerMap(Index) = ViagemMap And UZ) Or Player(Index).GravityHours > 0 Then Exit Sub

    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(Index, InvNum)).data2
        ItemNum = GetPlayerInvItemNum(Index, InvNum)
        
        ' Find out what kind of item it is
        Select Case Item(ItemNum).Type
            Case ITEM_TYPE_ARMOR
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Armor)
                End If

                SetPlayerEquipment Index, ItemNum, Armor
                PlayerMsg Index, printf("Você equipou %s", CheckGrammar(Item(ItemNum).Name)), brightgreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                Call SendStats(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_WEAPON
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, Weapon) > 0 Then
                    tempItem = GetPlayerEquipment(Index, Weapon)
                End If

                SetPlayerEquipment Index, ItemNum, Weapon
                PlayerMsg Index, printf("Você equipou %s", CheckGrammar(Item(ItemNum).Name)), brightgreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                Call SendStats(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_HELMET
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If

                If GetPlayerEquipment(Index, helmet) > 0 Then
                    tempItem = GetPlayerEquipment(Index, helmet)
                End If

                SetPlayerEquipment Index, ItemNum, helmet
                PlayerMsg Index, printf("Você equipou %s", CheckGrammar(Item(ItemNum).Name)), brightgreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                Call SendStats(Index)
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_SHIELD
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If
                
                If GetPlayerEquipment(Index, shield) > 0 Then
                    tempItem = GetPlayerEquipment(Index, shield)
                End If

                SetPlayerEquipment Index, ItemNum, shield
                PlayerMsg Index, printf("Você equipou %s", CheckGrammar(Item(ItemNum).Name)), brightgreen
                TakeInvItem Index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(Index, Vitals.HP)
                Call SendVital(Index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                Call SendStats(Index)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If
                
                ' add hp
                If Item(ItemNum).AddHP > 0 Then
                    Player(Index).Vital(Vitals.HP) = Player(Index).Vital(Vitals.HP) + Item(ItemNum).AddHP
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddHP, brightgreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, HP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                ' add mp
                If Item(ItemNum).AddMP > 0 Then
                    Player(Index).Vital(Vitals.MP) = Player(Index).Vital(Vitals.MP) + Item(ItemNum).AddMP
                    SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddMP, brightblue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SendVital Index, MP
                    ' send vitals to party if in one
                    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
                End If
                ' add exp
                If Item(ItemNum).AddEXP > 0 Then
                    'SetPlayerExp Index, GetPlayerExp(Index) + Item(ItemNum).AddEXP
                    'CheckPlayerLevelUp Index
                    'SendActionMsg GetPlayerMap(Index), "+" & Item(ItemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    'SendEXP Index
                    'Habilidade especial
                    If Item(ItemNum).AddEXP = 1 Then
                        PlayerWarp Index, 52, 10, 10
                        Exit Sub
                    End If
                End If
                Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index)
                Call TakeInvItem(Index, Player(Index).Inv(InvNum).Num, 1)
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_SPELL
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(Index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = Item(ItemNum).ClassReq Then
                        PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(Index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg Index, printf("Você não tem os requerimentos para esse item."), brightred
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(ItemNum).data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq

                        If i <= GetPlayerLevel(Index) Then
                            i = FindOpenSpellSlot(Index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(Index, n) Then
                                    If HasItem(Index, ItemNum) >= Item(ItemNum).data2 Then
                                        Call SetPlayerSpell(Index, i, n)
                                        Call SendAnimation(GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index)
                                        Call TakeInvItem(Index, ItemNum, Item(ItemNum).data2)
                                        Call PlayerMsg(Index, printf("Você aprendeu %s.", Trim$(Spell(n).Name)), brightgreen)
                                    Else
                                        Call PlayerMsg(Index, "Você precisa de " & Item(ItemNum).data2 & " cópias deste item para adquirir a magia", brightred)
                                    End If
                                Else
                                    Call PlayerMsg(Index, printf("Você ja conhece esta técnica."), brightred)
                                End If

                            Else
                                Call PlayerMsg(Index, printf("Você não pode aprender mais técnicas!."), brightred)
                            End If

                        Else
                            Call PlayerMsg(Index, printf("Você não tem os requerimentos para esse item."), brightred)
                        End If

                    Else
                        Call PlayerMsg(Index, printf("Você não tem os requerimentos para esse item."), brightred)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_ESOTERICA
                If Player(Index).EsoBonus > 0 Then
                    PlayerMsg Index, printf("Você já está sobre o efeito de uma esotérica!"), brightred
                    Exit Sub
                End If
                
                Player(Index).EsoBonus = Item(ItemNum).EsotericaBonus
                Player(Index).EsoTime = Item(ItemNum).EsotericaTime
                Player(Index).EsoNum = ItemNum
                
                Call TakeInvItem(Index, ItemNum, 1)
                
                'Notice
                PlayerMsg Index, printf("Parabéns! você recebeu o bonus de %d por cento de EXP por %d minutos!", Player(Index).EsoBonus & "," & Player(Index).EsoTime), Yellow
                
                ' send the sound and effects
                SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum
                SendAnimation GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index
                SendPlayerData Index
                SavePlayer Index
            Case ITEM_TYPE_TITULO
                If Player(Index).Titulo <> ItemNum Then
                    Player(Index).Titulo = ItemNum
                    PlayerMsg Index, printf("Seu título agora é: %s", Trim$(Item(ItemNum).Name)), Yellow
                Else
                    Player(Index).Titulo = 0
                    PlayerMsg Index, printf("Você removeu seu título"), Yellow
                End If
                SendPlayerData Index
                SavePlayer Index
            Case ITEM_TYPE_EXTRATOR
                PutExtractor Index, InvNum
            Case ITEM_TYPE_NAVE
                If Player(Index).GravityHours = 0 Then
                    If Not getProvação(GetPlayerMap(Index)) > 0 Then
                        If TempPlayer(Index).stopRegen = False Then
                            If Not GetPlayerMap(Index) = ViagemMap And Not GetPlayerMap(Index) = MapRespawn Then
                                TempPlayer(Index).Speed = Item(ItemNum).data2
                                TempPlayer(Index).Nave = ItemNum
                                Viajar Index, ViagemMap
                            End If
                        Else
                            PlayerMsg Index, "Você não pode usar a nave em combate", brightred
                        End If
                    Else
                        PlayerMsg Index, "Você não pode usar a nave em uma provação", brightred
                    End If
                End If
            Case ITEM_TYPE_COMBUSTIVEL
                EnhanceExtractor Index, InvNum
            Case ITEM_TYPE_VIP
                If Player(Index).VIP = 0 Then
                    Player(Index).VIP = 1
                    Player(Index).VIPData = Now
                    Player(Index).VIPDias = Item(ItemNum).data1
                    Player(Index).VIPExp = 0
                Else
                    Player(Index).VIPDias = Player(Index).VIPDias + Item(ItemNum).data1
                End If
                PlayerMsg Index, "Você ganhou " & Item(ItemNum).data1 & " dias vip", Yellow
                TakeInvItem Index, ItemNum, 1
                SavePlayer Index
                SendPlayerData Index
            Case ITEM_TYPE_CAPTURE
                CapturePlanet Index, ItemNum
            Case ITEM_TYPE_BAU
                Dim Value As Long
                Dim Total As Long
                If GetPlayerMap(Index) = RESPAWN_MAP Then Exit Sub
                If Item(ItemNum).data1 = 0 Then
                SendAnimation GetPlayerMap(Index), Item(ItemNum).Animation, 0, 0, GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index
                For i = 1 To 40
                    If Item(ItemNum).LuckySlot(i).ItemNum > 0 Then
                        Total = Total + Item(ItemNum).LuckySlot(i).Chance
                    End If
                Next i
                Value = rand(1, Total)
                Total = 0
                Dim Quant As Long
                For i = 1 To 40
                    If Value >= Total And Value <= Total + Item(ItemNum).LuckySlot(i).Chance Then
                        Quant = Item(ItemNum).LuckySlot(i).Quant
                        If Quant > 1 Then Quant = rand(Quant * 0.8, Quant)
GenerateCordinates:
                        X = rand(GetPlayerX(Index) - 1, GetPlayerX(Index) + 1)
                        Y = rand(GetPlayerY(Index) - 1, GetPlayerY(Index) + 1)
                        If X < 0 Or X > Map(GetPlayerMap(Index)).MaxX Or Y < 0 Or Y > Map(GetPlayerMap(Index)).MaxY Then GoTo GenerateCordinates
                        If Map(GetPlayerMap(Index)).Tile(X, Y).Type <> TileType.TILE_TYPE_WALKABLE Then GoTo GenerateCordinates
                        Call SpawnItem(Item(ItemNum).LuckySlot(i).ItemNum, Quant, GetPlayerMap(Index), X, Y)
                        TakeInvItem Index, ItemNum, 1
                        Exit Sub
                    End If
                    Total = Total + Item(ItemNum).LuckySlot(i).Chance
                Next i
                Else
                    Total = 0
                    For i = 1 To 40
                        If Item(ItemNum).LuckySlot(i).ItemNum > 0 Then
                            Total = Total + 1
                        End If
                    Next i
                    Dim ValidSpaces As Long
                    ValidSpaces = 0
                    For i = 1 To MAX_INV
                        If GetPlayerInvItemNum(Index, i) = 0 Then ValidSpaces = ValidSpaces + 1
                    Next i
                    If ValidSpaces >= Total Then
                        For i = 1 To 40
                            If Item(ItemNum).LuckySlot(i).ItemNum > 0 Then
                                GiveInvItem Index, Item(ItemNum).LuckySlot(i).ItemNum, Item(ItemNum).LuckySlot(i).Quant
                                PlayerMsg Index, "Você recebeu " & Item(ItemNum).LuckySlot(i).Quant & " " & Trim$(Item(Item(ItemNum).LuckySlot(i).ItemNum).Name) & "!", brightgreen
                            End If
                        Next i
                        TakeInvItem Index, ItemNum, 1
                    Else
                        PlayerMsg Index, "Você não tem espaço suficiente no inventário para os itens deste pacote! (Espaços necessários: " & Total & ")", brightred
                        Exit Sub
                    End If
                End If
            Case ITEM_TYPE_PLANETCHANGE
                Dim PlanetNum As Long
                Dim Changed As Boolean
                Dim MapNum As Long
                MapNum = GetPlayerMap(Index)
                PlanetNum = PlayerMapIndex(MapNum)
                If PlanetNum = 0 Then Exit Sub
                If Trim$(LCase(PlayerPlanet(PlanetNum).PlanetData.Owner)) <> Trim$(LCase(GetPlayerName(Index))) Then Exit Sub
                If Item(ItemNum).data1 = 1 Then
                    If PlayerPlanet(PlanetNum).PlanetData.Atmosfera - 5 > 0 Then
                        PlayerPlanet(PlanetNum).PlanetData.Atmosfera = PlayerPlanet(PlanetNum).PlanetData.Atmosfera - 5
                    Else
                        PlayerPlanet(PlanetNum).PlanetData.Atmosfera = 0
                    End If
                    Map(MapNum).FogOpacity = 255 - PlayerPlanet(PlanetNum).PlanetData.Atmosfera
                    PlayerMsg Index, "A poluição da atmosfera do seu planeta reduziu em 5%", brightgreen
                    SendSpecialEffect Index, EFFECT_TYPE_FOG, Map(MapNum).Fog, Map(MapNum).FogSpeed, Map(MapNum).FogOpacity
                    TakeInvItem Index, ItemNum, 1
                    Changed = True
                End If
                If Item(ItemNum).data1 = 2 Then
                    If Item(ItemNum).data2 = 0 Then
                        Map(MapNum).Red = Map(MapNum).Red + Item(ItemNum).data3
                        If Map(MapNum).Red > 255 Then Map(MapNum).Red = 255
                        If Map(MapNum).Red < 0 Then Map(MapNum).Red = 0
                    End If
                    If Item(ItemNum).data2 = 1 Then
                        Map(MapNum).Green = Map(MapNum).Green + Item(ItemNum).data3
                        If Map(MapNum).Green > 255 Then Map(MapNum).Green = 255
                        If Map(MapNum).Green < 0 Then Map(MapNum).Green = 0
                    End If
                    If Item(ItemNum).data2 = 2 Then
                        Map(MapNum).Blue = Map(MapNum).Blue + Item(ItemNum).data3
                        If Map(MapNum).Blue > 255 Then Map(MapNum).Blue = 255
                        If Map(MapNum).Blue < 0 Then Map(MapNum).Blue = 0
                    End If
                    If Item(ItemNum).data2 = 3 Then
                        Map(MapNum).Alpha = Map(MapNum).Alpha + Item(ItemNum).data3
                        If Map(MapNum).Alpha > 160 Then Map(MapNum).Alpha = 160
                        If Map(MapNum).Alpha < 0 Then Map(MapNum).Alpha = 0
                    End If
                    TakeInvItem Index, ItemNum, 1
                    SendSpecialEffect Index, EFFECT_TYPE_TINT, Map(MapNum).Red, Map(MapNum).Green, Map(MapNum).Blue, Map(MapNum).Alpha
                    PlayerMsg Index, "A cor de sua atmosfera sofreu modificação!", brightgreen
                    Changed = True
                End If
                If Item(ItemNum).data1 = 3 Then
                    If Item(ItemNum).data2 = 0 Then
                        If PlayerPlanet(PlanetNum).PlanetData.ColorR + Item(ItemNum).data3 > 255 Then
                            PlayerPlanet(PlanetNum).PlanetData.ColorR = 255
                        Else
                            PlayerPlanet(PlanetNum).PlanetData.ColorR = PlayerPlanet(PlanetNum).PlanetData.ColorR + Item(ItemNum).data3
                        End If
                        If PlayerPlanet(PlanetNum).PlanetData.ColorR < 80 Then PlayerPlanet(PlanetNum).PlanetData.ColorR = 80
                    End If
                    If Item(ItemNum).data2 = 1 Then
                        If PlayerPlanet(PlanetNum).PlanetData.ColorG + Item(ItemNum).data3 > 255 Then
                            PlayerPlanet(PlanetNum).PlanetData.ColorG = 255
                        Else
                            PlayerPlanet(PlanetNum).PlanetData.ColorG = PlayerPlanet(PlanetNum).PlanetData.ColorG + Item(ItemNum).data3
                        End If
                        If PlayerPlanet(PlanetNum).PlanetData.ColorG < 80 Then PlayerPlanet(PlanetNum).PlanetData.ColorG = 80
                    End If
                    If Item(ItemNum).data2 = 2 Then
                        If PlayerPlanet(PlanetNum).PlanetData.ColorB + Item(ItemNum).data3 > 255 Then
                            PlayerPlanet(PlanetNum).PlanetData.ColorB = 255
                        Else
                            PlayerPlanet(PlanetNum).PlanetData.ColorB = PlayerPlanet(PlanetNum).PlanetData.ColorB + Item(ItemNum).data3
                        End If
                        If PlayerPlanet(PlanetNum).PlanetData.ColorB < 80 Then PlayerPlanet(PlanetNum).PlanetData.ColorB = 80
                    End If
                    TakeInvItem Index, ItemNum, 1
                    PlayerMsg Index, "A cor de seu planeta sofreu modificação! Volte para o espaço para ver!", brightgreen
                    Changed = True
                End If
                If Item(ItemNum).data1 = 5 Then
                    Map(MapNum).Ambiente = Item(ItemNum).data2
                    TakeInvItem Index, ItemNum, 1
                    PlayerMsg Index, "Ecossistema ativado! Ambiente de criaturas modificado!", brightgreen
                    Changed = True
                End If
                If Item(ItemNum).data1 = 6 Then
                    Map(MapNum).Weather = Item(ItemNum).data2
                    Map(MapNum).WeatherIntensity = 100
                    SendSpecialEffect Index, EFFECT_TYPE_WEATHER, Item(ItemNum).data2, 100
                    TakeInvItem Index, ItemNum, 1
                    PlayerMsg Index, "Ecossistema ativado! Ambiente de criaturas modificado!", brightgreen
                    Changed = True
                End If
                If Item(ItemNum).data1 = 7 Then
                    Dim NpcCount As Long
                    For i = 1 To MAX_MAP_NPCS
                        If Map(MapNum).Npc(i) > 0 Then
                            If Trim$(LCase(Npc(Map(MapNum).Npc(i)).Name)) = Trim$(LCase(Npc(Item(ItemNum).data2).Name)) Then NpcCount = NpcCount + 1
                        End If
                    Next i
                    For i = 1 To PlayerPlanet(PlanetNum).TotalSaibamans
                        If PlayerPlanet(PlanetNum).Saibaman(i).Working = 1 Then
                            If PlayerPlanet(PlanetNum).Saibaman(i).TaskType = 0 Then
                                If Trim$(LCase(Npc(PlayerPlanet(PlanetNum).Saibaman(i).TaskResult).Name)) = Trim$(LCase(Npc(Item(ItemNum).data2).Name)) Then NpcCount = NpcCount + 1
                            End If
                        End If
                    Next i
                    If NucleoLevel(MapNum) = 0 Then
                        PlayerMsg Index, "Você precisa de um centro no seu planeta para fazer construções!", brightred
                        Exit Sub
                    End If
                    If NpcCount < 1 + Int(NucleoLevel(MapNum) / Item(ItemNum).data3) Then 'Nucleo Level
                        'If AddSaibaman(Index, PlanetNum, MapNum, 0, Item(ItemNum).Data2) Then
                        '    PlayerMsg Index, "Tarefa iniciada com sucesso!", brightgreen
                        '    TakeInvItem Index, ItemNum, 1
                        '    SavePlayerPlanet PlanetNum
                        'End If
                        If StartConstructNPC(Index, MapNum, PlanetNum, GetPlayerX(Index), GetPlayerY(Index), Item(ItemNum).data2) Then
                            TakeInvItem Index, ItemNum, 1
                        End If
                    Else
                        PlayerMsg Index, "Você atingiu o limite máximo de construções deste tipo! Evolua seu centro para liberar mais!", brightred
                    End If
                End If
                If Item(ItemNum).data1 = 8 Then
                    If Resource(Item(ItemNum).data2).NucleoLevel > 0 And NucleoLevel(MapNum) > 0 Then
                        PlayerMsg Index, "Seu planeta já tem um centro!", brightred
                        Exit Sub
                    End If
                    If NucleoLevel(MapNum) = 0 And Not Resource(Item(ItemNum).data2).NucleoLevel > 0 Then
                        PlayerMsg Index, "Você precisa de um centro no seu planeta para fazer construções!", brightred
                        Exit Sub
                    End If
                    If ResourceCount(MapNum, Item(ItemNum).data2) < 1 + Int(NucleoLevel(MapNum) / Item(ItemNum).data3) Then 'Nucleo Level
                        'If AddSaibaman(Index, PlanetNum, MapNum, 1, Item(ItemNum).Data2) Then
                        '    PlayerMsg Index, "Tarefa iniciada com sucesso!", brightgreen
                        '    TakeInvItem Index, ItemNum, 1
                        '    SavePlayerPlanet PlanetNum
                        'End If
                        If StartConstructResource(Index, Item(ItemNum).data2, MapNum, PlanetNum, GetPlayerX(Index), GetPlayerY(Index)) Then
                            TakeInvItem Index, ItemNum, 1
                        End If
                    Else
                        PlayerMsg Index, "Você atingiu o limite máximo de construções deste tipo! Evolua seu centro para liberar mais!", brightred
                    End If
                End If
                If Item(ItemNum).data1 = 10 Then
                    If PlayerPlanet(PlanetNum).TotalSaibamans < 5 Then
                        If PlayerPlanet(PlanetNum).TotalSaibamans + 1 = Item(ItemNum).data2 Then
                            PlayerPlanet(PlanetNum).TotalSaibamans = PlayerPlanet(PlanetNum).TotalSaibamans + 1
                            SavePlayerPlanet PlanetNum
                            TakeInvItem Index, ItemNum, 1
                            PlayerMsg Index, "Você ganhou um novo saibaman operário! Numéro atual (" & PlayerPlanet(PlanetNum).TotalSaibamans & "/5)", brightgreen
                        Else
                            PlayerMsg Index, "Este é o saibaman número " & Item(ItemNum).data2 & " e você está com " & PlayerPlanet(PlanetNum).TotalSaibamans & " atualmente!", brightred
                        End If
                    Else
                        PlayerMsg Index, "Você só pode ter até 5 saibamans em seu planeta!", brightred
                    End If
                End If
                If Changed Then
                    PlayerPlanet(PlanetNum).PlanetMap = Map(MapNum)
                    MapCache_Create MapNum
                    SendMap Index, MapNum
                    SavePlayerPlanet PlanetNum
                    SendPlayerPlanetToAll PlanetNum
                End If
        End Select
    End If
End Sub

Sub CheckEvent(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim Event_num As Long
    Dim Event_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    
    If GetPlayerMap(Index) > 0 Then
        If X > 0 And X < Map(GetPlayerMap(Index)).MaxX Then
            If Y > 0 And Y < Map(GetPlayerMap(Index)).MaxY Then
                If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_EVENT Then
                    Event_index = Map(GetPlayerMap(Index)).Tile(X, Y).data1
                End If
            End If
        End If
    End If
    
    If Event_index > 0 Then
        If Events(Event_index).Trigger > 0 Then
            InitEvent Index, Event_index
        End If
    End If
End Sub

Sub TransPlayer(ByVal Index As Long, ByVal SpellNum As Long)
    If TempPlayer(Index).Trans > 0 Then
        If Spell(TempPlayer(Index).Trans).HairChange = 5 Then 'Ozaru
            Call SetPlayerSprite(Index, GetPlayerNormalSprite(Index))
        End If
        If SpellNum = 0 Then Call SendAnimation(GetPlayerMap(Index), Spell(TempPlayer(Index).Trans).CastAnim + 1, GetPlayerX(Index), GetPlayerY(Index), GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index)
    End If
    
    TempPlayer(Index).Trans = SpellNum
    
    If SpellNum > 0 Then
        If Spell(SpellNum).HairChange < 5 Then 'Ozaru
            TempPlayer(Index).HairChange = Spell(SpellNum).HairChange
        Else
            Call SetPlayerSprite(Index, 40)
            TempPlayer(Index).HairChange = 5
        End If
    Else
        TempPlayer(Index).HairChange = 0
    End If
    
    SendPlayerData Index
End Sub

Function GetPlayerPDL(ByVal Index As Long) As Long
    
    If TempPlayer(Index).Trans = 0 Then
        GetPlayerPDL = Player(Index).PDL
    Else
        GetPlayerPDL = (Player(Index).PDL / 100) * (Spell(TempPlayer(Index).Trans).PDLBonus + 100)
    End If
    
End Function

Sub CheckVip(ByVal Index As Long)
    If Player(Index).VIP > 0 Then
        If DateDiff("d", Player(Index).VIPData, Date) >= Player(Index).VIPDias Then
            Player(Index).VIP = 0
            PlayerMsg Index, printf("Seus dias de VIP se encerraram, renove seu vip e não fique sem seus grandiosos bonus!"), brightred
            Exit Sub
        End If
        PlayerMsg Index, printf("Você ainda tem %d dias vip!", Val((Player(Index).VIPDias - DateDiff("d", Player(Index).VIPData, Date)))), brightgreen
    End If
End Sub

Function GetPlayerAccuracy(ByVal Index As Long) As Long
GetPlayerAccuracy = GetPlayerLevel(Index) + (GetPlayerStat(Index, Willpower))
End Function

Function GetPlayerEvasion(ByVal Index As Long) As Long
GetPlayerEvasion = GetPlayerStat(Index, agility)
End Function

Function IsHouseValid(ByVal Index As Long) As Boolean
    IsHouseValid = False
    
    Dim i As Long
    For i = 1 To TotalHouses
        If LCase(House(i).Proprietario) = LCase(GetPlayerName(Index)) Then
            If DateDiff("d", House(i).DataDeInicio, Date) < House(i).Dias Then
                IsHouseValid = True
                Exit Function
            End If
        End If
    Next i
End Function

Function isAFK(ByVal Index As Long) As Boolean
    If TempPlayer(Index).LastMove + AFKTime < GetTickCount And TempPlayer(Index).LastMove > 0 Then
        isAFK = True
    Else
        isAFK = False
    End If
End Function

Sub Report(ByVal Index As Long, ByVal Msg As String)
    Dim PlayerIndex As Long
    If TempPlayer(Index).Target = TARGET_TYPE_PLAYER Then
        PlayerIndex = TempPlayer(Index).Target
        If PlayerIndex > 0 Then
            Dim motivo As String
            motivo = Mid(LCase(Msg), 10)
            If Trim(motivo) = "" Then motivo = "Sem motivo"
            Call TextAdd("[REPORT] " & Trim$(Player(PlayerIndex).Name) & " reportado no mapa " & GetPlayerMap(Index) & " por " & GetPlayerName(Index) & " motivo '" & motivo & "'", ChatSystem)
            Call addReportLog(Trim$(Player(PlayerIndex).Name) & " reportado no mapa " & GetPlayerMap(Index) & " por " & GetPlayerName(Index) & " motivo '" & motivo & "'")
            Dim i As Long
            For i = 1 To Player_HighIndex
                If GetPlayerAccess(i) > 0 Then
                    Call PlayerMsg(i, printf("%s reportado no mapa %d por %s motivo '%s'", Trim$(Player(PlayerIndex).Name) & "," & GetPlayerMap(Index) & "," & GetPlayerName(Index) & "," & motivo), Yellow)
                End If
            Next i
        End If
    End If
End Sub

Function GetPlayerStatNextLevel(Index, ByVal stat As Stats)
    On Error Resume Next
    Dim StatLevel As Long
    StatLevel = GetPlayerRawStat(Index, stat)
    GetPlayerStatNextLevel = StatExperience(StatLevel)
End Function

Function GetPlayerStatLastLevel(Index, ByVal stat As Stats)
    On Error Resume Next
    Dim StatLevel As Long
    If GetPlayerStat(Index, stat, True) > 1 Then
        StatLevel = GetPlayerRawStat(Index, stat)
        GetPlayerStatLastLevel = StatExperience((StatLevel - 1))
    Else
        GetPlayerStatLastLevel = 0
    End If
End Function

Sub CheckPlayerStatLevelUp(ByVal Index As Long, ByVal stat As Stats)
    Dim statPoints As Long, NextLevel As Long
    statPoints = GetPlayerStatPoints(Index, stat)
    NextLevel = GetPlayerStatNextLevel(Index, stat)
    Dim LastLevel As Long
    LastLevel = GetPlayerStatLastLevel(Index, stat)
    If statPoints > NextLevel Then
        Do While statPoints > NextLevel
            statPoints = GetPlayerStatPoints(Index, stat)
            Player(Index).stat(stat) = Player(Index).stat(stat) + 1
            NextLevel = GetPlayerStatNextLevel(Index, stat)
        Loop
        Call SendPlayerData(Index)
        Call SavePlayer(Index)
        SendAnimation GetPlayerMap(Index), StatLevelUpAnim, GetPlayerX(Index), GetPlayerY(Index), GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index
    End If
End Sub

Sub UpgradeSpell(ByVal Index As Long, ByVal SpellNum As Long)
    
    'Dim SpellNum As Long
    'SpellNum = Player(Index).Spell(spellslot)
    
    If HasSpell(Index, SpellNum) Then
        'If Spell(SpellNum).Upgrade = 0 Then
        '    PlayerMsg Index, "Esta habilidade não pode evoluir mais!", brightred
        '    Exit Sub
        'End If
    Else
        Dim i As Long
        i = FindOpenSpellSlot(Index)
        If i = 0 Then
            PlayerMsg Index, "Você não tem espaço para novas habilidades!", brightred
            Exit Sub
        End If
    End If
    
            If Spell(SpellNum).Requisite > 0 Then
                If Item(Spell(SpellNum).Requisite).Type <> ItemType.ITEM_TYPE_TITULO Then
                    If Not HasItem(Index, Spell(SpellNum).Requisite) > 0 Then
                        Call PlayerMsg(Index, printf("Você não tem o requisito necessário para aprimorar esta habilidade"), brightred)
                        Exit Sub
                    End If
                Else
                    If Player(Index).Titulo > 0 Then
                        If Item(Player(Index).Titulo).LevelReq < Item(Spell(SpellNum).Requisite).LevelReq Then
                            Call PlayerMsg(Index, printf("Você não tem o cargo necessário para aprimorar esta habilidade"), brightred)
                            Exit Sub
                        End If
                    Else
                        Call PlayerMsg(Index, printf("Você não tem o cargo necessário ativo!"), brightred)
                        Exit Sub
                    End If
                End If
            End If
            
            If Spell(SpellNum).Item > 0 Then
                If Not HasItem(Index, Spell(SpellNum).Item) >= Spell(SpellNum).Price Then
                    Call PlayerMsg(Index, printf("Você não tem o valor para pagar por este aprimoramento"), brightred)
                    Exit Sub
                End If
                TakeInvItem Index, Spell(SpellNum).Item, Spell(SpellNum).Price
            End If
            
    If HasSpell(Index, SpellNum) Then
        'Evoluir
            'Change spells
            Dim spellslot As Long
            Dim lastspell As Long
            For i = 1 To MAX_SPELLS
                If Trim$(Spell(i).Name) = vbNullString Then
                    Exit For
                Else
                    If Spell(i).Upgrade = SpellNum Then
                        lastspell = i
                        Exit For
                    End If
                End If
            Next i
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(Index).Spell(i) = lastspell Then
                    spellslot = i
                    Exit For
                End If
            Next i
            If spellslot = 0 Then
                PlayerMsg Index, "Falha ao evoluir skill", brightred
                Exit Sub
            End If
            Player(Index).Spell(spellslot) = SpellNum
            
            For i = 1 To MAX_HOTBAR
                If Player(Index).Hotbar(i).sType = 2 Then
                    If Player(Index).Hotbar(i).slot = spellslot Then
                        Player(Index).Hotbar(i).slot = 0
                        Player(Index).Hotbar(i).sType = 0
                        SendHotbar Index
                    End If
                End If
            Next i
            
            PlayerMsg Index, printf("Sua habilidade evoluiu para %s!", Trim$(Spell(Player(Index).Spell(spellslot)).Name)), Yellow
            SendPlayerSpells Index

    Else
        'Aprender
        Player(Index).Spell(i) = SpellNum
        PlayerMsg Index, "Você aprendeu a habilidade: " & Trim$(Spell(SpellNum).Name) & "!", Yellow
        SendPlayerSpells Index
    End If
    
End Sub

Sub CheckProvacoesMap(ByVal Index As Long, Optional Logoff As Boolean = True)
    Dim i As Long
    For i = 1 To ProvaçãoCount
        If GetPlayerMap(Index) = Provação(i).Map Then
            If Logoff Then Call PlayerWarp(Index, RESPAWN_MAP, RESPAWN_X, RESPAWN_Y)
            Provação(i).ActualTick = 0
            If Not Logoff Then OnDeath Index
            Exit For
        End If
    Next i
End Sub

Function GetPlayerCargo(ByVal Index As Long) As Long
    Dim Cargo As Long
    
    If Player(Index).Titulo > 0 Then
        Cargo = Player(Index).Titulo
    Else
        Dim i As Long
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) > 0 Then
                If Item(GetPlayerInvItemNum(Index, i)).Type = ITEM_TYPE_TITULO Then
                    Cargo = GetPlayerInvItemNum(Index, i)
                    Exit For
                End If
            End If
        Next i
    End If

    GetPlayerCargo = Cargo
End Function

Sub InitProvação(ByVal Index As Long, ByVal ProvNum As Byte)
    
    'Requisitos
    If GetPlayerLevel(Index) < Provação(ProvNum).MinLevel Then
        Call PlayerMsg(Index, printf("O level mínimo para participar desta provação é: %d", Val(Provação(ProvNum).MinLevel)), brightred)
        Exit Sub
    End If
    
    If HasItem(Index, MoedaZ) < Provação(ProvNum).Cost Then
        Call PlayerMsg(Index, printf("Você precisa ter: %d Moedas Z para participar desta provação", Val(Provação(ProvNum).Cost)), brightred)
        Exit Sub
    End If
    
    ' Area em uso
    If Provação(ProvNum).ActualTick + 600000 > GetTickCount Then
        Call PlayerMsg(Index, printf("Atualmente a provação que você pretende entrar se encontra em período de intervalo, devido á entrada recente de algum jogador. Por favor aguarde %d minutos até a área ser liberada!", Val(Int((GetTickCount - Provação(ProvNum).ActualTick) / 60000))), brightred)
        Exit Sub
    End If
    
    If PlayersOnMap(Provação(ProvNum).Map) > 0 Then
        Call PlayerMsg(Index, printf("Atualmente existe um jogador presente no mapa das provações que você pretende entrar, aguarde até ele acabar a provação."), brightred)
        Exit Sub
    End If
    
    Dim ServicesReq As Long
    Dim Cargo As Long
    Cargo = GetPlayerCargo(Index)
    If Cargo < 115 Then
        ServicesReq = 10 + ((Cargo - 99) * 20)
    Else
        ServicesReq = 5
    End If
    If Player(Index).NumServices < ServicesReq Then
        Call PlayerMsg(Index, "Você não concluiu o número de serviços para subir de cargo (Atualmente: " & Player(Index).NumServices & "/" & ServicesReq & ")", Yellow)
        Exit Sub
    End If
    
    'Iniciar provação
    Call TakeInvItem(Index, MoedaZ, Provação(ProvNum).Cost)
    Provação(ProvNum).ActualTick = GetTickCount
    Provação(ProvNum).ActualWave = 0
    Provação(ProvNum).ProvaçãoIndex = Index
    Call PlayerWarp(Index, Provação(ProvNum).Map, Provação(ProvNum).X, Provação(ProvNum).Y)
    Call PlayerMsg(Index, printf("Bem-vindo á provação de evolução de cargo, elimine as ondas de inimigos para completar a missão!"), Yellow)
    Call GlobalMsg(printf("%s acabou de entrar em uma provação de cargo! Que a sorte esteja com ele!", GetPlayerName(Index)), brightgreen)
    Call SendProvacaoState(Index, 1)
    
    Dim i As Long
    For i = 1 To MAX_MAP_NPCS
        MapNpc(Provação(ProvNum).Map).Npc(i).Num = 0
        SpawnNpc i, Provação(ProvNum).Map
    Next i
    
End Sub
Sub HealPlayer(ByVal Index As Long, Amount As Long)
    If GetPlayerVital(Index, HP) + Amount > GetPlayerMaxVital(Index, HP) Then
        SetPlayerVital Index, HP, GetPlayerMaxVital(Index, HP)
    Else
        SetPlayerVital Index, HP, GetPlayerVital(Index, HP) + Amount
    End If
    SendVital Index, HP
End Sub
Sub HealPlayerMP(ByVal Index As Long, Amount As Long)
    If GetPlayerVital(Index, MP) + Amount > GetPlayerMaxVital(Index, MP) Then
        SetPlayerVital Index, MP, GetPlayerMaxVital(Index, MP)
    Else
        SetPlayerVital Index, MP, GetPlayerVital(Index, MP) + Amount
    End If
    SendVital Index, MP
End Sub
Function Fisgada(ByVal Index As Long) As Boolean
    Fisgada = False
    If TempPlayer(Index).NextFish > GetTickCount - MarginFish And TempPlayer(Index).NextFish < GetTickCount + MarginFish Then
        Fisgada = True
    End If
End Function
Sub HandleGravity(ByVal Index As Long)
    Dim diff As Long
    diff = DateDiff("h", Trim$(Player(Index).GravityInit), Now)
    If diff >= Player(Index).GravityHours Then
        Dim ExpBonus As Long
        Dim Level As Long
        Level = Int((Player(Index).GravityValue / 1500) * 95)
        If Level <= 0 Then Level = 1
        ExpBonus = ExperienceBase(Level) * Player(Index).GravityHours '(Experience(Level) * 0.1) * Player(Index).GravityHours
        GivePlayerEXP Index, ExpBonus
        'PlayerMsg Index, "Você finalizou o treino na sala de gravidade e recebeu: " & ExpBonus & " XP", brightgreen
        SendDialogue Index, "Conclusão de treinamento", "Você finalizou seu treinamento na sala de gravidade e recebeu " & ExpBonus & " XP"
        Player(Index).GravityInit = vbNullString
        Player(Index).GravityHours = 0
        PlayerWarp Index, START_MAP, Int(Map(START_MAP).MaxX / 2), Int(Map(START_MAP).MaxY / 2)
    End If
End Sub
Sub CheckItemDrop(ByVal Index As Long)
    Dim i As Long, ItemNum As Long
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, i) > 0 Then
            ItemNum = GetPlayerInvItemNum(Index, i)
            If Item(ItemNum).BindType = 3 Then 'Drop on death
                PlayerMapDropItem Index, i, GetPlayerInvItemValue(Index, i)
            End If
        End If
    Next i
End Sub
Sub RegisterAOE(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal SpellNum As Long)
    Dim AoEIndex As Integer
    
    AoEIndex = FindOpenAoEIndex
    AoEEffect(AoEIndex).X = X
    AoEEffect(AoEIndex).Y = Y
    AoEEffect(AoEIndex).Tick = GetTickCount
    AoEEffect(AoEIndex).Duration = Spell(SpellNum).AoEDuration
    AoEEffect(AoEIndex).CastTick = GetTickCount
    AoEEffect(AoEIndex).Map = GetPlayerMap(Index)
    AoEEffect(AoEIndex).Caster = GetPlayerName(Index)
    AoEEffect(AoEIndex).SpellNum = SpellNum
    AoEEffect(AoEIndex).CasterType = TARGET_TYPE_PLAYER
    
End Sub
Sub RegisterNPCAOE(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal SpellNum As Long)
    Dim AoEIndex As Integer
    
    AoEIndex = FindOpenAoEIndex
    AoEEffect(AoEIndex).X = X
    AoEEffect(AoEIndex).Y = Y
    AoEEffect(AoEIndex).Tick = GetTickCount
    AoEEffect(AoEIndex).Duration = Spell(SpellNum).AoEDuration
    AoEEffect(AoEIndex).CastTick = GetTickCount
    AoEEffect(AoEIndex).Map = MapNum
    AoEEffect(AoEIndex).Caster = Index
    AoEEffect(AoEIndex).SpellNum = SpellNum
    AoEEffect(AoEIndex).CasterType = TARGET_TYPE_NPC
    
End Sub
Function FindOpenAoEIndex() As Integer
    Dim i As Long
    
    For i = 1 To MAX_AOEEFFECTS
        If AoEEffect(i).Tick + AoEEffect(i).Duration < GetTickCount Then
            FindOpenAoEIndex = i
            Exit Function
        End If
    Next i
End Function
Function GetPlayerSkillDamage(ByVal Index As Long, ByVal Vital As Long) As Long
    Dim weaponNum As Long
    weaponNum = GetPlayerEquipment(Index, Weapon)
    If weaponNum > 0 Then
        GetPlayerSkillDamage = (Vital + GetPlayerStat(Index, Intelligence) + Item(weaponNum).data2) / 100 * rand(80, 100)
    Else
        GetPlayerSkillDamage = (Vital + GetPlayerStat(Index, Intelligence)) / 100 * rand(80, 100)
    End If
End Function
Function FindOpenGuildSlot() As Long
    Dim i As Long
    For i = 1 To MAX_GUILDS
        If Guild(i).Level = 0 Then
            FindOpenGuildSlot = i
            Exit Function
        End If
    Next i
    SetStatus "Sem espaço para novas guilds!"
End Function
Function MaxGuildMembers(Level As Long) As Byte
    MaxGuildMembers = 3 + Int(Level / 5)
    If MaxGuildMembers > 10 Then MaxGuildMembers = 10
End Function
Function FindOpenGuildMemberSlot(GuildNum As Long) As Long
    Dim i As Long
    For i = 1 To MaxGuildMembers(Guild(GuildNum).Level)
        If Guild(GuildNum).Member(i).Level = 0 Then
            FindOpenGuildMemberSlot = i
            Exit Function
        End If
    Next i
End Function
Sub AddMember(ByVal GuildNum As Long, Index As Long, Rank As Byte)
    Dim i As Long
    i = FindOpenGuildMemberSlot(GuildNum)
    If i > 0 Then
        Guild(GuildNum).Member(i).Name = GetPlayerName(Index)
        Guild(GuildNum).Member(i).Rank = Rank
        Guild(GuildNum).Member(i).Level = GetPlayerLevel(Index)
        Guild(GuildNum).Member(i).GuildExp = 0
        Guild(GuildNum).Member(i).Donations = 0
        Player(Index).Guild = GuildNum
        PlayerMsg Index, "Você foi aceito na guild " & Trim$(Guild(GuildNum).Name), Yellow
        PlayerMsg Index, "Você pode acessar o painel da sua guild apertando F1! Para falar no chat da guild utilize -(mensagem)", brightgreen
        SaveGuild GuildNum
        SavePlayer Index
        SendPlayerData Index
        SendUpdateGuildToAll GuildNum
    Else
        PlayerMsg Index, "A guild está cheia!", brightred
    End If
End Sub
Public Sub UpdateGuildLevel(ByVal Index As Long)
    Dim GuildNum As Long
    Dim GuildIndex As Long
    GuildNum = Player(Index).Guild
    If GuildNum > 0 Then
        GuildIndex = GetPlayerGuildIndex(Index)
        If GuildIndex > 0 Then
            Guild(GuildNum).Member(GuildIndex).Level = GetPlayerLevel(Index)
            SaveGuild GuildNum
            SendUpdateGuildToAll GuildNum
        End If
    End If
End Sub

Public Function GetPlayerGuildIndex(ByVal Index As Long) As Long
    Dim GuildNum As Long
    Dim i As Long
    
    GuildNum = Player(Index).Guild
    
    If GuildNum > 0 Then
        For i = 1 To 10
            If LCase(Trim$(Guild(GuildNum).Member(i).Name)) = LCase(Trim$(GetPlayerName(Index))) Then
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
Function IsGuildFull(ByVal GuildNum As Long) As Boolean
    IsGuildFull = (FindOpenGuildMemberSlot(GuildNum) = 0)
End Function
Function IsGuildEmpty(ByVal GuildNum As Long) As Boolean
    Dim i As Long
    For i = 1 To 10
        If Guild(GuildNum).Member(i).Level > 0 And Not Guild(GuildNum).Member(i).Rank = GuildRank.Mestre Then
            IsGuildEmpty = False
            Exit Function
        End If
    Next i
    IsGuildEmpty = True
End Function
Sub GiveGuildExp(ByVal Index As Long, GuildExp As Long, Level As Long)
    Dim ReqEspeciarias As Long, ReqGold As Long
    Dim GuildLevel As Long, GuildNum As Long
    Dim GuildMemberIndex As Long
    
    GuildNum = Player(Index).Guild
    If GuildNum > 0 Then
        If Not Guild(GuildNum).UpBlock = 1 Then
            GuildLevel = Guild(GuildNum).Level
            'ReqEspeciarias = Int(Experience(Level) * 0.001)
            ReqGold = Experience(Level)
            If Guild(GuildNum).Gold >= ReqGold Then
                'Guild(GuildNum).Red = Guild(GuildNum).Red - ReqEspeciarias
                'Guild(GuildNum).Blue = Guild(GuildNum).Blue - ReqEspeciarias
                'Guild(GuildNum).Yellow = Guild(GuildNum).Yellow - ReqEspeciarias
                Guild(GuildNum).Gold = Guild(GuildNum).Gold - ReqGold
                Guild(GuildNum).Exp = Guild(GuildNum).Exp + GuildExp
                GuildMemberIndex = GetPlayerGuildIndex(Index)
                Guild(GuildNum).Member(GuildMemberIndex).GuildExp = Guild(GuildNum).Member(GuildMemberIndex).GuildExp + GuildExp
                CheckGuildLevelUp GuildNum
                GuildMsg GuildNum, ReqGold & "z foram consumidos por " & GetPlayerName(Index) & " para garantir " & GuildExp & "xp para a guild!", brightgreen
                SaveGuild GuildNum
                SendUpdateGuildToAll GuildNum
            Else
                PlayerMsg Index, "Sua guild não tem fundos o suficiente para receber experiência (Requisitos: " & ReqGold & "z)", brightred
            End If
        Else
            PlayerMsg Index, "A evolução de sua guild foi travada pelo mestre", brightred
        End If
    End If
End Sub

Sub GiveGuildFreeExp(ByVal Index As Long, GuildExp As Long)
    Dim GuildNum As Long
    Dim GuildMemberIndex As Long
    
    GuildNum = Player(Index).Guild
    If GuildNum > 0 Then
        Guild(GuildNum).Exp = Guild(GuildNum).Exp + GuildExp
        GuildMemberIndex = GetPlayerGuildIndex(Index)
        Guild(GuildNum).Member(GuildMemberIndex).GuildExp = Guild(GuildNum).Member(GuildMemberIndex).GuildExp + GuildExp
        CheckGuildLevelUp GuildNum
        SaveGuild GuildNum
        SendUpdateGuildToAll GuildNum
    End If
End Sub

Sub CheckGuildLevelUp(ByVal GuildNum As Long)
    Dim i As Long
    Do While Guild(GuildNum).Exp >= Experience(Guild(GuildNum).Level)
        Guild(GuildNum).Exp = Guild(GuildNum).Exp - Experience(Guild(GuildNum).Level)
        Guild(GuildNum).Level = Guild(GuildNum).Level + 1
        Guild(GuildNum).TNL = Experience(Guild(GuildNum).Level)
        GuildMsg GuildNum, "A guild evoluiu para o nível " & Guild(GuildNum).Level & "!", brightgreen
    Loop
End Sub

Sub HandleMission(ByVal Index As Long)
    Dim NewMission As Boolean
    NewMission = False
    If IsDate(Player(Index).Daily.LastDate) = False Then
        NewMission = True
    Else
        Dim diff As Long
        diff = DateDiff("d", Trim$(Player(Index).Daily.LastDate), Now)
        If diff > 0 Then
            NewMission = True
            If diff > 1 Then Player(Index).Daily.DailyBonus = 0
        End If
    End If
    
    If NewMission Then
        Player(Index).Daily.LastDate = Now
        Player(Index).Daily.MissionActual = 0
        Player(Index).Daily.Completed = 0
        Player(Index).Daily.MissionIndex = rand(1, UBound(DailyMission))
        Player(Index).Daily.MissionObjective = DailyMission(Player(Index).Daily.MissionIndex).NumberFactory * GetPlayerLevel(Index)
        SavePlayer Index
    End If
End Sub

Function HandleLastLogin(ByVal Index As Long)
    If IsDate(Player(Index).LastLogin) = True Then
        If Player(Index).VIP > 0 Then
            Dim diff As Long
            diff = DateDiff("d", Trim$(Player(Index).LastLogin), Now)
            If diff > 1 Then
                PlayerMsg Index, "Você passou mais de um dia sem entrar no jogo e por isso a experiência de seu vip foi resetada!", brightred
                Player(Index).VIPExp = 0
            End If
        End If
    End If
    Player(Index).LastLogin = Now
End Function

Function IsDaily(ByVal Index As Long, Quest As Daily) As Boolean
    If Player(Index).Daily.Completed = 0 Then
        If Player(Index).Daily.MissionIndex = Quest Then IsDaily = True
    End If
End Function

Sub UpdateDaily(ByVal Index As Long, Optional Value As Long = 1)
    Player(Index).Daily.MissionActual = Player(Index).Daily.MissionActual + Value
    
    If Player(Index).Daily.MissionActual >= Player(Index).Daily.MissionObjective Then
        Dim Exp As Long
        Exp = ExperienceBase(GetPlayerLevel(Index)) * 3
        PlayerMsg Index, "Parabéns! Você completou sua missão diária e recebeu " & Exp & "xp", brightgreen
        If Player(Index).Guild > 0 Then
            GiveGuildFreeExp Index, GetPlayerLevel(Index)
            GuildMsg Player(Index).Guild, GetPlayerName(Index) & " completou sua missão diária e garantiu " & GetPlayerLevel(Index) & "xp para a guild!", brightgreen
        End If
        Player(Index).Daily.Completed = 1
        Player(Index).Daily.DailyBonus = Player(Index).Daily.DailyBonus + 1
        
        Dim ItemNum As Long
        Dim Quant As Long
        Select Case Player(Index).Daily.DailyBonus
            Case 1
                ItemNum = MoedaZ
                Quant = rand(50000, 80000)
            Case 2
                ItemNum = MoedaZ
                Quant = rand(80000, 120000)
            Case 3
                ItemNum = MoedaZ
                Quant = rand(120000, 200000)
            Case 4
                ItemNum = rand(80, 82)
                Quant = rand(500, 800) - ((ItemNum - 80) * 100)
            Case 5
                ItemNum = 141
                Quant = 1
            Case 6
                ItemNum = 141
                Quant = 2
            Case 7
                ItemNum = 144
                Quant = 1
        End Select
        
        GiveInvItem Index, ItemNum, Quant
        PlayerMsg Index, "Você recebeu: " & Quant & " " & Trim$(Item(ItemNum).Name) & "!", Yellow
        
        GiveInvItem Index, 185, 1
        PlayerMsg Index, "Você recebeu uma caixa planetária!", Yellow
        
        If Player(Index).Daily.DailyBonus >= 7 Then Player(Index).Daily.DailyBonus = 0
    End If
    
    SendPlayerDailyQuest Index
End Sub
Function VIPBonus(ByVal Index As Long) As Single
    VIPBonus = 1.5 + (0.05 * (Player(Index).VIP - 1))
End Function
Sub GivePlayerVIPExp(ByVal Index As Long, Exp As Long)
    If Player(Index).VIP > 0 Then
        Player(Index).VIPExp = Player(Index).VIPExp + Exp
        If Player(Index).VIP = 50 Then
            Player(Index).VIPExp = 0
            Exit Sub
        End If
        
        Do While Player(Index).VIPExp >= GetPlayerVipNextLevel(Index)
            Player(Index).VIPExp = Player(Index).VIPExp - GetPlayerVipNextLevel(Index)
            Player(Index).VIP = Player(Index).VIP + 1
            PlayerMsg Index, "Parabéns! Seu VIP evoluiu para o level " & Player(Index).VIP, Yellow
        Loop
    End If
End Sub
Sub CheckArenaStart()
    Dim i As Long
    Dim PlayerIndex As Long
    Dim Players As String
    For i = 1 To ArenaChallenge.TotalPlayers
        If ArenaChallenge.PlayerAccept(i) = False Then Exit Sub
    Next i
    
    'Start arena
    For i = 1 To ArenaChallenge.TotalPlayers
        PlayerIndex = FindPlayer(ArenaChallenge.Players(i))
        If PlayerIndex > 0 Then
            If Not HasItem(PlayerIndex, MoedaZ) >= ArenaChallenge.Aposta Then
                GlobalMsg "O jogador " & GetPlayerName(PlayerIndex) & " não tem o dinheiro necessário para a aposta de " & ArenaChallenge.Aposta & "z! Arena cancelada!", brightred
                ArenaChallenge.Active = 0
                Exit Sub
            Else
                If i < ArenaChallenge.TotalPlayers Then
                    Players = GetPlayerName(PlayerIndex) & ", "
                Else
                    Players = GetPlayerName(PlayerIndex)
                End If
            End If
        Else
            GlobalMsg "O jogador " & ArenaChallenge.Players(i) & " está offline, arena cancelada!", brightred
            ArenaChallenge.Active = 0
            Exit Sub
        End If
    Next i
    
    GlobalMsg "A arena entre os jogadores: " & Players & " irá começar! Façam suas apostas!", Yellow
    'Retirar dinheiro, stunar e teleportar
    ArenaChallenge.Active = 2
    ArenaChallenge.Tick = GetTickCount
    
    'Find Positions
    Dim X As Long, Y As Long, MapNum As Long
    Dim PositionX(1 To 6) As Long, PositionY(1 To 6) As Long
    Dim Pos As Long
    MapNum = ARENA_MAP
    For Y = 1 To Map(MapNum).MaxY
        For X = 1 To Map(MapNum).MaxX
            If Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_NPCAVOID Then
                Pos = Pos + 1
                PositionX(Pos) = X
                PositionY(Pos) = Y
                If Pos >= 6 Then Exit For
            End If
        Next X
    Next Y
    
    For i = 1 To ArenaChallenge.TotalPlayers
        PlayerIndex = FindPlayer(ArenaChallenge.Players(i))
        TakeInvItem PlayerIndex, MoedaZ, ArenaChallenge.Aposta
        PlayerWarp PlayerIndex, MapNum, PositionX(i), PositionY(i)
        ServerStunPlayer PlayerIndex, 3
    Next i
    
    ArenaChallenge.Count = 4
    ArenaChallenge.CountTick = GetTickCount
    
End Sub
Sub HandleArena()
    If ArenaChallenge.Active = 1 Then
        If ArenaChallenge.LastCall + 20000 < GetTickCount Then
            GlobalMsg "O pedido de arena excedeu o limite de tempo para as respostas, desafio cancelado!", brightred
            ArenaChallenge.Active = 0
        End If
    End If
    If ArenaChallenge.Active = 2 Then
        If ArenaChallenge.Tick + 120000 < GetTickCount Then
            GlobalMsg "O tempo limite da arena foi finalizado! Batalha cancelada!", brightred
            ArenaChallenge.Active = 0
            Evacuate ARENA_MAP, True
            Exit Sub
        End If
        If ArenaChallenge.Count > 0 Then
            If ArenaChallenge.CountTick + 1000 < GetTickCount Then
                If ArenaChallenge.Count > 1 Then
                    SendBossMsg ARENA_MAP, "A batalha começa em " & ArenaChallenge.Count - 1 & "s", Yellow
                Else
                    SendBossMsg ARENA_MAP, "BATALHEM!", Yellow
                End If
                ArenaChallenge.Count = ArenaChallenge.Count - 1
                ArenaChallenge.CountTick = GetTickCount
            End If
        End If
        Dim i As Long, Index As Long
        Dim AlivePlayers As Long
        Dim AliveIndex As Long
        For i = 1 To ArenaChallenge.TotalPlayers
            Index = FindPlayer(Trim$(ArenaChallenge.Players(i)))
            If Index > 0 Then
                If Player(Index).IsDead = 0 And GetPlayerMap(Index) = ARENA_MAP Then
                    AlivePlayers = AlivePlayers + 1
                    AliveIndex = Index
                End If
            End If
        Next i
        If AlivePlayers = 1 Then
            'Vencedor!
            Dim Prize As Long
            Dim Remove As Long
            Prize = (ArenaChallenge.Aposta * ArenaChallenge.TotalPlayers)
            PlayerWarp AliveIndex, START_MAP, START_X, START_Y
            If Prize = 0 Then
                GlobalMsg "O jogador " & GetPlayerName(AliveIndex) & " foi o vencedor do desafio da arena!", Yellow
            Else
                GlobalMsg "O jogador " & GetPlayerName(AliveIndex) & " foi o vencedor do desafio da arena ganhando um prêmio de " & Prize & "z!", Yellow
            End If
            GiveInvItem AliveIndex, MoedaZ, Prize
            ArenaChallenge.Active = 0
        End If
    End If
End Sub
Sub CheckSpecialEvent(ByVal Index As Long, ByVal TargetPlayer As Long, ByVal PartyNum As Long)
    Dim i As Long
    Dim CanGive As Boolean
    Dim CaptureItem As Long
    CanGive = True
    CaptureItem = 152
    For i = 1 To MAX_PARTY_MEMBERS
        If Party(PartyNum).Member(i) = 0 Then
            CanGive = False
        Else
            If HasItem(Party(PartyNum).Member(i), CaptureItem) > 0 Then
                CanGive = False
            Else
                If Player(Party(PartyNum).Member(i)).PlanetNum > 0 Then
                    CanGive = False
                End If
            End If
        End If
    Next
    
    If CanGive Then
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(PartyNum).Member(i) > 0 Then
                GiveInvItem Party(PartyNum).Member(i), CaptureItem, 1
                PlayerMsg Party(PartyNum).Member(i), "Parabéns! Você recebeu um item de captura de planeta! Escolha bem e capture já o seu! Acesse nosso fórum para tirar suas dúvidas!", brightgreen
            End If
        Next
    Else
        PlayerMsg Index, "Aproveite o evento deste fim de semana! Com um grupo de 4 pessoas que NÃO TEM UM PLANETA PRÓPRIO todos receberão o item para capturar um planeta!", Yellow
    End If
End Sub
Sub ProgressConquista(ByVal Index As Long, Conquista As Long, Optional Value As Long = 1)
    If Player(Index).Conquistas(Conquista) = 0 Then
        Player(Index).ConquistaProgress(Conquista) = Player(Index).ConquistaProgress(Conquista) + Value
        
        If Player(Index).ConquistaProgress(Conquista) >= Conquistas(Conquista).Progress Then
            Player(Index).Conquistas(Conquista) = 1
            
            GivePlayerEXP Index, Conquistas(Conquista).Exp
            PlayerMsg Index, "Você completou a conquista: " & Trim$(Conquistas(Conquista).Name), brightgreen
            
            Dim i As Long
            For i = 1 To 5
                If Conquistas(Conquista).Reward(i).Num > 0 Then
                    GiveInvItem Index, Conquistas(Conquista).Reward(i).Num, Conquistas(Conquista).Reward(i).Value
                    PlayerMsg Index, "Você recebeu " & Conquistas(Conquista).Reward(i).Value & " " & Trim$(Item(Conquistas(Conquista).Reward(i).Num).Name), brightgreen
                End If
            Next i
        End If
        
        SendPlayerConquista Index, Conquista
        SavePlayer Index
    End If
End Sub
