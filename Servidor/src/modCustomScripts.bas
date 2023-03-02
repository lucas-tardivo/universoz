Attribute VB_Name = "modCustomScripts"
Public Function CustomScript(Index As Long, caseID As Long) As Boolean
    Dim i As Long
    CustomScript = True
    If Not UZ Then
        Select Case caseID
        
            Case 2, 3, 4, 5 'Transportes
                i = caseID - 1
                    If Transporte(i).State = 1 Then
                        If Transporte(i).Tick + Transporte(i).Embarque < GetTickCount Then
                            If Transporte(i).Passaporte > 0 Then
                                If HasItem(Index, Transporte(i).Passaporte) = 0 Then
                                    Call PlayerMsg(Index, printf("Você precisa ter um passaporte para entrar neste transporte!"), brightred)
                                    Exit Function
                                End If
                                Call TakeInvItem(Index, Transporte(i).Passaporte, 1)
                            End If
                            Call PlayerWarp(Index, Transporte(i).LoadMap, Transporte(i).LoadX, Transporte(i).LoadY)
                            Exit Function
                        Else
                            Call PlayerMsg(Index, printf("Quanta pressa! O %s já está chegando!", Transporte(i).Nome), brightred)
                        End If
                    Else
                        Call PlayerMsg(Index, printf("O %s ainda não chegou!", Transporte(i).Nome), brightred)
                    End If
            Exit Function
    
            Case 6 'Vegeta
                Call SendAnimation(GetPlayerMap(Index), 29, MapNpc(GetPlayerMap(Index)).Npc(2).X, MapNpc(GetPlayerMap(Index)).Npc(2).Y, 1, , , , , 3)
                Call SpawnNpc(3, GetPlayerMap(Index), True, 1, DIR_DOWN)
                MapNpc(GetPlayerMap(Index)).Npc(3).X = MapNpc(GetPlayerMap(Index)).Npc(2).X
                MapNpc(GetPlayerMap(Index)).Npc(3).Y = MapNpc(GetPlayerMap(Index)).Npc(2).Y
                SendMapNpcXY 3, GetPlayerMap(Index)
            Exit Function
                
            Case 7 'Vegeta morrendo
                Call SendAnimation(GetPlayerMap(Index), 40, MapNpc(GetPlayerMap(Index)).Npc(7).X, MapNpc(GetPlayerMap(Index)).Npc(7).Y, 1)
            Exit Function
                
            Case 10, 11, 12, 13
                Dim ItemNum As Long, Value As Long
                If caseID = 10 Then
                    ItemNum = 18 + rand(0, 6)
                    Value = 20000
                End If
                If caseID = 11 Then
                    ItemNum = 51 + rand(0, 5)
                    Value = 15000
                End If
                If caseID = 12 Then
                    ItemNum = 57 + rand(0, 5)
                    Value = 32000
                End If
                If caseID = 13 Then
                    ItemNum = 45 + rand(0, 5)
                    Value = 8000
                End If
                Call GiveInvItem(Index, ItemNum, 1, True)
                Call TakeInvItem(Index, MoedaZ, Value)
                Call PlayerMsg(Index, "Você recebeu: " & Trim$(Item(ItemNum).Name), Yellow)
            Exit Function
                
            Case 15 'Refine
                SendOpenRefine Index
            Exit Function
                
            Case 18, 19, 20
                Dim prov As Byte
                prov = caseID - 17
                Call InitProvação(Index, prov)
            Exit Function
            
            Case 21 'Sala de gravidade
                SendGravidade Index
            Exit Function
            
            Case 25
                SendOpenGuildMaster Index
            Exit Function
            
            Case Else
                PlayerMsg Index, "You just activated custom script " & caseID & ". This script is not yet programmed.", brightred
        End Select
    Else
        Select Case caseID
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9
                prov = caseID
                Call InitProvação(Index, prov)
                CustomScript = False
            Exit Function
            
            Case 10 'Vegeta
                Call SendAnimation(GetPlayerMap(Index), 29, MapNpc(GetPlayerMap(Index)).Npc(6).X, MapNpc(GetPlayerMap(Index)).Npc(6).Y, 1, , , , , 7)
                Call SpawnNpc(7, GetPlayerMap(Index), True, 1, DIR_DOWN)
                MapNpc(GetPlayerMap(Index)).Npc(7).X = MapNpc(GetPlayerMap(Index)).Npc(6).X
                MapNpc(GetPlayerMap(Index)).Npc(7).Y = MapNpc(GetPlayerMap(Index)).Npc(6).Y
                SendMapNpcXY 7, GetPlayerMap(Index)
            Exit Function
                
            Case 11 'Vegeta morrendo
                Call SendAnimation(GetPlayerMap(Index), 40, MapNpc(GetPlayerMap(Index)).Npc(7).X, MapNpc(GetPlayerMap(Index)).Npc(7).Y, 1)
            Exit Function
            
            Case 15 'Refine
                SendOpenRefine Index
            Exit Function
            
            Case 21 'Sala de gravidade
                SendGravidade Index
            Exit Function
            
            Case 22
                If Player(Index).EsoNum <> ESOTERICAAUTO And Player(Index).EsoNum <> 0 Then
                    PlayerMsg Index, "Você não pode receber o bonus pois está sobre o efeito de outra esotérica!", brightred
                    Exit Function
                End If
                Player(Index).EsoBonus = Player(Index).EsoBonus + 10
                If Player(Index).EsoBonus > 30 Then Player(Index).EsoBonus = 30
                Player(Index).EsoTime = 30
                Player(Index).EsoNum = ESOTERICAAUTO
                Call PlayerMsg(Index, "Você recebeu o bonus de +10% de exp e dinheiro por meia hora! (Total de bonus: " & Player(Index).EsoBonus & "%)", brightgreen)
                SendPlayerData Index
                SavePlayer Index
            Exit Function
            
            Case 23
                Dim MapNum As Long, X As Long, Y As Long
                MapNum = TesouroMap + 1
                X = rand(1, Map(MapNum).MaxX)
                Y = rand(1, Map(MapNum).MaxY)
                
                Do While Map(MapNum).Tile(X, Y).Type <> TileType.TILE_TYPE_WALKABLE
                    X = rand(1, Map(MapNum).MaxX)
                    Y = rand(1, Map(MapNum).MaxY)
                Loop
                
                PlayerWarp Index, MapNum, X, Y
            Exit Function
            
            Case 24
                'Reset PDL
                For i = 1 To Stats.Stat_Count - 1
                    SetPlayerStatPoints Index, i, 0
                    SetPlayerStat Index, i, 1
                Next i
                SetPlayerPOINTS Index, GetPlayerLevel(Index) * 3
                SendStats Index
            Exit Function
            
            Case 25
                SendOpenGuildMaster Index
            Exit Function
            
            Case 26
                SendOpenArena Index
            Exit Function
            
            Case 27
                Viajar Index, VirgoMap
            Exit Function
            
            Case 28
                Call SetPlayerSprite(Index, GetPlayerNormalSprite(Index))
                Call PlayerWarp(Index, 2, Map(2).MaxX / 2, Map(2).MaxY / 2)
            Exit Function
            
            Case 29
                Call SendOpenTroca(Index)
            Exit Function
            
            Case 30
                Dim Difference As Long
                Dim DifferencePlanet As Long
                DifferencePlanet = 1
                Difference = Abs(GetPlayerLevel(Index) - Planets(1).Level)
                For i = 1 To MAX_PLANETS
                    If Planets(i).TimeToExplode = 0 Then
                        If Trim$(LCase(Planets(i).Name)) <> "planeta desconhecido" And Not PlanetInService(i) Then
                            If Planets(i).Level <= GetPlayerLevel(Index) Then
                                If Difference >= Abs(GetPlayerLevel(Index) - Planets(i).Level) Then
                                    If Difference = Abs(GetPlayerLevel(Index) - Planets(i).Level) And rand(1, 2) = 1 Then
                                        DifferencePlanet = i
                                        Difference = GetPlayerLevel(Index) - Planets(i).Level
                                    Else
                                        If TempPlayer(Index).PlanetService = 0 Or (TempPlayer(Index).PlanetService > 0 And i <> TempPlayer(Index).PlanetService) Then
                                            DifferencePlanet = i
                                            Difference = GetPlayerLevel(Index) - Planets(i).Level
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
                If TempPlayer(Index).PlanetService <> 0 And TempPlayer(Index).PlanetService <> DifferencePlanet Then PlanetInService(TempPlayer(Index).PlanetService) = False
                TempPlayer(Index).PlanetService = DifferencePlanet
                PlanetInService(DifferencePlanet) = True
                SendPlayerData Index
                PlayerMsg Index, "Você recebeu um serviço!", Yellow
            Exit Function
            
            Case 31
                Dim Total As Long
                Total = rand(TotalOnlinePlayers * 2, TotalOnlinePlayers * 5)
                For i = 1 To Total
                    X = rand(2, Map(2).MaxX - 2)
                    Y = rand(2, Map(2).MaxY - 2)
                    If Map(2).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE Then
                        If rand(1, 2) = 1 Then
                            ItemNum = 141
                        Else
                            ItemNum = 144
                        End If
                        Call SpawnItem(ItemNum, 1, 2, X, Y)
                        SendAnimation 2, Item(ItemNum).Animation, X, Y, 0
                    End If
                Next i
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If GetPlayerMap(i) = 2 Then
                            PlayerMsg i, "Você ajudou contra a invasão e recebeu " & (GetPlayerNextLevel(i) / 10) & "xp", brightgreen
                            GivePlayerEXP i, GetPlayerNextLevel(i) / 20
                        End If
                    End If
                Next i
                DesactivateEvent
            Exit Function
            
            Case 32
                If Player(Index).PlayerHouseNum > 0 Then
                    PlayerWarp Index, Player(Index).PlayerHouseNum, Map(Player(Index).PlayerHouseNum).MaxX / 2, Map(Player(Index).PlayerHouseNum).MaxY / 2
                Else
                    Viajar Index, 1
                End If
            Exit Function
            
            Case 33
                
            Exit Function
            
            Case Else
                PlayerMsg Index, "You just activated custom script " & caseID & ". This script is not yet programmed.", brightred
        End Select
    End If
End Function
