Attribute VB_Name = "modGameLogic"
Option Explicit
Private InitServerTick As Long

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal PlayerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, ItemVal, MapNum, X, Y, PlayerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal PlayerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
            MapItem(MapNum, i).PlayerName = PlayerName
            MapItem(MapNum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(MapNum, i).canDespawn = canDespawn
            MapItem(MapNum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(MapNum, i).Num = ItemNum
            MapItem(MapNum, i).Value = ItemVal
            MapItem(MapNum, i).X = X
            MapItem(MapNum, i).Y = Y
            ' send to map
            SendSpawnItemToMap MapNum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim X As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(X, Y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(X, Y).data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(MapNum).Tile(X, Y).data1).Stackable > 0 And Map(MapNum).Tile(X, Y).data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(X, Y).data1, 1, MapNum, X, Y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(X, Y).data1, Map(MapNum).Tile(X, Y).data2, MapNum, X, Y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Dim RandomNumber As Long
    Randomize RandomNumber
    RandomNumber = ((High - Low + 1) * Rnd) + Low
    Random = RandomNumber
End Function

Public Sub SpawnNpc(ByVal MapNPCNum As Long, ByVal MapNum As Long, Optional ForcedSpawn As Boolean = False, Optional SpawnDelay As Byte = 0, Optional Dir As Byte = 0)
    Dim Buffer As clsBuffer
    Dim NpcNum As Long
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    NpcNum = Map(MapNum).Npc(MapNPCNum)
    If ForcedSpawn = False And Map(MapNum).NpcSpawnType(MapNPCNum) = 1 Then NpcNum = 0
    If NpcNum > 0 Then
        MapNpc(MapNum).Npc(MapNPCNum).PDL = Npc(NpcNum).Level
        If UZ Then
            MapNpc(MapNum).Npc(MapNPCNum).Level = 1
            If MapNum >= PlanetStart And MapNum < PlanetStart + MAX_PLANET_BASE Then
                Dim PlanetNum As Long
                PlanetNum = GetPlanetNum(MapNum)
                If Planets(PlanetNum).Map = MapNum Then 'Confirmar
                    MapNpc(MapNum).Npc(MapNPCNum).Level = Planets(PlanetNum).Level
                    MapNpc(MapNum).Npc(MapNPCNum).PDL = (PDLBase(Planets(PlanetNum).Level) / 100) * Npc(NpcNum).Level
                End If
            End If
            Dim ProvaçãoNum As Long
            ProvaçãoNum = getProvação(MapNum)
            If ProvaçãoNum > 0 Then
                If Provação(ProvaçãoNum).MinLevel < MAX_LEVELS Then
                    MapNpc(MapNum).Npc(MapNPCNum).Level = Provação(ProvaçãoNum).MinLevel
                    MapNpc(MapNum).Npc(MapNPCNum).PDL = (PDLBase(Provação(ProvaçãoNum).MinLevel) / 100) * Npc(NpcNum).Level
                Else
                    MapNpc(MapNum).Npc(MapNPCNum).Level = MAX_LEVELS
                    MapNpc(MapNum).Npc(MapNPCNum).PDL = MAX_LEVELS
                End If
            End If
        End If
    
        MapNpc(MapNum).Npc(MapNPCNum).Num = NpcNum
        MapNpc(MapNum).Npc(MapNPCNum).Target = 0
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = 0 ' clear
        MapNpc(MapNum).Npc(MapNPCNum).Spawned = 0
        
        MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) = GetNpcMaxVital(MapNum, MapNPCNum, Vitals.HP)
        MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.MP) = GetNpcMaxVital(MapNum, MapNPCNum, Vitals.MP)
        
        MapNpc(MapNum).Npc(MapNPCNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For X = 0 To Map(MapNum).MaxX
            For Y = 0 To Map(MapNum).MaxY
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(MapNum).Tile(X, Y).data1 = MapNPCNum Then
                        MapNpc(MapNum).Npc(MapNPCNum).X = X
                        MapNpc(MapNum).Npc(MapNPCNum).Y = Y
                        MapNpc(MapNum).Npc(MapNPCNum).Dir = Map(MapNum).Tile(X, Y).data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next Y
        Next X
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                X = Random(0, Map(MapNum).MaxX)
                Y = Random(0, Map(MapNum).MaxY)
    
                If X > Map(MapNum).MaxX Then X = Map(MapNum).MaxX
                If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(MapNum, X, Y) Then
                    MapNpc(MapNum).Npc(MapNPCNum).X = X
                    MapNpc(MapNum).Npc(MapNPCNum).Y = Y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For X = 0 To Map(MapNum).MaxX
                For Y = 0 To Map(MapNum).MaxY

                    If NpcTileIsOpen(MapNum, X, Y) Then
                        MapNpc(MapNum).Npc(MapNPCNum).X = X
                        MapNpc(MapNum).Npc(MapNPCNum).Y = Y
                        Spawned = True
                    End If

                Next
            Next

        End If
        
        If Dir > 0 Then MapNpc(MapNum).Npc(MapNPCNum).Dir = Dir

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Num
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir
            Buffer.WriteByte SpawnDelay
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
            UpdateMapBlock MapNum, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, True
        End If
        
        SendMapNpcVitals MapNum, MapNPCNum
    Else
        MapNpc(MapNum).Npc(MapNPCNum).Num = 0
        MapNpc(MapNum).Npc(MapNPCNum).Target = 0
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = 0 ' clear
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong MapNPCNum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If

End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = X Then
                    If GetPlayerY(LoopI) = Y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum).Npc(LoopI).Num > 0 Then
            If MapNpc(MapNum).Npc(LoopI).X = X Then
                If MapNpc(MapNum).Npc(LoopI).Y = Y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next
    
    CacheMapBlocks MapNum

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub
Function CanNpcMove(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim X As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Function
    End If

    X = MapNpc(MapNum).Npc(MapNPCNum).X
    Y = MapNpc(MapNum).Npc(MapNPCNum).Y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP_LEFT



            ' Check to make sure not outside of boundries

            If Y > 0 And X > 0 Then

                n = Map(MapNum).Tile(X - 1, Y - 1).Type



                ' Check to make sure that the tile is walkable

                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then

                    CanNpcMove = False

                    Exit Function

                End If



                ' Check to make sure that there is not a player in the way

                For i = 1 To Player_HighIndex

                    If IsPlaying(i) Then

                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y - 1) Then

                            CanNpcMove = False

                            Exit Function

                        End If

                    End If

                Next



                ' Check to make sure that there is not another npc in the way

                For i = 1 To MAX_MAP_NPCS

                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X - 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y - 1) Then

                        CanNpcMove = False

                        Exit Function

                    End If

                Next

               

                ' Directional blocking

                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_UP + 1) And isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_LEFT + 1) Then

                    CanNpcMove = False

                    Exit Function

                End If

            Else

                CanNpcMove = False

            End If

           

        Case DIR_UP_RIGHT



            ' Check to make sure not outside of boundries

            If Y > 0 And X < Map(MapNum).MaxX Then

                n = Map(MapNum).Tile(X + 1, Y - 1).Type



                ' Check to make sure that the tile is walkable

                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then

                    CanNpcMove = False

                    Exit Function

                End If



                ' Check to make sure that there is not a player in the way

                For i = 1 To Player_HighIndex

                    If IsPlaying(i) Then

                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y - 1) Then

                            CanNpcMove = False

                            Exit Function

                        End If

                    End If

                Next



                ' Check to make sure that there is not another npc in the way

                For i = 1 To MAX_MAP_NPCS

                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X + 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y - 1) Then

                        CanNpcMove = False

                        Exit Function

                    End If

                Next

               

                ' Directional blocking

                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_UP + 1) And isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_RIGHT + 1) Then

                    CanNpcMove = False

                    Exit Function

                End If

            Else

                CanNpcMove = False

            End If

           

        Case DIR_DOWN_LEFT



            ' Check to make sure not outside of boundries

            If Y < Map(MapNum).MaxY And X > 0 Then

                n = Map(MapNum).Tile(X - 1, Y + 1).Type



                ' Check to make sure that the tile is walkable

                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then

                    CanNpcMove = False

                    Exit Function

                End If



                ' Check to make sure that there is not a player in the way

                For i = 1 To Player_HighIndex

                    If IsPlaying(i) Then

                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y + 1) Then

                            CanNpcMove = False

                            Exit Function

                        End If

                    End If

                Next



                ' Check to make sure that there is not another npc in the way

                For i = 1 To MAX_MAP_NPCS

                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X - 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y + 1) Then

                        CanNpcMove = False

                        Exit Function

                    End If

                Next

               

                ' Directional blocking

                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_DOWN + 1) And isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_LEFT + 1) Then

                    CanNpcMove = False

                    Exit Function

                End If

            Else

                CanNpcMove = False

            End If

           

        Case DIR_DOWN_RIGHT



            ' Check to make sure not outside of boundries

            If Y < Map(MapNum).MaxY And X < Map(MapNum).MaxX Then

                n = Map(MapNum).Tile(X + 1, Y + 1).Type



                ' Check to make sure that the tile is walkable

                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then

                    CanNpcMove = False

                    Exit Function

                End If



                ' Check to make sure that there is not a player in the way

                For i = 1 To Player_HighIndex

                    If IsPlaying(i) Then

                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y + 1) Then

                            CanNpcMove = False

                            Exit Function

                        End If

                    End If

                Next



                ' Check to make sure that there is not another npc in the way

                For i = 1 To MAX_MAP_NPCS

                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X + 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y + 1) Then

                        CanNpcMove = False

                        Exit Function

                    End If

                Next

               

                ' Directional blocking

                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_DOWN + 1) And isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_RIGHT + 1) Then

                    CanNpcMove = False

                    Exit Function

                End If

            Else

                CanNpcMove = False

            End If
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = Map(MapNum).Tile(X, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(X, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = Map(MapNum).Tile(X - 1, Y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X - 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If X < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(X + 1, Y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNPCNum).X + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNPCNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).X = MapNpc(MapNum).Npc(MapNPCNum).X + 1) And (MapNpc(MapNum).Npc(i).Y = MapNpc(MapNum).Npc(MapNPCNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(MapNPCNum).Dir = Dir
    UpdateMapBlock MapNum, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, False

    Select Case Dir
        Case DIR_UP_LEFT

            MapNpc(MapNum).Npc(MapNPCNum).Y = MapNpc(MapNum).Npc(MapNPCNum).Y - 1

            MapNpc(MapNum).Npc(MapNPCNum).X = MapNpc(MapNum).Npc(MapNPCNum).X - 1

            Set Buffer = New clsBuffer

            Buffer.WriteLong SNpcMove

            Buffer.WriteLong MapNPCNum

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir

            Buffer.WriteLong movement

            SendDataToMap MapNum, Buffer.ToArray()

            Set Buffer = Nothing

        Case DIR_UP_RIGHT

            MapNpc(MapNum).Npc(MapNPCNum).Y = MapNpc(MapNum).Npc(MapNPCNum).Y - 1

            MapNpc(MapNum).Npc(MapNPCNum).X = MapNpc(MapNum).Npc(MapNPCNum).X + 1

            Set Buffer = New clsBuffer

            Buffer.WriteLong SNpcMove

            Buffer.WriteLong MapNPCNum

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir

            Buffer.WriteLong movement

            SendDataToMap MapNum, Buffer.ToArray()

            Set Buffer = Nothing

        Case DIR_DOWN_LEFT

            MapNpc(MapNum).Npc(MapNPCNum).Y = MapNpc(MapNum).Npc(MapNPCNum).Y + 1

            MapNpc(MapNum).Npc(MapNPCNum).X = MapNpc(MapNum).Npc(MapNPCNum).X - 1

            Set Buffer = New clsBuffer

            Buffer.WriteLong SNpcMove

            Buffer.WriteLong MapNPCNum

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir

            Buffer.WriteLong movement

            SendDataToMap MapNum, Buffer.ToArray()

            Set Buffer = Nothing

        Case DIR_DOWN_RIGHT

            MapNpc(MapNum).Npc(MapNPCNum).Y = MapNpc(MapNum).Npc(MapNPCNum).Y + 1

            MapNpc(MapNum).Npc(MapNPCNum).X = MapNpc(MapNum).Npc(MapNPCNum).X + 1

            Set Buffer = New clsBuffer

            Buffer.WriteLong SNpcMove

            Buffer.WriteLong MapNPCNum

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y

            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir

            Buffer.WriteLong movement

            SendDataToMap MapNum, Buffer.ToArray()

            Set Buffer = Nothing
        Case DIR_UP
            MapNpc(MapNum).Npc(MapNPCNum).Y = MapNpc(MapNum).Npc(MapNPCNum).Y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(MapNum).Npc(MapNPCNum).Y = MapNpc(MapNum).Npc(MapNPCNum).Y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(MapNum).Npc(MapNPCNum).X = MapNpc(MapNum).Npc(MapNPCNum).X - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(MapNum).Npc(MapNPCNum).X = MapNpc(MapNum).Npc(MapNPCNum).X + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).X
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select
    
    UpdateMapBlock MapNum, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, True

End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal Dir As Long, Optional Update As Byte = 1)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(MapNPCNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong MapNPCNum
    Buffer.WriteLong Dir
    Buffer.WriteByte Update
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = MapNum And Player(i).IsDead = 0 Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Public Sub CacheResources(ByVal MapNum As Long)
    Dim X As Long, Y As Long, Resource_Count As Long, ExtractorCount As Long
    Resource_Count = 0
    ExtractorCount = 0

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                'If Resource(Map(MapNum).Tile(X, Y).Data1).health > 0 Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                ResourceCache(MapNum).ResourceData(Resource_Count).X = X
                ResourceCache(MapNum).ResourceData(Resource_Count).Y = Y
                ResourceCache(MapNum).ResourceData(Resource_Count).ResourceNum = Map(MapNum).Tile(X, Y).data1
                If IsPlayerMap(MapNum) Then
                    Dim i As Long, n As Long
                    For i = 1 To UBound(ResourceFactory)
                        For n = 1 To UBound(ResourceFactory(i).Resource)
                            If ResourceFactory(i).Resource(n) = Map(MapNum).Tile(X, Y).data1 Then
                                ResourceCache(MapNum).ResourceData(Resource_Count).ResourceState = 1
                                Exit For
                            End If
                        Next n
                    Next i
                End If
                    
                If Not UZ Then
                    ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(Map(MapNum).Tile(X, Y).data1).health
                Else
                    If MapNum >= PlanetStart And MapNum <= PlanetStart + MAX_PLANET_BASE Then
                        Dim PlanetNum As Long
                        PlanetNum = GetPlanetNum(MapNum)
                        ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = ((NPCBase(Planets(PlanetNum).Level).HP * 6) / 100) * Resource(Map(MapNum).Tile(X, Y).data1).health
                    Else
                        ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(Map(MapNum).Tile(X, Y).data1).health
                    End If
                    If Resource(Map(MapNum).Tile(X, Y).data1).ResourceType = 4 Then ExtractorCount = ExtractorCount + 1
                End If
                'End If
            End If

        Next
    Next

    ResourceCache(MapNum).Resource_Count = Resource_Count
    ResourceCache(MapNum).ExtractorCount = ExtractorCount
End Sub

Sub PlayerSwitchBankSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(Index, oldSlot)
    OldValue = GetPlayerBankItemValue(Index, oldSlot)
    NewNum = GetPlayerBankItemNum(Index, newSlot)
    NewValue = GetPlayerBankItemValue(Index, newSlot)
    
    SetPlayerBankItemNum Index, newSlot, OldNum
    SetPlayerBankItemValue Index, newSlot, OldValue
    
    SetPlayerBankItemNum Index, oldSlot, NewNum
    SetPlayerBankItemValue Index, oldSlot, NewValue
        
    SendBank Index
End Sub

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(Index, oldSlot)
    OldValue = GetPlayerInvItemValue(Index, oldSlot)
    NewNum = GetPlayerInvItemNum(Index, newSlot)
    NewValue = GetPlayerInvItemValue(Index, newSlot)
    SetPlayerInvItemNum Index, newSlot, OldNum
    SetPlayerInvItemValue Index, newSlot, OldValue
    SetPlayerInvItemNum Index, oldSlot, NewNum
    SetPlayerInvItemValue Index, oldSlot, NewValue
    SendInventory Index
End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(Index, oldSlot)
    NewNum = GetPlayerSpell(Index, newSlot)
    SetPlayerSpell Index, oldSlot, NewNum
    SetPlayerSpell Index, newSlot, OldNum
    SendPlayerSpells Index
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(Index, GetPlayerEquipment(Index, EqSlot)) > 0 Then
        If Item(GetPlayerEquipment(Index, EqSlot)).Stackable > 0 Then
            GiveInvItem Index, GetPlayerEquipment(Index, EqSlot), 1
        Else
            GiveInvItem Index, GetPlayerEquipment(Index, EqSlot), 0
        End If
        PlayerMsg Index, printf("Desequipou %s", CheckGrammar(Item(GetPlayerEquipment(Index, EqSlot)).Name)), Yellow
        ' send the sound
        SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipment(Index, EqSlot)
        ' remove equipment
        SetPlayerEquipment Index, 0, EqSlot
        SendWornEquipment Index
        SendMapEquipment Index
        SendStats Index
        ' send vitals
        Call SendVital(Index, Vitals.HP)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
    Else
        PlayerMsg Index, printf("Inventário cheio."), brightred
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function rand(ByVal Low As Long, ByVal High As Long) As Long
    Dim RandomNumber As Long
    'Randomize RandomNumber
    RandomNumber = Int((High - Low + 1) * Rnd) + Low
    rand = Int(((High - Low + 1) / 100) * Val(RandomString(2))) + Low  'RandomNumber
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal Index As Long)
Dim PartyNum As Long, i As Long

    PartyNum = TempPlayer(Index).inParty
    If PartyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers PartyNum
        ' make sure there's more than 2 people
        If Party(PartyNum).MemberCount > 2 Then
        
            ' check if leader
            If Party(PartyNum).Leader = Index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(PartyNum).Member(i) > 0 And Party(PartyNum).Member(i) <> Index Then
                        Party(PartyNum).Leader = Party(PartyNum).Member(i)
                        PartyMsg PartyNum, GetPlayerName(i) & " é o lider do grupo.", brightblue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg PartyNum, GetPlayerName(Index) & " saiu do grupo.", brightred
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(PartyNum).Member(i) = Index Then
                        Party(PartyNum).Member(i) = 0
                        TempPlayer(Index).inParty = 0
                        TempPlayer(Index).partyInvite = 0
                        Exit For
                        End If
                Next
                ' recount party
                Party_CountMembers PartyNum
                ' set update to all
                SendPartyUpdate PartyNum
                ' send clear to player
                SendPartyUpdateTo Index
            Else
                ' not the leader, just leave
                PartyMsg PartyNum, GetPlayerName(Index) & " saiu do grupo.", brightred
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(PartyNum).Member(i) = Index Then
                        Party(PartyNum).Member(i) = 0
                        TempPlayer(Index).inParty = 0
                        TempPlayer(Index).partyInvite = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers PartyNum
                ' set update to all
                SendPartyUpdate PartyNum
                ' send clear to player
                SendPartyUpdateTo Index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers PartyNum
            ' only 2 people, disband
            PartyMsg PartyNum, "Grupo desfeito.", brightred
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                Index = Party(PartyNum).Member(i)
                ' player exist?
                If Index > 0 Then
                    ' remove them
                    TempPlayer(Index).partyInvite = 0
                    TempPlayer(Index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo Index
                End If
            Next
            ' clear out the party itself
            ClearParty PartyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal Index As Long, ByVal TargetPlayer As Long)
Dim PartyNum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(TargetPlayer) Or Not IsPlaying(TargetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(TargetPlayer).partyInvite > 0 Or TempPlayer(TargetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg Index, printf("Jogador ocupado."), brightred
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(TargetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg Index, printf("Jogador já está em um grupo."), brightred
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(Index).inParty > 0 Then
        PartyNum = TempPlayer(Index).inParty
        ' make sure we're the leader
        If Party(PartyNum).Leader = Index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(PartyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite TargetPlayer, Index
                    ' set the invite target
                    TempPlayer(TargetPlayer).partyInvite = Index
                    ' let them know
                    PlayerMsg Index, printf("Convite enviado."), Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg Index, printf("Grupo cheio."), brightred
            Exit Sub
        Else
            ' not the leader
            PlayerMsg Index, printf("Você não é o lider do grupo."), brightred
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite TargetPlayer, Index
        ' set the invite target
        TempPlayer(TargetPlayer).partyInvite = Index
        ' let them know
        PlayerMsg Index, printf("Convite enviado."), Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal Index As Long, ByVal TargetPlayer As Long)
Dim PartyNum As Long, i As Long

    ' check if already in a party
    If TempPlayer(Index).inParty > 0 Then
        ' get the partynumber
        PartyNum = TempPlayer(Index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(PartyNum).Member(i) = 0 Then
                'add to the party
                Party(PartyNum).Member(i) = TargetPlayer
                ' recount party
                Party_CountMembers PartyNum
                ' send update to all - including new player
                SendPartyUpdate PartyNum
                SendPartyVitals PartyNum, TargetPlayer
                ' let everyone know they've joined
                PartyMsg PartyNum, GetPlayerName(TargetPlayer) & " entrou no grupo.", Pink
                ' add them in
                TempPlayer(TargetPlayer).inParty = PartyNum
                'CheckSpecialEvent Index, TargetPlayer, PartyNum
                Exit Sub
            End If
        Next
        
        ' no empty slots - let them know
        PlayerMsg Index, printf("Grupo cheio."), brightred
        PlayerMsg TargetPlayer, printf("Grupo cheio."), brightred
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                PartyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(PartyNum).MemberCount = 2
        Party(PartyNum).Leader = Index
        Party(PartyNum).Member(1) = Index
        Party(PartyNum).Member(2) = TargetPlayer
        SendPartyUpdate PartyNum
        SendPartyVitals PartyNum, Index
        SendPartyVitals PartyNum, TargetPlayer
        ' let them know it's created
        PartyMsg PartyNum, "Party created.", brightgreen
        PartyMsg PartyNum, GetPlayerName(Index) & " entrou no grupo.", Pink
        PartyMsg PartyNum, GetPlayerName(TargetPlayer) & " entrou no grupo.", Pink
        ' clear the invitation
        TempPlayer(TargetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(Index).inParty = PartyNum
        TempPlayer(TargetPlayer).inParty = PartyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal Index As Long, ByVal TargetPlayer As Long)
    PlayerMsg Index, printf("%s negou a se juntar ao grupo.", GetPlayerName(TargetPlayer)), brightred
    PlayerMsg TargetPlayer, printf("Você negou á se juntar ao grupo."), brightred
    ' clear the invitation
    TempPlayer(TargetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal PartyNum As Long)
Dim i As Long, highIndex As Long, X As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(PartyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(PartyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For X = i To MAX_PARTY_MEMBERS - 1
                    Party(PartyNum).Member(X) = Party(PartyNum).Member(X + 1)
                    Party(PartyNum).Member(X + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(PartyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(PartyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers PartyNum
End Sub

Public Sub Party_ShareEvent(ByVal PartyNum As Long, ByVal EventNum As Long, ByVal MapNum As Long)
Dim i As Long, tmpIndex As Long
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(PartyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) = MapNum Then
                    ' give them their share
                    InitEvent tmpIndex, EventNum
                End If
            End If
        End If
    Next
End Sub

Public Sub Party_ShareExp(ByVal PartyNum As Long, ByVal Exp As Long, ByVal Index As Long, ByVal MapNum As Long)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long, LoseMemberCount As Byte

    ' check if it's worth sharing
    If Not Exp >= Party(PartyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP Index, Exp
        Exit Sub
    End If
    
    ' check members in outhers maps
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(PartyNum).Member(i)
        If tmpIndex > 0 Then
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) <> MapNum Then
                    LoseMemberCount = LoseMemberCount + 1
                End If
            End If
        End If
    Next i
    
    ' find out the equal share
    expShare = Exp \ (Party(PartyNum).MemberCount - LoseMemberCount)
    leftOver = Exp Mod (Party(PartyNum).MemberCount - LoseMemberCount)
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(PartyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) = MapNum Then
                    ' give them their share
                    GivePlayerEXP tmpIndex, expShare
                End If
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(PartyNum).Member(rand(1, Party(PartyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver
End Sub

Public Sub GivePlayerEXP(ByVal Index As Long, ByVal Exp As Long)
    Dim ExpFinal As Long
    'calculate bonus
    ExpFinal = Exp
    
    If Index = 0 Then Exit Sub
    
    'esoterica
    If Player(Index).EsoBonus > 0 Then
        ExpFinal = ExpFinal + ((Exp / 100) * (100 + Player(Index).EsoBonus))
    End If
    
    'vip
    If Player(Index).VIP >= 1 Then
        ExpFinal = ExpFinal + (Exp * VIPBonus(Index))
    End If
    
    If Options.ExpFactor > 1 Then
        ExpFinal = ExpFinal * Options.ExpFactor
    End If

    If ExpFinal > 0 Then
        If Player(Index).IsGod = 0 Then
            ' give the exp
            Call SetPlayerExp(Index, GetPlayerExp(Index) + ExpFinal)
        Else
            Player(Index).GodExp = Player(Index).GodExp + ExpFinal
        End If
        SendEXP Index
        If ExpFinal = Exp Then
            SendActionMsg GetPlayerMap(Index), "+" & Exp & " EXP", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        Else
            SendActionMsg GetPlayerMap(Index), "+" & Exp & " EXP (+" & (ExpFinal - Exp) & " bonus!)", Yellow, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
        If Not Player(Index).IsGod > 0 Then
            'SendActionMsg GetPlayerMap(Index), actionf("+%d Poder de luta", Int(ExpFinal * 0.047)), brightgreen, 1, (GetPlayerX(Index) * 32) + 32, (GetPlayerY(Index) * 32) + 32
            Player(Index).PDL = Player(Index).PDL + Int(ExpFinal * ExpToPDL)
            ' check if we've leveled
            CheckPlayerLevelUp Index
        Else
            CheckPlayerGodLevelUp Index
        End If
        
        
        'If Int(ExpFinal * ExpToPDL) > 0 Then SendStats Index
    End If
End Sub

Function GetLevelPDL(ByVal Level As Long) As Long
    If Level = 1 Then
        GetLevelPDL = 10
    Else
        GetLevelPDL = GetLevelPDL(Level - 1) + (Experience(Level) * ExpToPDL) + LevelUpBonus
    End If
End Function

Sub CreateProjectile(ByVal MapNum As Long, ByVal Attacker As Long, ByVal victim As Long, ByVal TargetType As Long, ByVal Graphic As Long, ByVal RotateSpeed As Long, Optional NPCAttack As Byte = 0)
Dim Rotate As Long
Dim Buffer As clsBuffer
    
    ' ****** Initial Rotation Value ******
    Select Case TargetType
        Case TARGET_TYPE_PLAYER
            Rotate = Engine_GetAngle(GetPlayerX(Attacker), GetPlayerY(Attacker), GetPlayerX(victim), GetPlayerY(victim))
        Case TARGET_TYPE_NPC
            Rotate = Engine_GetAngle(GetPlayerX(Attacker), GetPlayerY(Attacker), MapNpc(MapNum).Npc(victim).X, MapNpc(MapNum).Npc(victim).Y)
    End Select
    
    If NPCAttack = 1 Then
        Rotate = Engine_GetAngle(MapNpc(MapNum).Npc(victim).X, MapNpc(MapNum).Npc(victim).Y, GetPlayerX(Attacker), GetPlayerY(Attacker))
    End If

    ' ****** Set Player Direction Based On Angle ******
    If NPCAttack = 0 Then
        If Rotate >= 315 And Rotate <= 360 Then
            Call SetPlayerDir(Attacker, DIR_UP)
        ElseIf Rotate >= 0 And Rotate <= 45 Then
            Call SetPlayerDir(Attacker, DIR_UP)
        ElseIf Rotate >= 225 And Rotate <= 315 Then
            Call SetPlayerDir(Attacker, DIR_LEFT)
        ElseIf Rotate >= 135 And Rotate <= 225 Then
            Call SetPlayerDir(Attacker, DIR_DOWN)
        ElseIf Rotate >= 45 And Rotate <= 135 Then
            Call SetPlayerDir(Attacker, DIR_RIGHT)
        End If

        Set Buffer = New clsBuffer
        Buffer.WriteLong SPlayerDir
        Buffer.WriteLong Attacker
        Buffer.WriteLong GetPlayerDir(Attacker)
        Call SendDataToMap(MapNum, Buffer.ToArray())
        Set Buffer = Nothing
    End If

    Call SendProjectile(MapNum, Attacker, victim, TargetType, Graphic, Rotate, RotateSpeed, NPCAttack)
End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal targetX As Integer, ByVal targetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = targetY Then
        'Check for going right (90 degrees)
        If CenterX < targetX Then
            Engine_GetAngle = 90
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
        
        'Exit the function
        Exit Function
    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = targetX Then
        'Check for going up (360 degrees)
        If CenterY > targetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function
    End If

    'Calculate Side C
    SideC = Sqr(Abs(targetX - CenterX) ^ 2 + Abs(targetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(targetX - CenterX) ^ 2 + targetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If targetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function
    Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

    Exit Function
End Function

' *****************
' ** Event Logic **
' *****************
Private Function IsForwardingEvent(ByVal EType As EventType) As Boolean
    Select Case EType
        Case Evt_Menu, Evt_Message
            IsForwardingEvent = False
        Case Else
            IsForwardingEvent = True
    End Select
End Function

Public Sub InitEvent(ByVal Index As Long, ByVal EventIndex As Long)
    If TempPlayer(Index).CurrentEvent > 0 And TempPlayer(Index).CurrentEvent <= MAX_EVENTS Then Exit Sub
    If Events(EventIndex).chkVariable > 0 Then
        If Not CheckComparisonOperator(Player(Index).Variables(Events(EventIndex).VariableIndex), Events(EventIndex).VariableCondition, Events(EventIndex).VariableCompare) = True Then
            Exit Sub
        End If
    End If
    
    If Events(EventIndex).chkSwitch > 0 Then
        If Not Player(Index).Switches(Events(EventIndex).SwitchIndex + 1) = Events(EventIndex).SwitchCompare Then
            Exit Sub
        End If
    End If
    
    If Events(EventIndex).chkHasItem > 0 Then
        If HasItem(Index, Events(EventIndex).HasItemIndex) = 0 Then
            Exit Sub
        End If
    End If
    
    TempPlayer(Index).CurrentEvent = EventIndex
    Call DoEventLogic(Index, 1)
End Sub

Public Function CheckComparisonOperator(ByVal numOne As Long, ByVal numTwo As Long, ByVal opr As ComparisonOperator) As Boolean
    CheckComparisonOperator = False
    Select Case opr
        Case GEQUAL
            If numOne >= numTwo Then CheckComparisonOperator = True
        Case LEQUAL
            If numOne <= numTwo Then CheckComparisonOperator = True
        Case GREATER
            If numOne > numTwo Then CheckComparisonOperator = True
        Case LESS
            If numOne < numTwo Then CheckComparisonOperator = True
        Case EQUAL
            If numOne = numTwo Then CheckComparisonOperator = True
        Case NOTEQUAL
            If Not (numOne = numTwo) Then CheckComparisonOperator = True
    End Select
End Function

Public Sub DoEventLogic(ByVal Index As Long, ByVal Opt As Long)
Dim X As Long, Y As Long, i As Long, Buffer As clsBuffer, ScriptContinue As Boolean
    
    If TempPlayer(Index).CurrentEvent <= 0 Or TempPlayer(Index).CurrentEvent > MAX_EVENTS Then GoTo EventQuit
    If Not (Events(TempPlayer(Index).CurrentEvent).HasSubEvents) Then GoTo EventQuit
    If Opt <= 0 Or Opt > UBound(Events(TempPlayer(Index).CurrentEvent).SubEvents) Then GoTo EventQuit
    ScriptContinue = True
    
        With Events(TempPlayer(Index).CurrentEvent).SubEvents(Opt)
            Select Case .Type
                Case Evt_Quit
                    GoTo EventQuit
                Case Evt_OpenShop
                    Call SendOpenShop(Index, .Data(1))
                    TempPlayer(Index).InShop = .Data(1)
                    GoTo EventQuit
                Case Evt_OpenBank
                    SendBank Index
                    TempPlayer(Index).InBank = True
                    GoTo EventQuit
                Case Evt_GiveItem
                    If .Data(1) > 0 And .Data(1) <= MAX_ITEMS Then
                        Select Case .Data(3)
                            Case 0: If .Data(1) > 0 Then Call TakeInvItem(Index, .Data(1), .Data(2))
                            'Case 1: Call SetPlayerItems(Index, .Data(1), .Data(2))
                            Case 2: Call GiveInvItem(Index, .Data(1), .Data(2), True)
                        End Select
                    End If
                    SendInventory Index
                Case Evt_ChangeLevel
                    Select Case .Data(2)
                        Case 0: Call SetPlayerLevel(Index, .Data(1))
                        Case 1: Call SetPlayerLevel(Index, GetPlayerLevel(Index) + .Data(1))
                        Case 2: Call SetPlayerLevel(Index, GetPlayerLevel(Index) - .Data(1))
                    End Select
                    SendPlayerData Index
                Case Evt_PlayAnimation
                    X = .Data(2)
                    Y = .Data(3)
                    If X < 0 Then X = GetPlayerX(Index)
                    If Y < 0 Then Y = GetPlayerY(Index)
                    If X >= 0 And Y >= 0 And X <= Map(GetPlayerMap(Index)).MaxX And Y <= Map(GetPlayerMap(Index)).MaxY Then Call SendAnimation(GetPlayerMap(Index), .Data(1), X, Y, GetPlayerDir(Index))
                Case Evt_Warp
                    If .Data(1) >= 1 And .Data(1) <= MAX_MAPS Then
                        If .Data(2) >= 0 And .Data(3) >= 0 And .Data(2) <= Map(.Data(1)).MaxX And .Data(3) <= Map(.Data(1)).MaxY Then Call PlayerWarp(Index, .Data(1), .Data(2), .Data(3))
                    End If
                Case Evt_GOTO
                    Call DoEventLogic(Index, .Data(1))
                    Exit Sub
                Case Evt_Switch
                    Player(Index).Switches(.Data(1)) = .Data(2)
                Case Evt_Variable
                    Select Case .Data(2)
                        Case 0: Player(Index).Variables(.Data(1)) = .Data(3)
                        Case 1: Player(Index).Variables(.Data(1)) = Player(Index).Variables(.Data(1)) + .Data(3)
                        Case 2: Player(Index).Variables(.Data(1)) = Player(Index).Variables(.Data(1)) - .Data(3)
                        Case 3: Player(Index).Variables(.Data(1)) = Random(.Data(3), .Data(4))
                    End Select
                Case Evt_AddText
                    Select Case .Data(2)
                        Case 0: PlayerMsg Index, Trim$(.Text(1)), .Data(1)
                        Case 1: MapMsg GetPlayerMap(Index), Trim$(.Text(1)), .Data(1)
                        Case 2: GlobalMsg Trim$(.Text(1)), .Data(1)
                    End Select
                Case Evt_Chatbubble
                    Select Case .Data(1)
                        Case 0: SendChatBubble GetPlayerMap(Index), Index, TARGET_TYPE_PLAYER, Trim$(.Text(1)), DarkBrown
                        Case 1: SendChatBubble GetPlayerMap(Index), .Data(2), TARGET_TYPE_NPC, Trim$(.Text(1)), DarkBrown
                    End Select
                Case Evt_Branch
                    Select Case .Data(1)
                        Case 0
                            If CheckComparisonOperator(Player(Index).Variables(.Data(6)), .Data(2), .Data(5)) Then
                                If .Data(3) <> Opt Then Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                If .Data(4) <> Opt Then Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                        Case 1
                            If Player(Index).Switches(.Data(5)) = .Data(2) Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                        Case 2
                            If HasItems(Index, .Data(2)) >= .Data(5) Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                        Case 3
                            If GetPlayerClass(Index) = .Data(2) Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                        Case 4
                            If HasSpell(Index, .Data(2)) Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                        Case 5
                            If CheckComparisonOperator(GetPlayerLevel(Index), .Data(2), .Data(5)) Then
                                Call DoEventLogic(Index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(Index, .Data(4))
                                Exit Sub
                            End If
                    End Select
                Case Evt_ChangeSkill
                    If .Data(2) = 0 Then
                        If FindOpenSpellSlot(Index) > 0 Then
                            If HasSpell(Index, .Data(1)) = False Then
                                SetPlayerSpell Index, FindOpenSpellSlot(Index), .Data(1)
                            End If
                        End If
                    Else
                        If HasSpell(Index, .Data(1)) = True Then
                            For i = 1 To MAX_PLAYER_SPELLS
                                If Player(Index).Spell(i) = .Data(1) Then
                                    SetPlayerSpell Index, i, 0
                                End If
                            Next
                        End If
                    End If
                    SendPlayerSpells Index
                Case Evt_ChangeSprite
                    SetPlayerSprite Index, .Data(1)
                    SendPlayerData Index
                Case Evt_ChangePK
                    SetPlayerPK Index, .Data(1)
                    SendPlayerData Index
                Case Evt_SpawnNPC
                    If .Data(1) > 0 Then
                        If MapNpc(GetPlayerMap(Index)).Npc(.Data(1)).Target = 0 Then
                            SpawnNpc .Data(1), GetPlayerMap(Index), True
                        End If
                    End If
                Case Evt_ChangeClass
                    SetPlayerClass Index, .Data(1)
                    SendPlayerData Index
                Case Evt_ChangeSex
                    Player(Index).Sex = .Data(1)
                    SendPlayerData Index
                Case Evt_ChangeExp
                    Select Case .Data(2)
                        Case 0: Call SetPlayerExp(Index, .Data(1))
                        Case 1: Call GivePlayerEXP(Index, .Data(1))
                        Case 2: Call SetPlayerExp(Index, GetPlayerExp(Index) - .Data(1))
                    End Select
                    CheckPlayerLevelUp Index
                    SendEXP Index
                Case Evt_SpecialEffect
                    Select Case .Data(1)
                        Case 0: SendSpecialEffect Index, EFFECT_TYPE_FADEOUT
                        Case 1: SendSpecialEffect Index, EFFECT_TYPE_FADEIN
                        Case 2: SendSpecialEffect Index, EFFECT_TYPE_FLASH
                        Case 3: SendSpecialEffect Index, EFFECT_TYPE_FOG, .Data(2), .Data(3), .Data(4)
                        Case 4: SendSpecialEffect Index, EFFECT_TYPE_WEATHER, .Data(2), .Data(3)
                        Case 5: SendSpecialEffect Index, EFFECT_TYPE_TINT, .Data(2), .Data(3), .Data(4), .Data(5)
                    End Select
                Case Evt_PlaySound
                    Set Buffer = New clsBuffer
                        Buffer.WriteLong SPlaySound
                        Buffer.WriteString Trim$(.Text(1))
                        SendDataTo Index, Buffer.ToArray
                    Set Buffer = Nothing
                Case Evt_PlayBGM
                    Set Buffer = New clsBuffer
                        Buffer.WriteLong SPlayBGM
                        Buffer.WriteString Trim$(.Text(1))
                        SendDataTo Index, Buffer.ToArray
                    Set Buffer = Nothing
                Case Evt_StopSound
                    Set Buffer = New clsBuffer
                        Buffer.WriteLong SStopSound
                        SendDataTo Index, Buffer.ToArray
                    Set Buffer = Nothing
                Case Evt_FadeoutBGM
                    Set Buffer = New clsBuffer
                        Buffer.WriteLong SFadeoutBGM
                        SendDataTo Index, Buffer.ToArray
                    Set Buffer = Nothing
                Case Evt_SetAccess
                    SetPlayerAccess Index, .Data(1)
                    SendPlayerData Index
                Case Evt_CustomScript
                    ScriptContinue = CustomScript(Index, .Data(1))
                Case Evt_Quest
                    X = .Data(1)
                    Player(Index).QuestState(X).State = .Data(2)
                    If .Data(2) <> 1 Then
                        Player(Index).QuestState(X).Date = Now
                    End If
                    Call SendPlayerQuest(Index, X)
                Case Evt_OpenEvent
                    X = .Data(1)
                    Player(Index).EventOpen(X) = YES
                    Player(Index).EventOpen(TempPlayer(Index).CurrentEvent) = NO
                    TempPlayer(Index).CurrentEvent = 0
                    Call InitEvent(Index, X)
                    Exit Sub
                    
                    'SendMapKey Index, YES, X
                    'Y = .Data(2)
                    'If .Data(3) = 0 Then
                    '    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_EVENT And Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(X, Y).Data1) = NO Then
                    '        Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(X, Y).Data1) = YES
                    '        Select Case .Data(4)
                    '            Case 0: SendMapKey Index, YES, Map(GetPlayerMap(Index)).Tile(X, Y).Data1
                    '            Case 1: SendMapKeyToMap GetPlayerMap(Index), YES, Map(GetPlayerMap(Index)).Tile(X, Y).Data1
                    '            Case 2: SendMapKeyToAll YES, Map(GetPlayerMap(Index)).Tile(X, Y).Data1
                    '        End Select
                    '    End If
                    'Else
                    '    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_EVENT And Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(X, Y).Data1) = YES Then
                    '        Player(Index).EventOpen(Map(GetPlayerMap(Index)).Tile(X, Y).Data1) = NO
                    '        Select Case .Data(4)
                    '            Case 0: SendMapKey Index, NO, Map(GetPlayerMap(Index)).Tile(X, Y).Data1
                    '            Case 1: SendMapKeyToMap GetPlayerMap(Index), NO, Map(GetPlayerMap(Index)).Tile(X, Y).Data1
                    '            Case 2: SendMapKeyToAll NO, Map(GetPlayerMap(Index)).Tile(X, Y).Data1
                    '        End Select
                    '    End If
                    'End If
            End Select
        End With
    
    'Make sure this is last
    If IsForwardingEvent(Events(TempPlayer(Index).CurrentEvent).SubEvents(Opt).Type) And ScriptContinue Then
        Call DoEventLogic(Index, Opt + 1)
    Else
        If Events(TempPlayer(Index).CurrentEvent).SubEvents(Opt).Type = Evt_CustomScript And Not ScriptContinue Then GoTo EventQuit
        Call Events_SendEventUpdate(Index, TempPlayer(Index).CurrentEvent, Opt)
    End If
    
    
    
Exit Sub
EventQuit:
    TempPlayer(Index).CurrentEvent = -1
    Events_SendEventQuit Index
    Exit Sub
End Sub
Sub CheckShenlong(ByVal Index As Long, ByVal MapNum As Long, X As Long, Y As Long)
    Dim i As Long
    Dim DragonballCheck(1 To 7) As Boolean

    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num > 0 Then
            If Item(MapItem(MapNum, i).Num).Type = ITEM_TYPE_DRAGONBALL Then
                DragonballCheck(Item(MapItem(MapNum, i).Num).Dragonball) = True
            End If
        End If
    Next i
    
    For i = 1 To 7
        If DragonballCheck(i) = False Then Exit Sub
    Next i
    
    If ShenlongTick + 600000 > GetTickCount Then
        PlayerMsg Index, "O shenlong ainda está em período de intervalo, aguarde 10 minutos", Yellow
        Exit Sub
    End If
    
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num > 0 Then
            If Item(MapItem(MapNum, i).Num).Type = ITEM_TYPE_DRAGONBALL Then
                Call ClearMapItem(i, MapNum)
                Call SpawnItemSlot(i, 0, 0, MapNum, 0, 0)
            End If
        End If
    Next i

    SendShenlong MapNum, 1, 1, , X, Y
    ShenlongTick = GetTickCount
    ShenlongActive = 1
    ShenlongMap = MapNum
    ShenlongX = X
    ShenlongY = Y
    ShenlongOwner = GetPlayerName(Index)
End Sub

Sub DoWish(ByVal Index As Long, ByVal Msg As String)
    Dim Wished As Boolean
    If GetPlayerMap(Index) <> ShenlongMap Then Exit Sub
    If GetPlayerName(Index) <> ShenlongOwner Then Exit Sub

    Dim i As Long
    For i = 1 To UBound(Wish)
        If Wish(i).Phrase = LCase(Msg) Then
            If Wish(i).Type = 0 Then
                Call InitEvent(Index, Wish(i).Event)
            Else
                Call GiveInvItem(Index, Wish(i).Item, Wish(i).ItemVal)
                Call PlayerMsg(Index, "Você recebeu: " & Wish(i).ItemVal & " " & Trim$(Item(Wish(i).Item).Name), brightgreen)
            End If
            Wished = True
            Exit For
        End If
    Next i
    
    If LCase(Msg) = "eu desejo me tornar um deus" Then
        'SSJ RED
        If TempPlayer(Index).inParty > 0 Then
            Dim tmpIndex As Long, PartyNum As Long, IsAllOk As Boolean
            PartyNum = TempPlayer(Index).inParty
            IsAllOk = True
            For i = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party(PartyNum).Member(i)
                ' existing member?
                If tmpIndex > 0 Then
                    ' playing?
                    If Not IsConnected(tmpIndex) Or IsPlaying(tmpIndex) Then
                        IsAllOk = False
                    Else
                        If GetPlayerMap(tmpIndex) = GetPlayerMap(Index) Then
                            IsAllOk = False
                        End If
                    End If
                Else
                    IsAllOk = False
                End If
            Next
        End If
        If HasItem(Index, 108) = 0 Then
            PlayerMsg Index, "Você deve ser um General do exército sayajin para fazer este desejo", brightred
        End If
        If Not IsAllOk Then
            PlayerMsg Index, "Você deve estar em um grupo cheio de Sayajins para fazer este pedido!", brightred
        Else
            If Player(Index).IsGod = 0 Then
                If GetPlayerLevel(Index) >= MAX_LEVELS - 1 Then
                    
                    i = FindOpenSpellSlot(Index)
                    
                    If i > 0 Then
                        SetPlayerSpell Index, i, 90
                        TransPlayer Index, 90
                        SendAnimation GetPlayerMap(Index), Spell(90).SpellAnim, GetPlayerX(Index), GetPlayerY(Index), GetPlayerDir(Index), TargetType.TARGET_TYPE_PLAYER, Index
                        SendPlayerSpells Index
                    Else
                        PlayerMsg Index, "Você não tem um espaço livre para uma nova habilidade!", brightred
                        Exit Sub
                    End If
                    
                    Player(Index).IsGod = 1
                    Player(Index).GodLevel = 1
                    Player(Index).GodExp = 0
                    
                    TakeInvItem Index, 108, 1
                    GiveInvItem Index, 109, 1
                    
                    PlayerMsg Index, "Parabéns! Você se tornou um Super Sayajin Deus!", brightred
                    Wished = True
                    SendPlayerData Index
                Else
                    PlayerMsg Index, "Você deve estar no level máximo para se tornar um Sayajin Deus!", brightred
                End If
            Else
                PlayerMsg Index, "Você já é um Sayajin Deus!", brightred
            End If
        End If
    End If
    
    If Wished = True Then
        SendShenlong GetPlayerMap(Index), 0, 1, , ShenlongX, ShenlongY
        ShenlongActive = 0
    End If
End Sub

Sub DropItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional Owner As String)
    Dim i As Long
        i = FindOpenMapItemSlot(MapNum)
            
        MapItem(MapNum, i).Num = ItemNum
        MapItem(MapNum, i).X = X
        MapItem(MapNum, i).Y = Y
        MapItem(MapNum, i).PlayerName = Owner
        MapItem(MapNum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
        MapItem(MapNum, i).canDespawn = True
        MapItem(MapNum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            
        Call SpawnItemSlot(i, ItemNum, ItemVal, MapNum, X, Y, Owner)
End Sub

Function GravityValue(ByVal Index As Long, ByVal Gravity As Long) As Long
    If Gravity < 10 Then Gravity = 10
    If Gravity > 1500 Then Gravity = 1500
    If Gravity > Int(10 + (GetPlayerLevel(Index) / MAX_LEVELS) * 1500) Then Gravity = Int(10 + (GetPlayerLevel(Index) / MAX_LEVELS) * 1500)
    GravityValue = Int(((Gravity / 1500) * 50000) + 5000)
End Function

Public Function Fat(ByVal Number As Long) As Long
    If Number = 1 Then
        Fat = 1
    Else
        Fat = Number * Fat(Number - 1)
    End If
End Function

Public Function RandomString( _
    ByVal Length As Long, _
    Optional charset As String = "0123456789" _
    ) As String
    Dim chars() As Byte, Value() As Byte, chrUprBnd As Long, i As Long
    If Length > 0& Then
        Randomize
        chars = charset
        chrUprBnd = Len(charset) - 1&
        Length = (Length * 2&) - 1&
        ReDim Value(Length) As Byte
        For i = 0& To Length Step 2&
            Value(i) = chars(CLng(chrUprBnd * Rnd) * 2&)
        Next
    End If
    RandomString = Value
End Function

Function InPartyWith(ByVal Index As Long, ByVal Name As String) As Boolean
    Dim i As Long
    Dim n As Long
    n = FindPlayer(Name)
    If n > 0 Then
        If TempPlayer(n).inParty = TempPlayer(Index).inParty Then
            InPartyWith = True
            Exit Function
        End If
    End If
End Function

