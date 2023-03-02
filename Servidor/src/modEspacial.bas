Attribute VB_Name = "modEspacial"
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Const UZ As Boolean = False
Public Const ViagemMap As Long = 1
Public Const ViagemX As Long = 50
Public Const ViagemY As Long = 50
Public Const VirgoMap As Long = 40

Public MAX_PLANETS As Long
Public Const MAX_PLANET_BASE As Long = 300
Public Const MAX_PLAYER_PLANETS As Long = 150
Public MAX_PLANET_CUSTOM As Long
Public Const PlanetStart As Long = 75
Public Const MapRespawn As Long = 4
Public Const RespawnTime As Long = 30000
Public Const ESOTERICAAUTO As Long = 63
Public Const TesouroMap As Long = 21
Public Const SpawnAnim As Long = 54
Public Const CUSTOM_PLANETS_RESPAWN As Long = 3600000
Public Const NPCS_BOSSES_START As Long = 200

Public PlayerMapIndex() As Long

Public LastCustomPlanetsRespawn As Long
Public PlanetStarted As Boolean
Public TesouroStarted As Boolean

Public Planets() As PlanetRec
Public PlayerPlanet() As PlayerPlanetRec
Public PlanetLocations() As Boolean
Public ResourceFactory(1 To 3) As ResourceFactoryRec
Public NpcFactory(1 To 3) As NpcFactoryRec
Public MapFactory() As PlanetConfig

Public MatchData() As MatchDataRec

Public PlanetTile() As Byte

Enum TileConfigEnum
    Solo = 1
    Detail
    Block
    
    TileConfigCount
End Enum

Type MatchDataRec
    Points As Long
    TotalNpcs As Long
    WaveNum As Long
    WaveTick As Long
    SpawnTick As Long
    Active As Byte
    Planet As Long
    Indexes() As Long
    Winner As Long
    HighLevel As Long
    Stars As Long
    EliteWave As Boolean
    PriceBonus As Long
End Type

Type NpcFactoryRec
    Npc() As Long
    Elites() As Long
End Type

Type ResourceFactoryRec
    Resource() As Long
End Type

Type MoonRec
    Size As Long
    ColorR As Long
    ColorG As Long
    ColorB As Long
    Pic As Long
    Speed As Long
End Type

Type TileConfig
    X As Long
    Y As Long
    Tileset As Long
    Layer As Long
End Type

Type PlanetConfig
    Tile(1 To 3) As TileConfig
End Type

Type PlanetRec
    Name As String * NAME_LENGTH
    Map As Long
    Owner As String * NAME_LENGTH
    State As Byte
    PointsToConquest As Long
    WaveDuration As Long
    WaveCooldown As Long
    
    EspeciariaAmarela As Byte
    EspeciariaVermelha As Byte
    EspeciariaAzul As Byte
    
    Level As Long
    Especie As Byte
    Habitantes As Long
    Gravidade As Long
    Atmosfera As Long
    Preco As Long
    
    Pic As Long
    X As Long
    Y As Long
    Size As Long
    ColorR As Byte
    ColorG As Byte
    ColorB As Byte
    TileConfig As PlanetConfig
    MoonData As MoonRec
    TimeToExplode As Long
    Type As Byte
End Type

Type ConstructionRec
    X As Long
    Y As Long
    ResourceNum As Long
End Type

Type SaibamanRec
    Working As Byte
    X As Long
    Y As Long
    TaskInit As String * 255
    TaskType As Byte
    TaskResult As Long
    Accelerate As Long
End Type

Type ExtratorRec
    TaskInit As String * 255
    X As Long
    Y As Long
    Used As Byte
    Acc As Long
    AccStart As String * 255
End Type

Type SementeRec
    Quant As Long
    Start As String * 255
    Fila As Long
End Type

Type PlayerPlanetRec
    PlanetMap As MapRec
    LastLogin As String * 255
    TotalSaibamans As Byte
    Saibaman(1 To 5) As SaibamanRec
    Extrator(1 To 20) As ExtratorRec
    Sementes(1 To 5) As SementeRec
    Soldados(1 To 5) As SementeRec
    SoldadosAcc As Long
    SoldadosStart As Long
    SementesAcc As Long
    SementesStart As Long
    PlanetData As PlanetRec
End Type

Sub StartPlanets()
    If UZ Then
        ReDim Planets(1 To MAX_PLANETS + 1)
        ReDim PlanetTile(1 To MAX_PLANETS + 1)
        ReDim PlanetLocations(0 To Map(ViagemMap).MaxX, 0 To Map(ViagemMap).MaxY)
        ReDim MatchData(1 To 1)
        ReDim PlayerMapIndex(1 To MAX_MAPS)
        ReDim PlanetInService(1 To MAX_PLANETS + 1)
        
        Call LoadMapFactory
        
        Dim i As Long
        For i = 1 To MAX_PLANET_BASE + 1
            Call CreatePlanet(i)
        Next i
        
        LoadPlayersPlanets
        CreateAllPlanetMaps
        PlanetStarted = True
        frmServer.Caption = "Loading..."
    End If
End Sub

Function GetMaxPlayerPlanets() As Long
    Dim n As Long
    n = 1
    Do While FileExist(App.path & "\data\player planets\planet" & n & ".bin", True)
        n = n + 1
    Loop
    n = n - 1
    GetMaxPlayerPlanets = n
End Function

Sub LoadPlayersPlanets()
    Dim n As Long
    n = GetMaxPlayerPlanets
    
    If n = 0 Then
        ReDim PlayerPlanet(1 To 1)
    Else
        ReDim PlayerPlanet(1 To n)
        
        Dim i As Long
        For i = 1 To n
            LoadPlayerPlanet i
            
            Dim MapNum As Long
            MapNum = PlanetStart + MAX_PLANET_BASE + i
            
            LoadPlayerPlanetMap MapNum, PlayerPlanet(i).PlanetMap, i
        Next i
    End If
End Sub

Sub LoadPlayerPlanet(ByVal PlanetNum As Long)
    Dim F As Long
    Dim i As Long
    Dim filename As String
    filename = App.path & "\data\player planets\planet" & PlanetNum & ".bin"
    F = FreeFile
    Open filename For Binary As #F
        Get #F, , PlayerPlanet(PlanetNum).PlanetData
        Get #F, , PlayerPlanet(PlanetNum).LastLogin
        Get #F, , PlayerPlanet(PlanetNum).TotalSaibamans
        For i = 1 To PlayerPlanet(PlanetNum).TotalSaibamans
            Get #F, , PlayerPlanet(PlanetNum).Saibaman(i)
        Next i
        For i = 1 To 20
            Get #F, , PlayerPlanet(PlanetNum).Extrator(i)
        Next i
        For i = 1 To 5
            Get #F, , PlayerPlanet(PlanetNum).Sementes(i)
            Get #F, , PlayerPlanet(PlanetNum).Soldados(i)
        Next i
        Get #F, , PlayerPlanet(PlanetNum).SoldadosStart
        Get #F, , PlayerPlanet(PlanetNum).SoldadosAcc
        Get #F, , PlayerPlanet(PlanetNum).SementesStart
        Get #F, , PlayerPlanet(PlanetNum).SementesAcc
    Close #F
    
    filename = App.path & "\data\player planets\maps\planetmap" & PlanetNum & ".bin"
    DeprecatedLoadMap filename, PlayerPlanet(PlanetNum).PlanetMap
    Reposition PlanetNum
End Sub

Sub SavePlayerPlanet(ByVal PlanetNum As Long)
    Dim filename As String
    Dim F As Long
    Dim i As Long
    filename = App.path & "\data\player planets\planet" & PlanetNum & ".bin"
    F = FreeFile
    If Not UZ Then Exit Sub
    Open filename For Binary As #F
        Put #F, , PlayerPlanet(PlanetNum).PlanetData
        Put #F, , PlayerPlanet(PlanetNum).LastLogin
        Put #F, , PlayerPlanet(PlanetNum).TotalSaibamans
        For i = 1 To PlayerPlanet(PlanetNum).TotalSaibamans
            Put #F, , PlayerPlanet(PlanetNum).Saibaman(i)
        Next i
        For i = 1 To 20
            Put #F, , PlayerPlanet(PlanetNum).Extrator(i)
        Next i
        For i = 1 To 5
            Put #F, , PlayerPlanet(PlanetNum).Sementes(i)
            Put #F, , PlayerPlanet(PlanetNum).Soldados(i)
        Next i
        Put #F, , PlayerPlanet(PlanetNum).SoldadosStart
        Put #F, , PlayerPlanet(PlanetNum).SoldadosAcc
        Put #F, , PlayerPlanet(PlanetNum).SementesStart
        Put #F, , PlayerPlanet(PlanetNum).SementesAcc
    Close #F
    
    filename = App.path & "\data\player planets\maps\planetmap" & PlanetNum & ".bin"
    SaveMapData PlayerPlanet(PlanetNum).PlanetMap, filename
End Sub

Sub Reposition(ByVal PlanetNum As Long)
    Dim X As Long, Y As Long
    
Position:
    X = rand(1, Map(VirgoMap).MaxX)
    Y = rand(1, Map(VirgoMap).MaxY)
    
    Dim i As Long
    For i = 1 To UBound(PlayerPlanet)
        If PlayerPlanet(i).PlanetData.X = X And PlayerPlanet(i).PlanetData.Y = Y Then GoTo Position
    Next i
End Sub

Sub LoadMapFactory()
    Dim i As Long
    i = 1
    Do While FileExist(App.path & "\data\mapfactory\" & i & ".ini", True)
        i = i + 1
    Loop
    i = i - 1
    
    ReDim MapFactory(1 To i)
    
    LoadMapFactories
    LoadResourceFactory
    LoadNpcFactory
    
End Sub

Sub LoadResourceFactory()
    Dim i As Long, n As Long
    For i = 1 To 3
        Dim tmpResource() As String
        Dim Text As String
        Text = GetVar(App.path & "\data\mapfactory\resources.ini", "Resources", STR(i))
        If Len(Trim(Text)) > 0 Then
            tmpResource = Split(Text, ",")
            ReDim ResourceFactory(i).Resource(1 To UBound(tmpResource) + 1)
            For n = 1 To UBound(tmpResource) + 1
                ResourceFactory(i).Resource(n) = Val(tmpResource(n - 1))
            Next n
        Else
            ReDim ResourceFactory(i).Resource(1 To 1)
            ResourceFactory(i).Resource(1) = 0
        End If
    Next i
End Sub

Sub LoadNpcFactory()
    Dim i As Long, n As Long
    For i = 1 To 3
        Dim tmpNpc() As String
        Dim Text As String
        Text = GetVar(App.path & "\data\species.ini", "SPECIES", STR(i))
        If Len(Trim(Text)) > 0 Then
            tmpNpc = Split(Text, ",")
            ReDim NpcFactory(i).Npc(1 To UBound(tmpNpc) + 1)
            For n = 1 To UBound(tmpNpc) + 1
                NpcFactory(i).Npc(n) = Val(tmpNpc(n - 1))
            Next n
        Else
            ReDim NpcFactory(i).Npc(1 To 1)
            NpcFactory(i).Npc(1) = 0
        End If
        
        Text = GetVar(App.path & "\data\species.ini", "ELITES", STR(i))
        If Len(Trim(Text)) > 0 Then
            tmpNpc = Split(Text, ",")
            ReDim NpcFactory(i).Elites(1 To UBound(tmpNpc) + 1)
            For n = 1 To UBound(tmpNpc) + 1
                NpcFactory(i).Elites(n) = Val(tmpNpc(n - 1))
            Next n
        Else
            ReDim NpcFactory(i).Elites(1 To 1)
            NpcFactory(i).Elites(1) = 0
        End If
    Next i
End Sub

Sub LoadMapFactories()
    Dim i As Long, n As Long
    Dim filename As String
    For i = 1 To UBound(MapFactory)
        filename = App.path & "\data\mapfactory\" & i & ".ini"
        For n = 1 To TileConfigEnum.TileConfigCount - 1
            MapFactory(i).Tile(n).Tileset = Val(GetVar(filename, STR(n), "Tileset"))
            MapFactory(i).Tile(n).X = Val(GetVar(filename, STR(n), "X"))
            MapFactory(i).Tile(n).Y = Val(GetVar(filename, STR(n), "Y"))
            MapFactory(i).Tile(n).Layer = Val(GetVar(filename, STR(n), "Layer"))
        Next n
    Next i
End Sub

Sub Viajar(ByVal Index As Long, MapNum As Long, Optional X As Long = -1, Optional Y As Long = -1)
    If UZ Then
        RemoveTemporaryItems Index
        If TempPlayer(Index).Trans > 0 Then
            TransPlayer Index, 0
        End If
        Call SetPlayerSprite(Index, GetPlayerNave(Index))
        Dim i As Long
        TempPlayer(Index).Speed = 1
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) > 0 Then
                If Item(GetPlayerInvItemNum(Index, i)).Type = ItemType.ITEM_TYPE_NAVE Then
                    If TempPlayer(Index).Speed < Item(Player(Index).Inv(i).Num).data2 Then
                        TempPlayer(Index).Speed = Item(GetPlayerInvItemNum(Index, i)).data2
                    End If
                End If
            End If
        Next i
        If Not IsPlayerMap(GetPlayerMap(Index)) Then
            If GetPlayerMap(Index) < PlanetStart Or GetPlayerMap(Index) >= PlanetStart + MAX_PLANETS + 1 Then
                If X = -1 Then X = ViagemX
                If Y = -1 Then Y = ViagemY
                For PlanetNum = MAX_PLANET_BASE + 3 To (MAX_PLANET_BASE + 3 + MAX_PLANET_CUSTOM - 2)
                    If Planets(PlanetNum).Map = GetPlayerMap(Index) Then Exit For
                Next PlanetNum
                If PlanetNum < (MAX_PLANET_BASE + 3 + MAX_PLANET_CUSTOM - 2) Then
                    If Planets(PlanetNum).Level <= 25 Then MapNum = ViagemMap
                    If Planets(PlanetNum).Level > 25 And Planets(PlanetNum).Level <= 50 Then MapNum = 53
                    If Planets(PlanetNum).Level > 50 And Planets(PlanetNum).Level <= 75 Then MapNum = 54
                    Call PlayerWarp(Index, MapNum, Planets(PlanetNum).X, Planets(PlanetNum).Y - 1)
                Else
                    Call PlayerWarp(Index, MapNum, X, Y)
                End If
            Else
                PlanetNum = GetPlanetNum(GetPlayerMap(Index))
                If Planets(PlanetNum).Level <= 25 Then MapNum = ViagemMap
                If Planets(PlanetNum).Level > 25 And Planets(PlanetNum).Level <= 50 Then MapNum = 53
                If Planets(PlanetNum).Level > 50 And Planets(PlanetNum).Level <= 75 Then MapNum = 54
                Call PlayerWarp(Index, MapNum, Planets(PlanetNum).X, Planets(PlanetNum).Y - 1)
            End If
        Else
            PlanetNum = PlayerMapIndex(GetPlayerMap(Index))
            MapNum = VirgoMap
            Call PlayerWarp(Index, MapNum, PlayerPlanet(PlanetNum).PlanetData.X, PlayerPlanet(PlanetNum).PlanetData.Y - 1)
        End If
        If MapNum = ViagemMap Then
            Call SendPlanets(Index)
        Else
            Call SendPlayerPlanets(Index)
        End If
    End If
End Sub

Function IsPlayerMap(ByVal MapNum As Long)
    IsPlayerMap = (MapNum > PlanetStart + MAX_PLANET_BASE And MapNum <= PlanetStart + MAX_PLANET_BASE + GetMaxPlayerPlanets)
End Function

Function GetPlayerNormalSprite(ByVal Index As Long) As Long
    GetPlayerNormalSprite = Player(Index).RealSprite  'ToDo
    TempPlayer(Index).Speed = 4
End Function

Function GetPlayerNave(ByVal Index As Long) As Long
    If TempPlayer(Index).Nave = 0 Then
        Dim i As Long
        For i = 1 To MAX_INV
            If Player(Index).Inv(i).Num > 0 Then
                If Item(GetPlayerInvItemNum(Index, i)).Type = ITEM_TYPE_NAVE Then
                    TempPlayer(Index).Nave = GetPlayerInvItemNum(Index, i)
                    Exit For
                End If
            End If
        Next i
    End If
    GetPlayerNave = Item(TempPlayer(Index).Nave).data1
End Function

Function MakeAlpha(ByVal Total As Long) As String
    Dim i As Long
    For i = 1 To Total
        MakeAlpha = MakeAlpha & Chr(rand(65, 90))
    Next i
End Function

Function MakeNumber(ByVal Total As Long) As String
    Dim i As Long
    For i = 1 To Total
        MakeNumber = MakeNumber & rand(0, 9)
    Next i
End Function

Sub SendSellPlanet(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMatchData
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteByte 0
    Buffer.WriteLong Index

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMatchData(ByVal MatchIndex As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMatchData
    Buffer.WriteLong MatchData(MatchIndex).Points
    Buffer.WriteLong MatchData(MatchIndex).TotalNpcs
    Buffer.WriteLong MatchData(MatchIndex).Stars
    Buffer.WriteByte MatchData(MatchIndex).Active
    Buffer.WriteLong Planets(MatchData(MatchIndex).Planet).PointsToConquest
    Buffer.WriteLong MatchData(MatchIndex).Winner
    
    Dim i As Long
    For i = 1 To UBound(MatchData(MatchIndex).Indexes)
        SendDataTo MatchData(MatchIndex).Indexes(i), Buffer.ToArray()
    Next i
    
    Set Buffer = Nothing
End Sub

Sub SendPlanets(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlanets
    
    Buffer.WriteByte 0
    Buffer.WriteByte 0
    Buffer.WriteLong MAX_PLANETS
    Dim i As Long
    For i = 1 To MAX_PLANETS
        Dim PlanetsSize As Long
        Dim PlanetsData() As Byte
        PlanetsSize = LenB(Planets(i))
        ReDim PlanetsData(PlanetsSize - 1)
        CopyMemory PlanetsData(0), ByVal VarPtr(Planets(i)), PlanetsSize
        Buffer.WriteBytes PlanetsData
    Next i
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlanet(ByVal Index As Long, ByVal PlanetNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlanets
    
    Buffer.WriteByte 0
    Buffer.WriteByte 1
    Buffer.WriteLong PlanetNum
    Dim PlanetsSize As Long
    Dim PlanetsData() As Byte
    PlanetsSize = LenB(Planets(PlanetNum))
    ReDim PlanetsData(PlanetsSize - 1)
    CopyMemory PlanetsData(0), ByVal VarPtr(Planets(PlanetNum)), PlanetsSize
    Buffer.WriteBytes PlanetsData
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlanetToAll(ByVal PlanetNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlanets
    
    Buffer.WriteByte 0
    Buffer.WriteByte 1
    Buffer.WriteLong PlanetNum
    Dim PlanetsSize As Long
    Dim PlanetsData() As Byte
    PlanetsSize = LenB(Planets(PlanetNum))
    ReDim PlanetsData(PlanetsSize - 1)
    CopyMemory PlanetsData(0), ByVal VarPtr(Planets(PlanetNum)), PlanetsSize
    Buffer.WriteBytes PlanetsData
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerPlanets(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlanets
    
    Buffer.WriteByte 1
    Buffer.WriteByte 0
    Buffer.WriteLong UBound(PlayerPlanet)
    Dim i As Long
    For i = 1 To UBound(PlayerPlanet)
        Dim PlanetsSize As Long
        Dim PlanetsData() As Byte
        PlanetsSize = LenB(PlayerPlanet(i).PlanetData)
        ReDim PlanetsData(PlanetsSize - 1)
        CopyMemory PlanetsData(0), ByVal VarPtr(PlayerPlanet(i).PlanetData), PlanetsSize
        Buffer.WriteBytes PlanetsData
    Next i
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerPlanet(ByVal Index As Long, ByVal PlanetNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlanets
    
    Buffer.WriteByte 1
    Buffer.WriteByte 1
    Buffer.WriteLong PlanetNum
    Dim PlanetsSize As Long
    Dim PlanetsData() As Byte
    PlanetsSize = LenB(PlayerPlanet(PlanetNum).PlanetData)
    ReDim PlanetsData(PlanetsSize - 1)
    CopyMemory PlanetsData(0), ByVal VarPtr(PlayerPlanet(PlanetNum).PlanetData), PlanetsSize
    Buffer.WriteBytes PlanetsData
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerPlanetToAll(ByVal PlanetNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlanets
    
    Buffer.WriteByte 1
    Buffer.WriteByte 1
    Buffer.WriteLong PlanetNum
    Dim PlanetsSize As Long
    Dim PlanetsData() As Byte
    PlanetsSize = LenB(PlayerPlanet(PlanetNum))
    ReDim PlanetsData(PlanetsSize - 1)
    CopyMemory PlanetsData(0), ByVal VarPtr(PlayerPlanet(PlanetNum)), PlanetsSize
    Buffer.WriteBytes PlanetsData
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendSaibamansToMap(ByVal PlanetNum As Long, ByVal MapNum As Long)
    Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                SendSaibamans i, PlanetNum
            End If
        End If
    Next i
End Sub
Sub SendSaibamans(ByVal Index As Long, ByVal PlanetNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSaibamans
    
    Buffer.WriteLong GetPlayerMap(Index)
    Buffer.WriteByte PlayerPlanet(PlanetNum).TotalSaibamans
    Dim i As Long
    For i = 1 To PlayerPlanet(PlanetNum).TotalSaibamans
        Buffer.WriteByte PlayerPlanet(PlanetNum).Saibaman(i).Working
        Buffer.WriteLong PlayerPlanet(PlanetNum).Saibaman(i).X
        Buffer.WriteLong PlayerPlanet(PlanetNum).Saibaman(i).Y
        Buffer.WriteString PlayerPlanet(PlanetNum).Saibaman(i).TaskInit
        If PlayerPlanet(PlanetNum).Saibaman(i).Working = 1 Then
            Dim Minutes As Long
            If PlayerPlanet(PlanetNum).Saibaman(i).TaskType = 0 Then
                Minutes = Npc(PlayerPlanet(PlanetNum).Saibaman(i).TaskResult).TimeToEvolute - (PlayerPlanet(PlanetNum).Saibaman(i).Accelerate * 60)
            End If
            If PlayerPlanet(PlanetNum).Saibaman(i).TaskType = 1 Then
                Minutes = Resource(PlayerPlanet(PlanetNum).Saibaman(i).TaskResult).TimeToEvolute - (PlayerPlanet(PlanetNum).Saibaman(i).Accelerate * 60)
            End If
            Buffer.WriteLong (Minutes - DateDiff("n", PlayerPlanet(PlanetNum).Saibaman(i).TaskInit, Now))
        Else
            Buffer.WriteLong 0
        End If
    Next i
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub ClearPlanet(ByVal Index As Long)
    With Planets(Index)
        RemovePlanetLocation .X, .Y, .Size
        .Name = vbNullString
        .Pic = 0
        .Size = 0
        .X = 0
        .Y = 0
        .Owner = vbNullString
    End With
End Sub


Sub CreatePlanet(ByVal Index As Long)
    
    Dim MuitoBaixo As Long, Baixo As Long, Medio As Long, Alto As Long, MuitoAlto As Long
    For i = 1 To MAX_PLANETS
        If Planets(i).Level < 10 Then
            MuitoBaixo = MuitoBaixo + 1
        End If
        If Planets(i).Level >= 10 And Planets(i).Level < 20 Then
            Baixo = Baixo + 1
        End If
        If Planets(i).Level >= 20 And Planets(i).Level < 45 Then
            Medio = Medio + 1
        End If
    Next i
    
    With Planets(Index)
        
        .Name = GerarNome(rand(4, 12)) 'Trim$(MakeAlpha(rand(3, 5)) & "-" & MakeNumber(rand(3, 6)))
        
        .Pic = rand(1, 2)
MakePosition:

        .X = rand(5, Map(ViagemMap).MaxX - 5)
        .Y = rand(5, Map(ViagemMap).MaxY - 5)
        
        If Map(ViagemMap).Tile(.X, .Y).Type = tile_type_blocked Or Map(ViagemMap).Tile(.X, .Y).Type = tile_type_warp Or PlanetLocations(.X, .Y) = True Then GoTo MakePosition
        
        For i = 1 To Index
            If Index <> i Then
                If Planets(i).X = .X And Planets(i).Y = .Y And Planets(i).Name = .Name Then
                    SetStatus "Duplicada encontrada: " & Index & " com " & i
                End If
            End If
            DoEvents
        Next i
        
        .TimeToExplode = 0
        
        Dim Distance As Long, DistanceX As Long, DistanceY As Long
        DistanceX = .X - (Map(ViagemMap).MaxX / 2)
        If DistanceX < 0 Then DistanceX = DistanceX * -1
        DistanceY = .Y - (Map(ViagemMap).MaxY / 2)
        If DistanceY < 0 Then DistanceY = DistanceY * -1
        Distance = DistanceX + DistanceY
        Dim MapSize As Long
        MapSize = (Map(ViagemMap).MaxX / 2) + (Map(ViagemMap).MaxY / 2)
        
        If MuitoBaixo < 25 Then
            .Size = rand(24, 28)
        Else
            If Distance < (MapSize * 0.2) Then .Size = rand(24, 32)
            If Distance >= (MapSize * 0.2) And Distance < (MapSize * 0.3) Then .Size = rand(32, 64)
            If Distance >= (MapSize * 0.3) And Distance < (MapSize * 0.45) Then .Size = rand(64, 82)
            If Distance >= (MapSize * 0.45) Then .Size = rand(82, 96)
        End If
        
        'If Baixo < 10 Then .Size = rand(28, 32)
        'If MuitoBaixo < 10 Then .Size = rand(24, 28)
        
        AddPlanetLocation .X, .Y, Int(.Size / 24)
        
        If .Size > 32 Then
            'Tem lua?
            If rand(1, 10) > 7 Then
                .MoonData.ColorR = rand(0, 255)
                .MoonData.ColorG = rand(0, 255)
                .MoonData.ColorB = rand(0, 255)
                .MoonData.Size = rand(16, 24)
                .MoonData.Speed = rand(25, 200)
                .MoonData.Pic = rand(1, 2)
            End If
        End If
        
        .ColorR = rand(0, 255)
        .ColorG = rand(0, 255)
        .ColorB = rand(0, 255)
        .State = 0
        
        Dim Factory As Byte
        Factory = rand(1, UBound(MapFactory))
        .TileConfig = MapFactory(Factory)
        PlanetTile(Index) = Factory
        .Especie = rand(1, 3)
        
        If .ColorR < 50 And .ColorG < 50 And .ColorB < 50 Then
            .ColorR = .ColorR + 100
            .ColorG = .ColorG + 100
            .ColorB = .ColorB + 100
        End If
        
        'Level
        Dim Factor As Long
        Factor = rand(1, 30)
        Factor = Factor + (.Size * 0.75)
        
        If Factor < 30 Then
            .Habitantes = rand(1, 5) 'Pouquissima habitação (1 - 5)
            .Level = MakeLevel(.Habitantes, 1, 5, 1, 5)
        End If
        If Factor >= 30 And Factor < 50 Then
            .Habitantes = rand(5, 10) 'Pouca habitação (5 - 15)
            .Level = MakeLevel(.Habitantes, 5, 15, 5, 10)
        End If
        If Factor >= 50 And Factor < 65 Then
            .Habitantes = rand(10, 20) 'Média habitação (15 - 25)
            .Level = MakeLevel(.Habitantes, 15, 25, 10, 20)
        End If
        If Factor >= 65 And Factor < 80 Then
            .Habitantes = rand(20, 35) 'Média alta habitação (25 - 55)
            .Level = MakeLevel(.Habitantes, 25, 55, 20, 35)
        End If
        If Factor >= 80 And Factor < 90 Then
            .Habitantes = rand(60, 80) 'Alta habitação (55 - 80)
            .Level = MakeLevel(.Habitantes, 55, 80, 60, 80)
        End If
        If Factor >= 90 Then
            .Habitantes = rand(110, 200) 'Muito alta habitação (80 - 95)
            .Level = MakeLevel(.Habitantes, 80, 95, 110, 200)
        End If
            
        Factor = rand(1, 30)
        Factor = Factor + (.Size * 0.75)
        
        If Factor < 30 Then
            .Gravidade = rand(5, 15)  'Pouquissima gravidade (1 - 5)
            .Level = (.Level + MakeLevel(.Gravidade, 1, 5, 5, 15)) / 2
        End If
        If Factor >= 30 And Factor < 50 Then
            .Gravidade = rand(15, 25) 'Pouca gravidade (5 - 15)
            .Level = (.Level + MakeLevel(.Gravidade, 5, 15, 15, 25)) / 2
        End If
        If Factor >= 50 And Factor < 65 Then
            .Gravidade = rand(25, 60) 'Média gravidade (15 - 25)
            .Level = (.Level + MakeLevel(.Gravidade, 15, 25, 25, 60)) / 2
        End If
        If Factor >= 65 And Factor < 80 Then
            .Gravidade = rand(60, 80) 'Média alta gravidade (25 - 55)
            .Level = (.Level + MakeLevel(.Gravidade, 25, 55, 60, 80)) / 2
        End If
        If Factor >= 80 And Factor < 90 Then
            .Gravidade = rand(200, 350) 'Alta gravidade (55 - 80)
            .Level = (.Level + MakeLevel(.Gravidade, 55, 80, 200, 350)) / 2
        End If
        If Factor >= 90 Then
            .Gravidade = rand(500, 1500) 'Muito alta gravidade (80 - 95)
            .Level = (.Level + MakeLevel(.Gravidade, 80, 95, 500, 1500)) / 2
        End If
        
        If .Level > MAX_LEVELS Then .Level = MAX_LEVELS
        If .Level <= 6 Then .Level = rand(1, 6)
        
        If LCase(Trim$(Planets(Index).Name)) <> "planeta desconhecido" Then
            If .Level > 8 Then
                .Type = rand(0, 5)
            Else
                .Type = 0
            End If
        Else
            .Type = 6
        End If
        
        
        If .Type <> 1 Then
            .PointsToConquest = 50
            .WaveDuration = 6000
            .WaveCooldown = 15000
            If .Level > 5 And .Level <= 10 Then
                .PointsToConquest = 75
                .WaveDuration = 6000
                .WaveCooldown = 15000
            ElseIf .Level > 10 And .Level <= 20 Then
                .PointsToConquest = 100
                .WaveDuration = 8000
                .WaveCooldown = 15000
            ElseIf .Level > 20 And .Level <= 30 Then
                .PointsToConquest = 200
                .WaveDuration = 11000
                .WaveCooldown = 15000
            ElseIf .Level > 30 And .Level <= 40 Then
                .PointsToConquest = 300
                .WaveDuration = 14000
                .WaveCooldown = 15000
            ElseIf .Level > 40 And .Level <= 50 Then
                .PointsToConquest = 450
                .WaveDuration = 16000
                .WaveCooldown = 15000
            ElseIf .Level > 50 And .Level <= 60 Then
                .PointsToConquest = 550
                .WaveDuration = 18000
                .WaveCooldown = 15000
            ElseIf .Level > 60 And .Level <= 70 Then
                .PointsToConquest = 700
                .WaveDuration = 20000
                .WaveCooldown = 15000
            ElseIf .Level > 70 And .Level <= 80 Then
                .PointsToConquest = 850
                .WaveDuration = 20000
                .WaveCooldown = 15000
            ElseIf .Level > 80 And .Level <= 90 Then
                .PointsToConquest = 1000
                .WaveDuration = 20000
                .WaveCooldown = 15000
            ElseIf .Level > 95 Then
                .PointsToConquest = 1500
                .WaveDuration = 22000
                .WaveCooldown = 15000
            End If
        End If
        
        If .Type = 2 Then
            .PointsToConquest = rand(20, 30) + (.Level * 2)
        End If
        
        If .Type = 4 Then
            .PointsToConquest = rand(60, 90) + .Level
            .WaveCooldown = 10000
        End If
        
        'Preco
        Factor = rand(1, 100)
        
        If Factor < 35 Then
            .Atmosfera = rand(80, 95)  'Pesada (0 - 2)
        End If
        If Factor >= 35 And Factor < 60 Then
            .Atmosfera = rand(45, 60) 'Mediana (8 - 15)
        End If
        If Factor >= 60 And Factor < 85 Then
            .Atmosfera = rand(10, 25) 'Pouca (30 - 50)
        End If
        If Factor >= 85 Then
            .Atmosfera = 0 'Limpa (65 - 80)
        End If
        
        Dim PriceFactor As Long
        
        PriceFactor = 250
        If .Size > 24 Then PriceFactor = 350
        If .Size > 32 Then PriceFactor = 500
        If .Size > 48 Then PriceFactor = 750
        If .Size > 64 Then PriceFactor = 1000
        If .Size > 80 Then PriceFactor = 1500
        
        Dim AtmosferaPrice As Double
        AtmosferaPrice = (0.5 + ((100 - .Atmosfera) / 100))
        
        .Preco = (.Size * PriceFactor) * AtmosferaPrice
        
        'Especiaria
        .EspeciariaVermelha = rand(0, 100)
        .EspeciariaAzul = rand(0, (100 - .EspeciariaVermelha))
        .EspeciariaAmarela = rand(0, (100 - .EspeciariaAzul - .EspeciariaVermelha))
        
    End With
End Sub

Function MakeLevel(Value As Long, MinLevel As Long, MaxLevel As Long, Min As Long, Max As Long) As Long
    MakeLevel = MinLevel + ((MaxLevel - MinLevel) * ((Value - Min) / (Max - Min)))
End Function


Sub CreateAllPlanetMaps()
    Dim i As Long
    
    Dim GlobalTick As Long, Tick As Long
    
    For i = 1 To MAX_PLANET_BASE + 1
        Tick = GetTickCount
        CreatePlanetMap i
        Tick = GetTickCount - Tick
        If GlobalTick = 0 Then
            GlobalTick = Tick
        Else
            GlobalTick = (GlobalTick + Tick) / 2
        End If
        frmServer.Caption = "Criando mapas dos planetas: " & ((i / MAX_PLANETS) * 100) & "% (Tempo médio: " & GlobalTick & "ms)"
    Next i
End Sub

Sub CreatePlanetMap(ByVal Index As Long)
    Dim X As Long, Y As Long, n As Long, i As Long
    Dim LastX As Long, LastY As Long
    Dim MapNum As Long
    Dim Ambient As Byte
    
    Randomize
    
    Select Case PlanetTile(Index)
        Case 1
            Ambient = 1
        Case 2
            Ambient = 2
        Case 3
            Ambient = 2
        Case 4
            Ambient = 3
        Case 5
            Ambient = 3
    End Select
    
    MapNum = PlanetStart + (Index - 1)
    Planets(Index).Map = MapNum
    
    Map(MapNum).Name = Planets(Index).Name
    
    Dim Music As Long
    Music = rand(1, 8)
    If Music < 6 Then
        Map(MapNum).Music = "BATLE 0" & Music & ".mid"
    Else
        Map(MapNum).Music = "BATLE 0" & Music & ".mp3"
    End If
    
    If Planets(Index).Type <> 1 Then
        Map(MapNum).MaxX = 14 + Int(Planets(Index).Size / 2)
        Map(MapNum).MaxY = 14 + Int(Planets(Index).Size / 2)
    Else
        Map(MapNum).MaxX = MAX_MAPX
        Map(MapNum).MaxY = MAX_MAPY
    End If
    
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    
    If Planets(Index).Type = 2 Then
        Dim SelectedResource As Long
        SelectedResource = rand(56, 58)
    End If
    
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            For n = 1 To 8
                For i = 1 To TileConfigEnum.TileConfigCount - 1
                    If Planets(Index).TileConfig.Tile(i).Layer = n Then
                        If i > 1 Then
                            If rand(1, 100) < 95 Or Map(MapNum).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then GoTo NextTile
                        End If
                        Map(MapNum).Tile(X, Y).Layer(n).Tileset = Planets(Index).TileConfig.Tile(i).Tileset
                        Map(MapNum).Tile(X, Y).Layer(n).X = Planets(Index).TileConfig.Tile(i).X
                        Map(MapNum).Tile(X, Y).Layer(n).Y = Planets(Index).TileConfig.Tile(i).Y
                        If i = 3 Then 'block
                            Map(MapNum).Tile(X, Y).Type = tile_type_blocked
                        End If
                    End If
NextTile:
                Next i
            Next n
            If (rand(1, 100) >= 96 And Planets(Index).Type <> 1) Then
                If X + 2 > LastX Or X - 2 < LastX Or Y + 2 > LastY Or Y - 2 < LastY Or LastX = 0 Or LastY = 0 Then
                    'Resource
                    Dim ResourceType As Long
                    Static LastResource As Long
                    If Planets(i).Level < 10 Then ResourceType = 1 'Apenas pequenos
                    If Planets(i).Level >= 10 And Planets(i).Level < 20 Then ResourceType = rand(1, 2) 'Pequenos e médios
                    If Planets(i).Level >= 20 Then ResourceType = rand(1, 3) 'Pequenos médios e grandes
                    If ResourceFactory(ResourceType).Resource(1) <> 0 Then
                        Map(MapNum).Tile(X, Y).Type = TILE_TYPE_RESOURCE
                        Dim ResourceNum As Long
                        If Planets(Index).Type <> 2 Then
                            If rand(1, 2) = 1 Or Planets(Index).Type = 3 Then
                                ResourceNum = rand(1, UBound(ResourceFactory(ResourceType).Resource))
                                If ResourceNum = LastResource Then
                                    Do While ResourceNum = LastResource
                                        ResourceNum = rand(1, UBound(ResourceFactory(ResourceType).Resource))
                                    Loop
                                End If
                                Map(MapNum).Tile(X, Y).data1 = ResourceFactory(ResourceType).Resource(ResourceNum)
                            Else
                               'Decoração
                                Select Case Ambient
                                    Case 1 'Grama
                                        ResourceNum = rand(1, 2)
                                   Case 2 'Arvores decaidas 2 e 3
                                        ResourceNum = rand(59, 61)
                                    Case 3 'Pedras 4 e 5
                                        ResourceNum = rand(62, 63)
                                End Select
                                Map(MapNum).Tile(X, Y).data1 = ResourceNum
                            End If
                        Else
                            ResourceNum = SelectedResource
                            Map(MapNum).Tile(X, Y).data1 = SelectedResource
                        End If
                        
                        LastResource = ResourceNum
                        LastX = X
                        LastY = Y
                    End If
                End If
            End If
        Next Y
        DoEvents
    Next X
    
    Map(MapNum).Red = Planets(Index).ColorR
    Map(MapNum).Green = Planets(Index).ColorG
    Map(MapNum).Blue = Planets(Index).ColorB
    Map(MapNum).Alpha = rand(20, 80)
    
    If Planets(Index).Atmosfera > 0 Then 'Fog
        Map(MapNum).Fog = rand(1, 10)
        Map(MapNum).FogOpacity = 255 - Planets(Index).Atmosfera
        
        Dim FogSpeed As Long
        FogSpeed = rand(1, 100)
        
        'Low
        If FogSpeed < 30 Then
            Map(MapNum).FogSpeed = rand(10, 20)
        End If
        
        'Medium
        If FogSpeed >= 30 And FogSpeed < 80 Then
            Map(MapNum).FogSpeed = rand(80, 400)
        End If
        
        'High
        If FogSpeed >= 80 Then
            Map(MapNum).FogSpeed = rand(600, 1000)
        End If
    End If
    
    'Clima
    If rand(1, 100) > 50 Then
        Map(MapNum).Weather = rand(1, 5)
        Map(MapNum).WeatherIntensity = rand(50, 100)
    End If
    
    Select Case Planets(Index).Type
        Case 1 'Mini boss
            Dim NpcIndex As Long
            For NpcIndex = NPCS_BOSSES_START To MAX_NPCS
                If Len(Trim$(Npc(NpcIndex).Name)) = 0 Then
                    Exit For
                End If
            Next NpcIndex
            If NpcIndex = MAX_NPCS Then
                SetStatus "Falta de npcs para gerar bosses!"
            End If
            Npc(NpcIndex).Name = GerarNome(rand(4, 10))
            Npc(NpcIndex).Level = Planets(Index).Level
            Npc(NpcIndex).Sprite = rand(68, 137)
            Npc(NpcIndex).ND = 0
            Npc(NpcIndex).Damage = NPCBase(Planets(i).Level).Damage
            Npc(NpcIndex).HP = NPCBase(Planets(Index).Level).HP * rand(15, 25)
            Npc(NpcIndex).stat(Stats.agility) = NPCBase(Planets(Index).Level).Acc * 2
            Npc(NpcIndex).stat(Stats.Willpower) = NPCBase(Planets(Index).Level).Esq
            Npc(NpcIndex).Exp = Int(ExperienceBase(Planets(Index).Level) / 2)
            Npc(NpcIndex).Behaviour = 0
            Npc(NpcIndex).Drop(1).Num = MoedaZ
            Npc(NpcIndex).Drop(1).Value = Planets(Index).Preco / 2
            Npc(NpcIndex).Speed = rand(3, 10)
            Npc(NpcIndex).Range = rand(5, 10)
            Npc(NpcIndex).Animation = 3
            If rand(0, 100) <= 75 Then
                Npc(NpcIndex).Ranged = 1
                Npc(NpcIndex).ArrowAnim = rand(3, 27)
                Npc(NpcIndex).ArrowDamage = rand(80, 120)
                Npc(NpcIndex).ArrowAnimation = PlayerAttackAnim
            End If
            If rand(0, 100) <= 75 Then
                Npc(NpcIndex).IA(NPCIA.Shunppo).Data(1) = 1
                Npc(NpcIndex).IA(NPCIA.Shunppo).Data(2) = rand(10, 15)
                Npc(NpcIndex).IA(NPCIA.Shunppo).Data(3) = rand(0, 1)
                Npc(NpcIndex).IA(NPCIA.Shunppo).Data(4) = 21
            End If
            If rand(0, 100) <= 75 Then
                Npc(NpcIndex).IA(NPCIA.Stun).Data(1) = 1
                Npc(NpcIndex).IA(NPCIA.Stun).Data(2) = rand(10, 15)
                Npc(NpcIndex).IA(NPCIA.Stun).Data(3) = rand(1, 3)
                Npc(NpcIndex).IA(NPCIA.Stun).Data(4) = rand(1, 3)
                Npc(NpcIndex).IA(NPCIA.Stun).Data(5) = 2
            End If
            Map(MapNum).Npc(1) = NpcIndex
            SpawnNpc 1, MapNum, True
        Case 5
            For NpcIndex = 1 To MAX_MAP_NPCS
                Dim NpcNum As Long, Min As Long, Max As Long, Specie As Long, SelectNPC As Long
                Specie = Planets(Index).Especie
                If rand(0, 100) < 10 Then
                    NpcNum = rand(1, UBound(NpcFactory(Specie).Elites))
                    If NpcNum > UBound(NpcFactory(Specie).Elites) Then NpcNum = UBound(NpcFactory(Specie).Elites)
                    SelectNPC = NpcFactory(Specie).Elites(NpcNum)
                Else
                    Min = 1
                    Max = 1 + Int(NpcIndex / 3)
                    If Max > UBound(NpcFactory(Specie).Npc) Then Max = UBound(NpcFactory(Specie).Npc)
                    NpcNum = rand(Min, Max)
                    SelectNPC = NpcFactory(Specie).Npc(NpcNum)
                End If
                Map(MapNum).Npc(NpcIndex) = SelectNPC
                SpawnNpc NpcIndex, MapNum, True
            Next
    End Select
    
    CacheResources MapNum
    If Decorations And Planets(Index).Type <> 2 Then MakeDecorations MapNum
    
End Sub

Sub MakeDecorations(ByVal MapNum As Long)
    Dim ResourceNum As Long
    Dim ResourcesSelected() As Boolean
    Dim TotalPaths As Long
    Dim Start As Long, Finish As Long
    
    If ResourceCache(MapNum).Resource_Count = 0 Then Exit Sub
    ReDim ResourcesSelected(1 To ResourceCache(MapNum).Resource_Count)
    
    TotalPaths = rand(16, 32)
    If TotalPaths > ResourceCache(MapNum).Resource_Count Then TotalPaths = ResourceCache(MapNum).Resource_Count
    
    Dim i As Long
    For i = 1 To ResourceCache(MapNum).Resource_Count
        If Resource(ResourceCache(MapNum).ResourceData(i).ResourceNum).ResourceType < 6 And ResourcesSelected(i) = False Then
            If Start = 0 Then
                Start = i
            Else
                Finish = i
                Exit For
            End If
        End If
    Next i
    
    If Start > 0 And Finish > 0 Then
    
MakePath:

        Dim ChooseDir As Long
        ChooseDir = rand(1, 2)
    
        ResourcesSelected(Start) = True
        ResourcesSelected(Finish) = True
        
        Dim X As Long
        Dim Y As Long
        
        X = ResourceCache(MapNum).ResourceData(Start).X
        Y = ResourceCache(MapNum).ResourceData(Start).Y
        
        
        Do Until X = ResourceCache(MapNum).ResourceData(Finish).X And Y = ResourceCache(MapNum).ResourceData(Finish).Y
            If ChooseDir = 1 Then
                If X < ResourceCache(MapNum).ResourceData(Finish).X Then
                    X = X + 1
                Else
                    If X > ResourceCache(MapNum).ResourceData(Finish).X Then
                        X = X - 1
                    Else
                        If Y < ResourceCache(MapNum).ResourceData(Finish).Y Then
                            Y = Y + 1
                        Else
                            If Y > ResourceCache(MapNum).ResourceData(Finish).Y Then
                                Y = Y - 1
                            End If
                        End If
                    End If
                End If
            Else
                If Y < ResourceCache(MapNum).ResourceData(Finish).Y Then
                    Y = Y + 1
                Else
                    If Y > ResourceCache(MapNum).ResourceData(Finish).Y Then
                        Y = Y - 1
                    Else
                        If X < ResourceCache(MapNum).ResourceData(Finish).X Then
                            X = X + 1
                        Else
                            If X > ResourceCache(MapNum).ResourceData(Finish).X Then
                                X = X - 1
                            End If
                        End If
                    End If
                End If
            End If
            Map(MapNum).Tile(X, Y).Autotile(2) = 1
            Map(MapNum).Tile(X, Y).Layer(2).Tileset = 1
            Map(MapNum).Tile(X, Y).Layer(2).X = 2
            Map(MapNum).Tile(X, Y).Layer(2).Y = 28
            Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE
        Loop
        
        Dim Steps As Long
        Steps = 0
        Start = Finish
        i = Finish
        Do Until (Resource(ResourceCache(MapNum).ResourceData(i).ResourceNum).ResourceType < 6 And ResourcesSelected(i) = False And (Abs(ResourceCache(MapNum).ResourceData(i).X - ResourceCache(MapNum).ResourceData(Start).X) + Abs(ResourceCache(MapNum).ResourceData(i).Y - ResourceCache(MapNum).ResourceData(Start).Y) > 20)) Or Steps > ResourceCache(MapNum).Resource_Count
            i = rand(1, ResourceCache(MapNum).Resource_Count)
            Steps = Steps + 1
        Loop
        Finish = i
        TotalPaths = TotalPaths - 1
        If Steps <= ResourceCache(MapNum).Resource_Count And TotalPaths > 0 Then GoTo MakePath
    End If
    
End Sub

Sub AddPlanetLocation(ByVal X As Long, ByVal Y As Long, ByVal Size As Long)
    Dim rX As Long, rY As Long
    For rX = X - Size To X + Size
        For rY = Y - Size To Y + Size
            If rX >= 0 And rX <= Map(ViagemMap).MaxX Then
                If rY >= 0 And rY <= Map(ViagemMap).MaxY Then
                    PlanetLocations(rX, rY) = True
                End If
            End If
        Next rY
    Next rX
End Sub

Sub RemovePlanetLocation(ByVal X As Long, ByVal Y As Long, ByVal Size As Long)
    Dim rX As Long, rY As Long
    For rX = X - Size To X + Size
        For rY = Y - Size To Y + Size
            If rX >= 0 And rX <= Map(ViagemMap).MaxX Then
                If rY >= 0 And rY <= Map(ViagemMap).MaxY Then
                    PlanetLocations(rX, rY) = False
                End If
            End If
        Next rY
    Next rX
End Sub

Sub StartMatch(ByVal Index As Long, PlanetNum As Long)
    Dim MaxMatchs As Long
    Dim MatchNum As Long
    Dim PartyNum As Long
    Dim LoseMemberCount As Long
    MaxMatchs = UBound(MatchData)
    
    For i = 1 To MaxMatchs
        If MatchData(i).Active = 0 Then
            MatchNum = i
            Exit For
        End If
    Next i
    
    If MatchNum = 0 Then
        ReDim Preserve MatchData(1 To MaxMatchs + 1) As MatchDataRec
        MatchNum = MaxMatchs + 1
    End If
    
    With MatchData(MatchNum)
        .Active = 1
        .Points = 0
        .SpawnTick = GetTickCount
        .WaveTick = GetTickCount + Planets(PlanetNum).WaveDuration
        .TotalNpcs = 0
        .HighLevel = 0
        .Planet = PlanetNum
        .Stars = 0
        .PriceBonus = 0
        If TempPlayer(Index).inParty > 0 Then
            PartyNum = TempPlayer(Index).inParty
            ' check members in outhers maps
            Dim Count As Long
            Count = 0
            For i = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party(PartyNum).Member(i)
                If tmpIndex > 0 Then
                    If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                        If GetPlayerMap(tmpIndex) = ViagemMap Then
                            Count = Count + 1
                            ReDim Preserve .Indexes(1 To Count)
                            .Indexes(Count) = tmpIndex
                            If GetPlayerLevel(tmpIndex) > .HighLevel Then .HighLevel = GetPlayerLevel(tmpIndex)
                        End If
                    End If
                End If
            Next i
        Else
            ReDim .Indexes(1 To 1)
            .Indexes(1) = Index
            .HighLevel = GetPlayerLevel(Index)
        End If
        .WaveNum = 1
        .Winner = 0
    End With
    
    For i = 1 To MAX_MAP_NPCS
        DespawnNPC Planets(PlanetNum).Map, Val(i)
    Next i
    
    CacheResources Planets(PlanetNum).Map
    
    For i = 1 To UBound(MatchData(MatchNum).Indexes)
        TempPlayer(MatchData(MatchNum).Indexes(i)).MatchIndex = MatchNum
        Call SetPlayerSprite(MatchData(MatchNum).Indexes(i), GetPlayerNormalSprite(Index))
        Call PlayerWarp(MatchData(MatchNum).Indexes(i), Planets(PlanetNum).Map, Map(Planets(PlanetNum).Map).MaxX / 2, Map(Planets(PlanetNum).Map).MaxY / 2)
    Next i
End Sub

Function NpcCount(ByVal MapNum As Long) As Long
    Dim i As Long
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(MapNum).Npc(i).Num > 0 Then
            If MapNpc(MapNum).Npc(i).Vital(Vitals.HP) > 0 Then
                NpcCount = NpcCount + 1
            End If
        End If
    Next i
End Function

Function SelectNPC(ByVal MatchIndex As Long) As Long
    Dim NpcNum As Long, Min As Long, Max As Long, Specie As Long
    Specie = Planets(MatchData(MatchIndex).Planet).Especie
    If Specie <= 0 Then Exit Function
    If EliteWave(MatchIndex) Then
        NpcNum = Int(MatchData(MatchIndex).WaveNum / 10) + 1
        If NpcNum > UBound(NpcFactory(Specie).Elites) Then NpcNum = UBound(NpcFactory(Specie).Elites)
        SelectNPC = NpcFactory(Specie).Elites(NpcNum)
    Else
        Min = 1
        Max = 1 + Int(MatchData(MatchIndex).WaveNum / 3)
        If Max > UBound(NpcFactory(Specie).Npc) Then Max = UBound(NpcFactory(Specie).Npc)
        NpcNum = rand(Min, Max)
        SelectNPC = NpcFactory(Specie).Npc(NpcNum)
    End If
End Function

Function EliteWave(ByVal MatchIndex As Long) As Boolean
    EliteWave = MatchData(MatchIndex).EliteWave
End Function

Sub DespawnNPC(ByVal MapNum As Long, ByVal MapNPCNum As Long)
    MapNpc(MapNum).Npc(MapNPCNum).Num = 0
    MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) = 0
    UpdateMapBlock MapNum, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, False
    
    Dim i As Long
    'Loop through entire map and purge NPC from targets
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And IsConnected(i) Then
            If Player(i).Map = MapNum Then
                If TempPlayer(i).TargetType = TARGET_TYPE_NPC Then
                    If TempPlayer(i).Target = MapNPCNum Then
                        TempPlayer(i).Target = 0
                        TempPlayer(i).TargetType = TARGET_TYPE_NONE
                        SendTarget i
                    End If
                End If
            End If
        End If
    Next
    
    ' send death to the map
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDead
    Buffer.WriteLong MapNPCNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetNPCSlot(ByVal MapNum As Long) As Long
    Dim i As Long
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(MapNum).Npc(i).Num = 0 Or MapNpc(MapNum).Npc(i).Vital(Vitals.HP) <= 0 Then
            GetNPCSlot = i
            Exit Function
        End If
    Next i
End Function

Sub SpawnEnemy(ByVal MatchIndex As Long)
    Dim MapNum As Long
    MapNum = Planets(MatchData(MatchIndex).Planet).Map
    
    Dim NpcOnMap As Long
    NpcOnMap = NpcCount(MapNum)
    
    If NpcOnMap + 1 <= MAX_MAP_NPCS Then
        Dim NpcNum As Long
        NpcNum = SelectNPC(MatchIndex)
        Dim NpcSlot As Long
        NpcSlot = GetNPCSlot(MapNum)
        SpawnWaveNPC MapNum, NpcSlot, NpcNum, MatchData(MatchIndex).Indexes(rand(1, UBound(MatchData(MatchIndex).Indexes))), NpcNum
        
        Dim X As Long, Y As Long
        Dim XLimit As Long, YLimit As Long
        Dim SelectedIndex As Long
        SelectedIndex = MatchData(MatchIndex).Indexes(rand(1, UBound(MatchData(MatchIndex).Indexes)))
        X = GetPlayerX(SelectedIndex) - 10
        If X < 0 Then X = 0
        Y = GetPlayerY(SelectedIndex) - 10
        If Y < 0 Then Y = 0
        XLimit = GetPlayerX(SelectedIndex) + 10
        If XLimit > Map(MapNum).MaxX Then XLimit = Map(MapNum).MaxX
        YLimit = GetPlayerY(SelectedIndex) + 10
        If YLimit > Map(MapNum).MaxY Then YLimit = Map(MapNum).MaxY
        
        MapNpc(MapNum).Npc(NpcSlot).X = rand(X, XLimit)
        MapNpc(MapNum).Npc(NpcSlot).Y = rand(Y, YLimit)
        SendMapNpcXY NpcSlot, MapNum
        
        SendAnimation MapNum, SpawnAnim, MapNpc(MapNum).Npc(NpcSlot).X, MapNpc(MapNum).Npc(NpcSlot).Y, MapNpc(MapNum).Npc(NpcSlot).Dir
        MatchData(MatchIndex).TotalNpcs = NpcOnMap + 1
        SendMatchData MatchIndex
    Else
        'Perdeu a invasão
        MatchData(MatchIndex).Active = 0
        Dim n As Long
        For n = 1 To MAX_MAP_NPCS
            DespawnNPC MapNum, n
        Next n
        n = 0
        MaxIndexes = UBound(MatchData(MatchIndex).Indexes)
        For n = 1 To MaxIndexes
            If n <= UBound(MatchData(MatchIndex).Indexes) Then
                If IsPlaying(MatchData(MatchIndex).Indexes(n)) Then
                    GiveConquistaReward MatchData(MatchIndex).Indexes(n), UBound(MatchData(MatchIndex).Indexes), MatchData(MatchIndex).HighLevel, (MatchData(MatchIndex).Points / Planets(MatchData(MatchIndex).Planet).PointsToConquest) * 100
                    Dim MyViagemMap As Long
                    If Planets(MatchData(MatchIndex).Planet).Level <= 25 Then MyViagemMap = ViagemMap
                    If Planets(MatchData(MatchIndex).Planet).Level > 25 And Planets(MatchData(MatchIndex).Planet).Level <= 50 Then
                        MyViagemMap = 53
                    End If
                    If Planets(MatchData(MatchIndex).Planet).Level > 50 Then
                        MyViagemMap = 54
                    End If
                    Viajar MatchData(MatchIndex).Indexes(n), MyViagemMap, Planets(MatchData(MatchIndex).Planet).X, Planets(MatchData(MatchIndex).Planet).Y
                    SendBossMsg MyViagemMap, "Você não conseguiu dominar este planeta o suficiente!", brightred, MatchData(MatchIndex).Indexes(n)
                End If
            End If
        Next n
        SendMatchData MatchIndex
    End If
    
End Sub

Sub EnterPlanet(ByVal Index As Long, ByVal PlanetIndex As Long)
    If Planets(PlanetIndex).State = 0 Then
        'Iniciar invasão
        If Planets(PlanetIndex).Especie > 0 And Planets(PlanetIndex).Type = 0 Then
            Call SendBossMsg(GetPlayerMap(Index), "Você iniciou uma invasão no planeta " & Trim$(Planets(PlanetIndex).Name) & " prepare-se para a onda de inimigos!", Yellow, Index)
            Call StartMatch(Index, PlanetIndex)
            Planets(PlanetIndex).State = 1
            SendPlanetToAll PlanetIndex
        Else
            If LCase(Trim$(Planets(PlanetIndex).Name)) <> "planeta desconhecido" Then
                Select Case Planets(PlanetIndex).Type
                    Case 1: Call SendBossMsg(Planets(PlanetIndex).Map, "Derrote o chefe que reside neste planeta!", Yellow, Index)
                    Case 2
                        Call SendBossMsg(Planets(PlanetIndex).Map, "Destrua " & Planets(PlanetIndex).PointsToConquest & " pedras preciosas para coletar seus recursos!", Yellow, Index)
                        Call StartMatch(Index, PlanetIndex)
                        Planets(PlanetIndex).State = 1
                        SendPlanetToAll PlanetIndex
                    Case 3
                        Call SendBossMsg(Planets(PlanetIndex).Map, "Destrua todas as construções deste planeta!", Yellow, Index)
                        Call StartMatch(Index, PlanetIndex)
                        Planets(PlanetIndex).State = 1
                        SendPlanetToAll PlanetIndex
                    Case 4
                        Call SendBossMsg(Planets(PlanetIndex).Map, "Destrua os habitantes deste planeta em busca de " & Planets(PlanetIndex).PointsToConquest & " tesouros!", Yellow, Index)
                        Call StartMatch(Index, PlanetIndex)
                        Planets(PlanetIndex).State = 1
                        SendPlanetToAll PlanetIndex
                    Case 5
                        Call SendBossMsg(Planets(PlanetIndex).Map, "Piratas espaciais estão atacando este planeta em posse dos Sayajins! Derrote todos!", Yellow, Index)
                        Dim HaveNPCs As Boolean
                        For n = 1 To MAX_MAP_NPCS
                            If MapNpc(Planets(PlanetIndex).Map).Npc(n).Num > 0 And MapNpc(Planets(PlanetIndex).Map).Npc(n).Vital(Vitals.HP) > 0 Then
                                MapNpc(Planets(PlanetIndex).Map).Npc(n).Target = Index
                                MapNpc(Planets(PlanetIndex).Map).Npc(n).TargetType = TargetType.TARGET_TYPE_PLAYER
                                HaveNPCs = True
                            End If
                        Next n
                        If Not HaveNPCs Then
                            For n = 1 To MAX_MAP_NPCS
                                SpawnNpc n, Planets(PlanetIndex).Map, True
                            Next n
                        End If
                End Select
            End If
            Call SetPlayerSprite(Index, GetPlayerNormalSprite(Index))
            Call PlayerWarp(Index, Planets(PlanetIndex).Map, Map(Planets(PlanetIndex).Map).MaxX / 2, Map(Planets(PlanetIndex).Map).MaxY / 2)
        End If
        Call SetPlayerVital(Index, MP, GetPlayerMaxVital(Index, MP))
        SendVital Index, MP
    Else
        Call SetPlayerSprite(Index, GetPlayerNormalSprite(Index))
        Call PlayerWarp(Index, Planets(PlanetIndex).Map, Map(Planets(PlanetIndex).Map).MaxX / 2, Map(Planets(PlanetIndex).Map).MaxY / 2)
    End If
End Sub

Sub RemoveFromMatchData(ByVal Index As Long)
    If UBound(MatchData(TempPlayer(Index).MatchIndex).Indexes) - 1 = 0 Then
        'Perdeu
        MatchData(TempPlayer(Index).MatchIndex).Active = 0
        SendMatchData TempPlayer(Index).MatchIndex
    Else
        'Ainda há jogadores
        Dim NewIndexes() As Long
        Dim i As Long, z As Long
        
        ReDim NewIndexes(1 To UBound(MatchData(TempPlayer(Index).MatchIndex).Indexes) - 1)
        
        z = 1
        For i = 1 To UBound(MatchData(TempPlayer(Index).MatchIndex).Indexes) - 1
            If MatchData(TempPlayer(Index).MatchIndex).Indexes(i) <> Index Then
                NewIndexes(z) = MatchData(TempPlayer(Index).MatchIndex).Indexes(i)
                z = z + 1
            End If
        Next i
        
        ReDim MatchData(TempPlayer(Index).MatchIndex).Indexes(1 To UBound(NewIndexes)) As Long
        MatchData(TempPlayer(Index).MatchIndex).Indexes = NewIndexes
    End If
    TempPlayer(Index).MatchIndex = 0
End Sub
Function GetPlanetMaxExtractor(ByVal PlanetNum As Long) As Byte
    GetPlanetMaxExtractor = 1 + Int((Planets(PlanetNum).Size - 24) / 16)
End Function
Function CountExtractor(ByVal MapNum As Long) As Long
    CountExtractor = ResourceCache(MapNum).ExtractorCount
End Function
Function GetPlanetNum(ByVal MapNum As Long) As Long
    GetPlanetNum = (MapNum - PlanetStart) + 1
End Function
Sub PutExtractor(ByVal Index As Long, InvNum As Long)
    Dim ItemNum As Long, X As Long, Y As Long, MapNum As Long
    Dim PlanetNum As Long, Resource_Count As Long
    
    ItemNum = GetPlayerInvItemNum(Index, InvNum)
    MapNum = GetPlayerMap(Index)
    If MapNum >= PlanetStart And MapNum <= PlanetStart + MAX_PLANET_BASE Then
        PlanetNum = GetPlanetNum(MapNum)
        X = GetPlayerX(Index)
        Y = GetPlayerY(Index) - 1
        
        If Planets(PlanetNum).State = 2 Then
            If Trim$(Planets(PlanetNum).Owner) = Trim$(GetPlayerName(Index)) Then
                If Y >= 0 And Y <= Map(MapNum).MaxY Then
                    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                        If CountExtractor(MapNum) < GetPlanetMaxExtractor(PlanetNum) Then
                            Map(MapNum).Tile(X, Y).Type = TILE_TYPE_RESOURCE
                            Map(MapNum).Tile(X, Y).data1 = Item(ItemNum).data2
                            Resource_Count = ResourceCache(MapNum).Resource_Count
                            Resource_Count = Resource_Count + 1
                            ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                            ResourceCache(MapNum).ResourceData(Resource_Count).X = X
                            ResourceCache(MapNum).ResourceData(Resource_Count).Y = Y
                            ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = 1
                            ResourceCache(MapNum).ResourceData(Resource_Count).ResourceTimer = GetTickCount
                            ResourceCache(MapNum).ResourceData(Resource_Count).ResourceState = 1
                            ResourceCache(MapNum).ResourceData(Resource_Count).ResourceNum = Item(ItemNum).data2
                            ResourceCache(MapNum).ExtractorCount = ResourceCache(MapNum).ExtractorCount + 1
                            ResourceCache(MapNum).Resource_Count = Resource_Count
                            TakeInvItem Index, ItemNum, 1
                            SendResourceCacheToMap MapNum
                        Else
                            PlayerMsg Index, "Você não pode colocar mais de " & GetPlanetMaxExtractor(PlanetNum) & " extratores neste planeta! (Maximo de 1 para cada 8,000 km de raio do planeta)", brightred
                        End If
                    Else
                        PlayerMsg Index, "Você não pode colocar seu extrator aqui!", brightred
                    End If
                End If
            Else
                PlayerMsg Index, "Você não é o dono deste planeta!", brightred
            End If
        Else
            PlayerMsg Index, "É necessário capturar um planeta primeiro para poder inserir um extrator!", brightred
        End If
    Else
        PlayerMsg Index, "É necessário colocar o extrator em um planeta capturado!", brightred
    End If
End Sub
Sub ExtractorGet(ByVal Index As Long, ResourceNum As Long, ResourceIndex As Long)
    Dim Quant As Long
    Dim Perc As Long
    Dim Color As Long
    Dim PlanetNum As Long
    Dim ResourceTimer As Long
    ResourceTimer = (Resource(ResourceNum).RespawnTime * (10 * (100 - ResourceCache(Player(Index).Map).ResourceData(ResourceIndex).cur_health)))
    Dim Collects As Double
    Collects = (GetTickCount - ResourceCache(Player(Index).Map).ResourceData(ResourceIndex).ResourceTimer)
    Collects = (Collects / ResourceTimer) + 1
    
    PlanetNum = (GetPlayerMap(Index) - PlanetStart) + 1
    If LCase(Trim$(Item(Resource(ResourceNum).ItemReward).Name)) = "esp. amarela" Then
        Perc = Planets(PlanetNum).EspeciariaAmarela
        Color = Yellow
    End If
    If LCase(Trim$(Item(Resource(ResourceNum).ItemReward).Name)) = "esp. vermelha" Then
        Perc = Planets(PlanetNum).EspeciariaVermelha
        Color = brightred
    End If
    If LCase(Trim$(Item(Resource(ResourceNum).ItemReward).Name)) = "esp. azul" Then
        Perc = Planets(PlanetNum).EspeciariaAzul
        Color = brightblue
    End If
    
    Quant = Int((rand(50, 100) / 100) * Perc)
    Quant = Int(Quant * Options.ResourceFactor)
    Quant = Quant * Collects
    Dim bonus As Long
    If Planets(PlanetNum).Level > 10 Then bonus = (Planets(PlanetNum).Level - 10) * 10
    Quant = (Quant / 100) * (100 + bonus)
    SendActionMsg GetPlayerMap(Index), "+" & Quant, Color, 1, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
    GiveInvItem Index, Resource(ResourceNum).ItemReward, Quant, True
    
    If LCase(Trim$(Item(Resource(ResourceNum).ItemReward).Name)) = "esp. amarela" Then
        If IsDaily(Index, GetYellow) Then UpdateDaily Index, Quant
    End If
    If LCase(Trim$(Item(Resource(ResourceNum).ItemReward).Name)) = "esp. vermelha" Then
        If IsDaily(Index, GetRed) Then UpdateDaily Index, Quant
    End If
    If LCase(Trim$(Item(Resource(ResourceNum).ItemReward).Name)) = "esp. azul" Then
        If IsDaily(Index, GetBlue) Then UpdateDaily Index, Quant
    End If
    If IsDaily(Index, GetAny) Then UpdateDaily Index, Quant
    
End Sub
Sub Evacuate(ByVal MapNum As Long, Optional StartMap As Boolean = False)
    Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If Not StartMap Then
                    Viajar i, ViagemMap
                Else
                    PlayerWarp i, START_MAP, START_X, START_Y
                End If
            End If
        End If
    Next i
End Sub

Sub GiveConquistaReward(ByVal Index As Long, ByVal PlayerDivision As Long, HigherLevel As Long, ByVal PorcComplete As Long)
    Dim PlanetNum As Long
    PlanetNum = MatchData(TempPlayer(Index).MatchIndex).Planet
    
    If Planets(PlanetNum).Type <> 0 Then
        RemoveTemporaryItems Index
        Exit Sub
    End If
        
    Dim Exp As Long, GuildExp As Long
    Exp = ExperienceBase(Planets(PlanetNum).Level)
    GuildExp = Planets(PlanetNum).Level
    
    'Level factor
    Dim LevelDifference As Long
    LevelDifference = GetPlayerLevel(Index) - Planets(PlanetNum).Level
    If LevelDifference > 0 Then
        Exp = (Exp / 100) * (100 - (20 * Int(LevelDifference / 3)))
    End If
    
    Dim HighExp As Long
    If HigherLevel = 0 Then HigherLevel = GetPlayerLevel(Index)
    HighExp = Experience(HigherLevel)
    Dim Porc As Double
    Porc = GetPlayerNextLevel(Index) / HighExp
    If Porc > 1 Then Porc = 1
    
    Exp = Int(Exp * Porc)
    
    'Bonus
    Exp = (Exp / 100) * (100 + (10 * (PlayerDivision - 1)))
    
    If Exp < 125 Then Exp = 125
    
    If MatchData(TempPlayer(Index).MatchIndex).Stars > 0 Then
        Dim StarPercentage As Long
        StarPercentage = Int(MatchData(TempPlayer(Index).MatchIndex).Stars / 20)
        If StarPercentage < 1 Then StarPercentage = 1
        If StarPercentage > 3 Then StarPercentage = 3
        Exp = (Exp / 100) * (100 + MatchData(TempPlayer(Index).MatchIndex).Stars * StarPercentage)
        GuildExp = (GuildExp / 100) * (100 + MatchData(TempPlayer(Index).MatchIndex).Stars * StarPercentage)
        If MatchData(TempPlayer(Index).MatchIndex).Stars > Player(Index).TopStars Then
            Player(Index).TopStars = MatchData(TempPlayer(Index).MatchIndex).Stars
            UpdateWebRank LCase(GetPlayerLogin(Index)), "stars", Player(Index).TopStars
        End If
    End If
    
    Exp = Exp * Val(frmServer.txtExpFactor)
    
    Dim ExpFinal As Long
    If Player(Index).EsoBonus > 0 Then
        ExpFinal = ((Exp / 100) * (100 + Player(Index).EsoBonus))
    Else
        ExpFinal = Exp
    End If
    
    If PorcComplete < 100 Then
        Exp = ((Exp / 2) / 100) * PorcComplete
        ExpFinal = Exp
        Call PlayerMsg(Index, "Você não completou a invasão, e recebeu apenas: " & ExpFinal & "xp", brightred)
    Else
        If IsDaily(Index, DestruaPlanetas) Then UpdateDaily Index
        If Planets(PlanetNum).Level <= GetPlayerLevel(Index) / 2 And IsDaily(Index, DestruaPlanetasHalf) Then UpdateDaily Index
        If MatchData(TempPlayer(Index).MatchIndex).Stars > 0 Then
            PlayerMsg Index, "Parabéns! Você concluiu a invasão com " & MatchData(TempPlayer(Index).MatchIndex).Stars & " estrelas!", Yellow
        End If
        PlayerMsg Index, "-- Recompensa pela missão --", brightgreen
        Call PlayerMsg(Index, "Você recebeu: " & ExpFinal & "xp", brightgreen)
        Call PlayerMsg(Index, "Seu poder de luta aumentou em: " & Int(ExpFinal * ExpToPDL), brightgreen)
        GivePlayerVIPExp Index, Planets(Index).Level
        If Player(Index).Guild > 0 Then GiveGuildExp Index, GuildExp, Planets(PlanetNum).Level
    End If
    
    GivePlayerEXP Index, Exp
    
    If TempPlayer(Index).PlanetService = PlanetNum And PorcComplete >= 100 Then
        CompleteService Index
    End If
End Sub

Sub EnhanceExtractor(ByVal Index As Long, ByVal InvNum As Long)
    Dim Resource_num As Long, i As Long, Resource_index As Long, X As Long, Y As Long
    
    ' Check tradeskills
    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            X = GetPlayerX(Index) - 1
            Y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
            X = GetPlayerX(Index) + 1
            Y = GetPlayerY(Index)
    End Select
    
    Resource_num = 0
    Resource_index = Map(GetPlayerMap(Index)).Tile(X, Y).data1
    
    If Resource_index > 0 Then
    If Resource(Resource_index).ResourceType = 4 Then
        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            If ResourceCache(GetPlayerMap(Index)).ResourceData(i).X = X Then
                If ResourceCache(GetPlayerMap(Index)).ResourceData(i).Y = Y Then
                    Resource_num = i
                End If
            End If
        Next
        
        Dim bonus As Long
        bonus = Item(GetPlayerInvItemNum(Index, InvNum)).data1
        
        If ResourceCache(Player(Index).Map).ResourceData(Resource_num).cur_health < bonus Then
            ResourceCache(Player(Index).Map).ResourceData(Resource_num).cur_health = bonus
            TakeInvItem Index, GetPlayerInvItemNum(Index, InvNum), 1
            SendActionMsg GetPlayerMap(Index), "Combustível adicionado (+" & bonus & "% velocidade)", brightgreen, 1, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        Else
            PlayerMsg Index, "Este extrator ja está com este combustivel! Espere até que ele termine de extrair para colocar outro!", brightred
        End If
    End If
    End If
End Sub
Sub SendDragonballs()
    Dim ItemNum(1 To 7) As Long
    Dim i As Long, PlanetNum As Long
    Dim n As Long
    If Not UZ Then Exit Sub
    For i = 1 To MAX_ITEMS
        If Item(i).Type = ITEM_TYPE_DRAGONBALL Then
            ItemNum(Item(i).Dragonball) = i
        End If
    Next i
    
    For i = 1 To 7
        If ItemNum(i) > 0 Then
            If DragonballInfo(i) = 0 Then
                PlanetNum = rand(1, MAX_PLANET_BASE)
                Dim X As Long, Y As Long
define:
                X = rand(1, Map(Planets(PlanetNum).Map).MaxX)
                Y = rand(1, Map(Planets(PlanetNum).Map).MaxY)
                If Map(Planets(PlanetNum).Map).Tile(X, Y).Type <> TILE_TYPE_WALKABLE Then GoTo define
                n = FindOpenMapItemSlot(Planets(PlanetNum).Map)
                MapItem(Planets(PlanetNum).Map, n).Num = ItemNum(i)
                MapItem(Planets(PlanetNum).Map, n).X = X
                MapItem(Planets(PlanetNum).Map, n).Y = Y
                MapItem(Planets(PlanetNum).Map, n).canDespawn = False
                SpawnItemSlot n, ItemNum(i), 1, Planets(PlanetNum).Map, X, Y, , False
                DragonballInfo(i) = Planets(PlanetNum).Map
            End If
        End If
    Next i
    UpdateDragonballList
End Sub

Sub UpdateDragonballList()
    Dim i As Long
    frmServer.lstEsferas.Clear
    For i = 1 To 7
        If DragonballInfo(i) = 0 Then
            frmServer.lstEsferas.AddItem "Esfera " & i & " capturada"
        Else
            frmServer.lstEsferas.AddItem "Esfera " & i & " no mapa " & DragonballInfo(i)
        End If
    Next i
End Sub

Sub PlanetaDoTesouro()
    Dim PlanetNum As Long
    PlanetNum = MAX_PLANET_BASE + 2
    
    CreatePlanet PlanetNum
    
    Planets(PlanetNum).Name = "Planeta do tesouro"
    Planets(PlanetNum).Map = TesouroMap
    Planets(PlanetNum).Size = 96
    Planets(PlanetNum).ColorR = 255
    Planets(PlanetNum).ColorG = 255
    Planets(PlanetNum).ColorB = 20
    Planets(PlanetNum).EspeciariaAmarela = 100
    Planets(PlanetNum).EspeciariaVermelha = 100
    Planets(PlanetNum).EspeciariaAzul = 100
    Planets(PlanetNum).Atmosfera = 0
    Planets(PlanetNum).Gravidade = 10
    Planets(PlanetNum).Level = 1
    Planets(PlanetNum).Especie = 0
    Planets(PlanetNum).Preco = 1000000000
    
    GlobalMsg "O Planeta do tesouro acabou de aparecer! Ele só ficará por algumas horas portanto seja rápido!", brightgreen
    
    SendPlanetToAll PlanetNum
    TesouroStarted = True
    
End Sub

Sub CreateCustomPlanets()
    Dim PlanetNum As Long, i As Long, filename As String
    i = MAX_PLANET_BASE + 3
    
    Dim n As Long
    n = 0
    
    For PlanetNum = i To (i + MAX_PLANET_CUSTOM - 2)
        n = n + 1
        filename = App.path & "\data\planets\" & n & ".ini"
        
        CreatePlanet PlanetNum
        Planets(PlanetNum).Name = Trim$(GetVar(filename, "PLANET", "Name"))
        Planets(PlanetNum).Map = Val(GetVar(filename, "PLANET", "Map"))
        Planets(PlanetNum).Size = Val(GetVar(filename, "PLANET", "Size"))
        Planets(PlanetNum).ColorR = Val(GetVar(filename, "PLANET", "Red"))
        Planets(PlanetNum).ColorG = Val(GetVar(filename, "PLANET", "Green"))
        Planets(PlanetNum).ColorB = Val(GetVar(filename, "PLANET", "Blue"))
        Planets(PlanetNum).EspeciariaAmarela = 0
        Planets(PlanetNum).EspeciariaVermelha = 0
        Planets(PlanetNum).EspeciariaAzul = 0
        Planets(PlanetNum).Atmosfera = Val(GetVar(filename, "PLANET", "Atmosfera"))
        Planets(PlanetNum).Gravidade = Val(GetVar(filename, "PLANET", "Gravidade"))
        Planets(PlanetNum).Level = Val(GetVar(filename, "PLANET", "Level"))
        Planets(PlanetNum).Preco = 0
        Planets(PlanetNum).Especie = 0
        SendPlanetToAll PlanetNum
        
    Next PlanetNum
    
    'Planeta HIDRA
    'PlanetNum = PlanetNum + 1
    'CreatePlanet PlanetNum
    'Planets(PlanetNum).name = "Planeta desconhecido"
    'Planets(PlanetNum).Map = 32
    'Planets(PlanetNum).Size = 40
    'Planets(PlanetNum).ColorR = 0
    'Planets(PlanetNum).ColorG = 0
    'Planets(PlanetNum).ColorB = 255
    'Planets(PlanetNum).EspeciariaAmarela = 0
    'Planets(PlanetNum).EspeciariaVermelha = 0
    'Planets(PlanetNum).EspeciariaAzul = 0
    'Planets(PlanetNum).Atmosfera = 25
    'Planets(PlanetNum).Gravidade = 10
    'Planets(PlanetNum).Level = 20
    'Planets(PlanetNum).Preco = 0
    'Planets(PlanetNum).Especie = 0
    'SendPlanetToAll PlanetNum
    
End Sub

Sub DestroyTesouro()
    Dim PlanetNum As Long
    PlanetNum = MAX_PLANET_BASE + 2
    
    Evacuate Planets(PlanetNum).Map
    
    Call ZeroMemory(ByVal VarPtr(Planets(PlanetNum)), LenB(Planets(PlanetNum)))
    GlobalMsg "O Planeta do tesouro acabou de desaparecer!", brightred
    TesouroStarted = False
    SendPlanetToAll PlanetNum
End Sub
Function PlayerAlive(ByVal MatchIndex As Long) As Boolean
    Dim n As Long
    For n = 1 To UBound(MatchData(MatchIndex).Indexes)
        If IsPlaying(MatchData(MatchIndex).Indexes(n)) Then
            If Player(MatchData(MatchIndex).Indexes(n)).IsDead = 0 Then
                PlayerAlive = True
                Exit Function
            End If
        End If
    Next n
End Function
Function IsElite(ByVal NpcNum As Long) As Boolean
    Dim Name As String
    Name = LCase(Trim$(Npc(NpcNum).Name))
    IsElite = (Mid(Name, Len(Name) - 4, 5) = "elite")
End Function
Sub RollWave(ByVal MatchIndex As Long)
    MatchData(MatchIndex).EliteWave = False
    If MatchData(MatchIndex).WaveNum >= 5 Then
        If rand(1, 10) = 1 Then
            MatchData(MatchIndex).EliteWave = True
        End If
    End If
End Sub
Sub ClearWaveItems(ByVal MapNum As Long)
    Dim i As Long, ItemNum As Long
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num > 0 Then
            ItemNum = MapItem(MapNum, i).Num
            If Item(ItemNum).Type = ItemType.ITEM_TYPE_CONSUME Then
                ClearMapItem i, MapNum
                Call SpawnItemSlot(i, 0, 0, MapNum, 0, 0)
            End If
        End If
    Next i
End Sub
Sub HandleScriptedResource(ByVal Index As Long, ByVal ResourceIndex As Long)
    Dim MatchIndex As Long, PlanetNum As Long
    MatchIndex = TempPlayer(Index).MatchIndex
    PlanetNum = MatchData(MatchIndex).Planet
    If MatchIndex = 0 Then Exit Sub
    Select Case Resource(ResourceIndex).ItemReward
        Case 1 'Spawn elite
            MatchData(MatchIndex).EliteWave = True
            SpawnEnemy MatchIndex
            MatchData(MatchIndex).EliteWave = False
            SendBossMsg GetPlayerMap(Index), "Você destruiu um centro militar e chamou a atenção do exército!", Yellow, Index
        Exit Sub
        Case 2 'Estrela extra
            MatchData(MatchIndex).Stars = MatchData(MatchIndex).Stars + 4
            SendMatchData MatchIndex
        Exit Sub
        Case 3 'Próxima onda
            MatchData(MatchIndex).WaveTick = -(MatchData(TempPlayer(Index).MatchIndex).WaveTick - GetTickCount)
            SendActionMsg GetPlayerMap(Index), "Próxima onda ativada!", brightred, ActionMsgType.ACTIONMSG_STATIC, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        Exit Sub
        Case 4 'Aumentar preço em 10%
            MatchData(MatchIndex).PriceBonus = MatchData(MatchIndex).PriceBonus + 10
            SendBossMsg GetPlayerMap(Index), "Wow as financias do planeta estão ótimas! O valor do terreno subiu em " & MatchData(MatchIndex).PriceBonus & "%", Yellow, Index
        Exit Sub
        Case 5 'Aumentar destruição reduzir preço
            MatchData(MatchIndex).Points = MatchData(MatchIndex).Points + (Planets(PlanetNum).PointsToConquest * 0.1)
            MatchData(MatchIndex).Stars = MatchData(MatchIndex).Stars + 4
            If MatchData(MatchIndex).Points > Planets(PlanetNum).PointsToConquest Then MatchData(MatchIndex).Points = Planets(PlanetNum).PointsToConquest
            Planets(PlanetNum).Preco = Planets(PlanetNum).Preco * 0.9
            SendMatchData MatchIndex
            SendBossMsg GetPlayerMap(Index), "Wow isso deve deixar um baita estrago... pena que os compradores não gostarão... (Preço reduzido para: " & Planets(PlanetNum).Preco & "z)", Yellow, Index
    End Select
End Sub

Function ExperienceBase(ByVal Level As Long) As Long
    ExperienceBase = Experience(Level) / Int(Level * 2)
End Function

Function GetTotalCustomPlanets() As Long
    Dim i As Long
    i = 1
    Do While FileExist(App.path & "\data\planets\" & i & ".ini", True)
        i = i + 1
    Loop
    i = i - 1
    GetTotalCustomPlanets = i
End Function

Function GerarNome(NumCaracteres As Long) As String
    Dim Name As String
    Dim i As Long
    Name = GetLetra
    For i = 2 To NumCaracteres
        Dim Char As String
        Dim DoubleChar As String
        Char = Mid(Name, i - 1, 1)
        If i > 2 Then DoubleChar = Mid(Name, i - 2, 2)
        If Char = "Q" Then
            Name = Name & "U"
        Else
            If DoubleChar = "QU" Then
                Name = Name & GetVogal(True)
            Else
                If CanGetConsoante(Char, DoubleChar) Then
                    Name = Name & GetConsoante
                Else
                    Name = Name & GetVogal
                End If
            End If
        End If
    Next i
    GerarNome = Name
End Function
Function CorrectCase(ByVal Name As String) As String
    CorrectCase = UCase(Mid(Name, 1, 1)) & LCase(Mid(Name, 2, Len(Name)))
End Function
Function CanGetConsoante(ByVal Char As String, DoubleChar As String) As Boolean
    If IsVogal(GetAscii(Char)) Then CanGetConsoante = True
    If Char = "R" And DoubleChar <> "RR" Then CanGetConsoante = True
    If Char = "P" And DoubleChar <> "PP" Then CanGetConsoante = True
    If Char = "L" And DoubleChar <> "LL" Then CanGetConsoante = True
    If Char = "N" And DoubleChar <> "NN" Then CanGetConsoante = True
    If Char = "C" And DoubleChar <> "CC" Then CanGetConsoante = True
    If Char = "S" And DoubleChar <> "SS" Then CanGetConsoante = True
    If Not IsVogal(GetAscii(Mid(DoubleChar, 1, 1))) And Not IsVogal(GetAscii(Mid(DoubleChar, 2, 1))) Then CanGetConsoante = False
End Function
Function GetAscii(ByVal Char As String) As Byte
    Dim i As Long
    For i = 1 To 255
        If Chr(i) = Char Then
            GetAscii = i
            Exit Function
        End If
    Next i
End Function
Function IsVogal(ByVal Ascii As Long) As Boolean
    IsVogal = False
    Select Case Ascii
        Case 65: IsVogal = True
        Case 69: IsVogal = True
        Case 73: IsVogal = True
        Case 79: IsVogal = True
        Case 85: IsVogal = True
    End Select
End Function
Function GetVogal(Optional NotU As Boolean = False) As String
    Dim i As Long
    If NotU Then
        i = rand(1, 4)
    Else
        i = rand(1, 5)
    End If
    Select Case i
        Case 1: GetVogal = "A"
        Case 2: GetVogal = "E"
        Case 3: GetVogal = "I"
        Case 4: GetVogal = "O"
        Case 5: GetVogal = "U"
    End Select
End Function
Function GetConsoante() As String
    Dim i As Long
Gerar:
    i = rand(65, 90)
    If IsVogal(i) Then GoTo Gerar
    GetConsoante = Chr(i)
End Function
Function GetLetra() As String
    Dim i As Long
    i = rand(65, 90)
    GetLetra = Chr(i)
End Function

Sub LoadPlayerPlanetMap(ByVal MapNum As Long, ByRef Origin As MapRec, ByVal PlayerPlanetNum As Long)
    'CopyMemory ByVal VarPtr(Map(MapNum)), ByVal VarPtr(Origin), LenB(Origin)
    Map(MapNum) = Origin
    PlayerMapIndex(MapNum) = PlayerPlanetNum
    
    ReDim Preserve Map(MapNum).Tile(0 To Origin.MaxX, 0 To Origin.MaxY)
    
    MapCache_Create MapNum
    CacheResources MapNum
    
    PlayerPlanet(PlayerPlanetNum).PlanetData.Map = MapNum
End Sub

Sub SetPlayerPlanetMap(ByVal MapNum As Long, ByVal Origin As Long, ByVal PlayerPlanetNum As Long, ByVal PlanetNum As Long)
    PlayerPlanet(PlayerPlanetNum).PlanetMap = Map(Origin)
    PlayerPlanet(PlayerPlanetNum).PlanetData = Planets(PlanetNum)
    LoadPlayerPlanetMap MapNum, Map(Origin), PlayerPlanetNum
End Sub

Sub ExplodePlanet(ByVal PlanetNum As Long)
    SendAnimation Planets(PlanetNum).Map, 38, Planets(PlanetNum).X, Planets(PlanetNum).Y, 0
    Evacuate Planets(PlanetNum).Map
    ClearPlanet PlanetNum
    CreatePlanet PlanetNum
    CreatePlanetMap PlanetNum
    MapCache_Create Planets(PlanetNum).Map
    SendPlanetToAll PlanetNum
End Sub

Sub CapturePlanet(ByVal Index As Long, ItemNum As Long, Optional Force As Boolean = False)
    Dim PlanetNum As Long
    If GetPlayerMap(Index) >= PlanetStart And GetPlayerMap(Index) <= PlanetStart + MAX_PLANET_BASE Then
        PlanetNum = GetPlanetNum(GetPlayerMap(Index))
        If Planets(PlanetNum).State = 2 And Trim$(LCase(Planets(PlanetNum).Owner)) = Trim$(LCase(GetPlayerName(Index))) Then
            If Player(Index).PlanetNum = 0 Or Force Then
                Dim PlayerPlanetNum As Long, MapNum As Long, CreateNew As Boolean
                If Player(Index).PlanetNum > 0 Then
                    PlayerPlanetNum = Player(Index).PlanetNum
                    If PlayerPlanetNum > UBound(PlayerPlanet) Then
                        CreateNew = True
                    Else
                        If Trim$(LCase(PlayerPlanet(PlayerPlanetNum).PlanetData.Owner)) <> Trim$(LCase(GetPlayerName(Index))) Then
                            CreateNew = True
                        End If
                    End If
                Else
                    CreateNew = True
                End If
                
                If CreateNew Then
                    PlayerPlanetNum = GetMaxPlayerPlanets + 1
                    ReDim Preserve PlayerPlanet(1 To PlayerPlanetNum)
                End If
                PlayerPlanet(PlayerPlanetNum).TotalSaibamans = 1
                Player(Index).PlanetNum = PlayerPlanetNum
                
                MapNum = GetPlayerMap(Index)
                
                Dim n As Long
                For n = 1 To MAX_MAP_NPCS
                    Map(MapNum).Npc(n) = 0
                Next n
                
                MapNum = PlanetStart + MAX_PLANET_BASE + PlayerPlanetNum
                
                Planets(PlanetNum).Level = GetPlayerLevel(Index)
                
                SetPlayerPlanetMap MapNum, GetPlayerMap(Index), PlayerPlanetNum, PlanetNum
                Reposition PlayerPlanetNum
                
                PlayerWarp Index, MapNum, GetPlayerX(Index), GetPlayerY(Index)
                ExplodePlanet PlanetNum
                CacheResources MapNum
                MapCache_Create MapNum
                SendMap Index, MapNum
                SendResourceCacheToMap MapNum, , True
                SavePlayer Index
                SavePlayerPlanet PlayerPlanetNum
                SendPlayerPlanetToAll PlayerPlanetNum
                TakeInvItem Index, ItemNum, 1
                PlayerMsg Index, "Parabéns! Você capturou este planeta!", brightgreen
            Else
                'PlayerMsg Index, "Você já tem um planeta capturado!", brightred
                TempPlayer(Index).ConfirmationVar = ItemNum
                SendConfirmation Index, "Você já tem um planeta capturado! Se capturar este, o seu antigo planeta será destruído com tudo que existe nele e nada será reembolsado, deseja capturar este planeta?", 1
            End If
        Else
            PlayerMsg Index, "Este planeta não é seu para capturar!", brightred
        End If
    Else
        PlayerMsg Index, "Você deve estar em um planeta dominado para poder capturar!", brightred
    End If
End Sub
Function AddSaibaman(ByVal Index As Long, ByVal PlanetNum As Long, ByVal MapNum As Long, ByVal TaskType As Long, ByVal TaskResult As Long, Optional X As Long = -1, Optional Y As Long = -1) As Boolean
    Dim i As Long
    If X = -1 Then X = GetPlayerX(Index)
    If Y = -1 Then Y = GetPlayerY(Index)
    If PlayerPlanet(PlanetNum).TotalSaibamans = 0 Then PlayerPlanet(PlanetNum).TotalSaibamans = 1
    For i = 1 To PlayerPlanet(PlanetNum).TotalSaibamans
        If PlayerPlanet(PlanetNum).Saibaman(i).Working = 1 Then
            If PlayerPlanet(PlanetNum).Saibaman(i).X = X And PlayerPlanet(PlanetNum).Saibaman(i).Y = Y Then
                PlayerMsg Index, "Já existe um saibaman trabalhando neste local!", brightred
                Exit Function
            End If
        End If
    Next i
    For i = 1 To PlayerPlanet(PlanetNum).TotalSaibamans
        If PlayerPlanet(PlanetNum).Saibaman(i).Working = 0 Then
            PlayerPlanet(PlanetNum).Saibaman(i).Working = 1
            PlayerPlanet(PlanetNum).Saibaman(i).Accelerate = 0
            PlayerPlanet(PlanetNum).Saibaman(i).TaskInit = Now
            PlayerPlanet(PlanetNum).Saibaman(i).TaskResult = TaskResult
            PlayerPlanet(PlanetNum).Saibaman(i).TaskType = TaskType
            PlayerPlanet(PlanetNum).Saibaman(i).X = X
            PlayerPlanet(PlanetNum).Saibaman(i).Y = Y
            SendSaibamans Index, PlanetNum
            AddSaibaman = True
            Exit Function
        End If
    Next i
    PlayerMsg Index, "Você não tem saibamans operários disponíveis! Limite atual: " & PlayerPlanet(PlanetNum).TotalSaibamans, brightred
End Function
Sub UpdateSaibaman(PlanetNum As Long, SaibamanIndex As Long)
    Dim i As Long, n As Long
    Dim Minutes As Long
    Dim MapNum As Long
    MapNum = PlayerPlanet(PlanetNum).PlanetData.Map
    i = PlanetNum
    n = SaibamanIndex
    
    If PlayerPlanet(i).Saibaman(n).Working = 1 Then
        If PlayerPlanet(i).Saibaman(n).TaskType = 0 Then
            Minutes = Npc(PlayerPlanet(i).Saibaman(n).TaskResult).TimeToEvolute
        End If
        If PlayerPlanet(i).Saibaman(n).TaskType = 1 Then
            Minutes = Resource(PlayerPlanet(i).Saibaman(n).TaskResult).TimeToEvolute
        End If
        If PlayerPlanet(i).Saibaman(n).Accelerate = 1 Then Minutes = 0
        If DateDiff("n", PlayerPlanet(i).Saibaman(n).TaskInit, Now) >= Minutes Then
            
            If PlayerPlanet(i).Saibaman(n).TaskType = 0 Then
                Dim X As Long
                For X = 1 To MAX_MAP_NPCS
                    If Map(MapNum).Npc(X) = 0 Then
                        PlayerPlanet(i).Saibaman(n).Working = 0
                        Map(MapNum).Npc(X) = PlayerPlanet(i).Saibaman(n).TaskResult
                        Map(MapNum).Tile(PlayerPlanet(i).Saibaman(n).X, PlayerPlanet(i).Saibaman(n).Y).Type = TileType.TILE_TYPE_NPCSPAWN
                        Map(MapNum).Tile(PlayerPlanet(i).Saibaman(n).X, PlayerPlanet(i).Saibaman(n).Y).data1 = X
                        SpawnNpc X, MapNum
                        SendMapNpcsToMap MapNum
                        
                        PlayerPlanet(i).PlanetMap = Map(MapNum)
                        MapCache_Create MapNum
                        Dim Index As Long
                        ' Refresh map for everyone online
                        For Index = 1 To Player_HighIndex
                            If IsPlaying(Index) And GetPlayerMap(Index) = MapNum Then
                                SendMap Index, MapNum
                            End If
                        Next Index
                        SavePlayerPlanet i
                        SendPlayerPlanetToAll i
                        Exit For
                    End If
                Next X
            End If
            
            If PlayerPlanet(i).Saibaman(n).TaskType = 1 Then
                PlayerPlanet(i).Saibaman(n).Working = 0
                Map(MapNum).Tile(PlayerPlanet(i).Saibaman(n).X, PlayerPlanet(i).Saibaman(n).Y).Type = TILE_TYPE_RESOURCE
                Map(MapNum).Tile(PlayerPlanet(i).Saibaman(n).X, PlayerPlanet(i).Saibaman(n).Y).data1 = PlayerPlanet(i).Saibaman(n).TaskResult
                If Resource(PlayerPlanet(i).Saibaman(n).TaskResult).ItemReward > 0 Then AddExtrator i, PlayerPlanet(i).Saibaman(n).X, PlayerPlanet(i).Saibaman(n).Y
                Dim Resource_Count As Long
                Resource_Count = ResourceCache(MapNum).Resource_Count
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                ResourceCache(MapNum).ResourceData(Resource_Count).X = PlayerPlanet(i).Saibaman(n).X
                ResourceCache(MapNum).ResourceData(Resource_Count).Y = PlayerPlanet(i).Saibaman(n).Y
                ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(PlayerPlanet(i).Saibaman(n).TaskResult).health
                ResourceCache(MapNum).ResourceData(Resource_Count).ResourceTimer = GetTickCount
                ResourceCache(MapNum).ResourceData(Resource_Count).ResourceState = 0
                ResourceCache(MapNum).ResourceData(Resource_Count).ResourceNum = PlayerPlanet(i).Saibaman(n).TaskResult
                ResourceCache(MapNum).Resource_Count = Resource_Count
                MapCache_Create MapNum
                SendResourceCacheToMap MapNum, Resource_Count, True
                PlayerPlanet(i).PlanetMap = Map(MapNum)
                SavePlayerPlanet i
                SendPlayerPlanetToAll i
            End If
            
        End If
    End If
End Sub

Sub AddExtrator(ByVal PlanetNum As Long, X As Long, Y As Long, Optional When As String = vbNullString)
    Dim i As Long
    For i = 1 To 20
        If PlayerPlanet(PlanetNum).Extrator(i).Used = 0 Then
            If When = vbNullString Then
                PlayerPlanet(PlanetNum).Extrator(i).TaskInit = Now
            Else
                PlayerPlanet(PlanetNum).Extrator(i).TaskInit = When
            End If
            PlayerPlanet(PlanetNum).Extrator(i).X = X
            PlayerPlanet(PlanetNum).Extrator(i).Y = Y
            PlayerPlanet(PlanetNum).Extrator(i).Used = 1
            Exit Sub
        End If
    Next i
End Sub

Sub RemoveExtrator(ByVal PlanetNum As Long, X As Long, Y As Long)
    Dim i As Long
    For i = 1 To 20
        If PlayerPlanet(PlanetNum).Extrator(i).Used = 1 And PlayerPlanet(PlanetNum).Extrator(i).X = X And PlayerPlanet(PlanetNum).Extrator(i).Y = Y Then
            PlayerPlanet(PlanetNum).Extrator(i).TaskInit = vbNullString
            PlayerPlanet(PlanetNum).Extrator(i).X = 0
            PlayerPlanet(PlanetNum).Extrator(i).Y = 0
            PlayerPlanet(PlanetNum).Extrator(i).Used = 0
            Exit Sub
        End If
    Next i
End Sub

Sub SwitchExtrator(ByVal PlanetNum As Long, X As Long, Y As Long, NewX As Long, NewY As Long)
    Dim i As Long
    For i = 1 To 20
        If PlayerPlanet(PlanetNum).Extrator(i).Used = 1 And PlayerPlanet(PlanetNum).Extrator(i).X = X And PlayerPlanet(PlanetNum).Extrator(i).Y = Y Then
            AddExtrator PlanetNum, NewX, NewY, PlayerPlanet(PlanetNum).Extrator(i).TaskInit
            RemoveExtrator PlanetNum, X, Y
            Exit Sub
        End If
    Next i
End Sub

Function GetExtratorIndex(ByVal PlanetNum As Long, X As Long, Y As Long) As Long
    Dim i As Long
    For i = 1 To 20
        If PlayerPlanet(PlanetNum).Extrator(i).Used = 1 And PlayerPlanet(PlanetNum).Extrator(i).X = X And PlayerPlanet(PlanetNum).Extrator(i).Y = Y Then
            GetExtratorIndex = i
            Exit Function
        End If
    Next i
End Function

Sub UpdatePlanets()
    If UZ Then
        Dim MapNum As Long
        Dim i As Long, n As Long
        Dim Minutes As Long
        For i = 1 To UBound(PlayerPlanet)
            MapNum = PlayerPlanet(i).PlanetData.Map
            If MapNum > 0 Then
            For n = 1 To PlayerPlanet(i).TotalSaibamans
                UpdateSaibaman i, n
            Next n
            If PlayersOnMap(MapNum) Then SendSaibamansToMap i, MapNum
            For n = 1 To 5
                If PlayerPlanet(i).Sementes(n).Fila > 0 Then
                    Minutes = Fat(n) / (PlayerPlanet(i).SementesAcc + 1)
                    If PlayerPlanet(i).SementesAcc > 0 Then
                        If DateDiff("h", PlayerPlanet(i).SementesStart, Now) > 24 Then
                            PlayerPlanet(i).SementesAcc = 0
                        End If
                    End If
                    If DateDiff("n", PlayerPlanet(i).Sementes(n).Start, Now) >= Minutes Then
                        PlayerPlanet(i).Sementes(n).Fila = PlayerPlanet(i).Sementes(n).Fila - 1
                        PlayerPlanet(i).Sementes(n).Quant = PlayerPlanet(i).Sementes(n).Quant + 1
                        PlayerPlanet(i).Sementes(n).Start = Now
                        SavePlayerPlanet i
                    End If
                End If
                If PlayerPlanet(i).Soldados(n).Fila > 0 Then
                    Minutes = Fat(n) / (PlayerPlanet(i).SoldadosAcc + 1)
                    If PlayerPlanet(i).SoldadosAcc > 0 Then
                        If DateDiff("h", PlayerPlanet(i).SoldadosStart, Now) > 24 Then
                            PlayerPlanet(i).SoldadosAcc = 0
                        End If
                    End If
                    If DateDiff("n", PlayerPlanet(i).Soldados(n).Start, Now) >= Minutes Then
                        PlayerPlanet(i).Soldados(n).Fila = PlayerPlanet(i).Soldados(n).Fila - 1
                        PlayerPlanet(i).Soldados(n).Quant = PlayerPlanet(i).Soldados(n).Quant + 1
                        PlayerPlanet(i).Soldados(n).Start = Now
                        SavePlayerPlanet i
                    End If
                End If
            Next n
            End If
        Next i
    End If
End Sub

Function StartConstructResource(Index As Long, ResourceNum As Long, MapNum As Long, PlanetNum As Long, X As Long, Y As Long, Optional IsEvolution As Boolean = False) As Boolean
    Dim EvNum As Long
    EvNum = ResourceNum
    If NucleoLevel(MapNum) < Resource(EvNum).MinLevel Then
        PlayerMsg Index, "Você precisa de um centro nível: " & Resource(EvNum).MinLevel & " para esta construção!", brightred
        Exit Function
    End If
    If IsEvolution Then
        If HasItem(Index, MoedaZ) < Resource(EvNum).ECostGold Then
            PlayerMsg Index, "Você não tem moedas z suficiente!", brightred
            Exit Function
        End If
        If HasItem(Index, EspV) < Resource(EvNum).ECostRed Then
            PlayerMsg Index, "Você não tem especiaria vermelha suficiente!", brightred
            Exit Function
        End If
        If HasItem(Index, EspAz) < Resource(EvNum).ECostBlue Then
            PlayerMsg Index, "Você não tem especiaria azul suficiente!", brightred
            Exit Function
        End If
        If HasItem(Index, EspAm) < Resource(EvNum).ECostYellow Then
            PlayerMsg Index, "Você não tem especiaria amarela suficiente!", brightred
            Exit Function
        End If
    End If
    If AddSaibaman(Index, PlanetNum, MapNum, 1, EvNum, X, Y) Then
        If IsEvolution Then
            TakeInvItem Index, MoedaZ, Resource(EvNum).ECostGold
            TakeInvItem Index, EspV, Resource(EvNum).ECostRed
            TakeInvItem Index, EspAz, Resource(EvNum).ECostBlue
            TakeInvItem Index, EspAm, Resource(EvNum).ECostYellow
        End If
        If IsEvolution Then
            Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE
            If Resource(Map(MapNum).Tile(X, Y).data1).ItemReward > 0 Then RemoveExtrator PlanetNum, X, Y
        End If
        PlayerPlanet(PlanetNum).PlanetMap = Map(MapNum)
        PlayerMsg Index, "Tarefa iniciada com sucesso!", brightgreen
        SendSaibamans Index, PlanetNum
        SavePlayerPlanet PlanetNum
        MapCache_Create MapNum
        SendMap Index, MapNum
        CacheResources MapNum
        SendResourceCacheToMap MapNum, , True
        StartConstructResource = True
    End If
End Function

Function StartConstructNPC(ByVal Index As Long, MapNum As Long, PlanetNum As Long, X As Long, Y As Long, NpcNum As Long, Optional IsEvolution As Boolean = False) As Boolean
    EvNum = NpcNum
    If NucleoLevel(MapNum) < Npc(EvNum).MinLevel Then
        PlayerMsg Index, "Você precisa de um centro nível: " & Npc(EvNum).MinLevel & " para esta construção!", brightred
        Exit Function
    End If
    If IsEvolution Then
        If HasItem(Index, MoedaZ) < Npc(EvNum).ECostGold Then
            PlayerMsg Index, "Você não tem moedas z suficiente!", brightred
            Exit Function
        End If
        If HasItem(Index, EspV) < Npc(EvNum).ECostRed Then
            PlayerMsg Index, "Você não tem especiaria vermelha suficiente!", brightred
            Exit Function
        End If
        If HasItem(Index, EspAz) < Npc(EvNum).ECostBlue Then
            PlayerMsg Index, "Você não tem especiaria azul suficiente!", brightred
            Exit Function
        End If
        If HasItem(Index, EspAm) < Npc(EvNum).ECostYellow Then
            PlayerMsg Index, "Você não tem especiaria amarela suficiente!", brightred
            Exit Function
        End If
    End If
    If AddSaibaman(Index, PlanetNum, MapNum, 0, EvNum, X, Y) Then
        If IsEvolution Then
            TakeInvItem Index, MoedaZ, Npc(EvNum).ECostGold
            TakeInvItem Index, EspV, Npc(EvNum).ECostRed
            TakeInvItem Index, EspAz, Npc(EvNum).ECostBlue
            TakeInvItem Index, EspAm, Npc(EvNum).ECostYellow
        End If
        If IsEvolution Then
            Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE
            Map(MapNum).Npc(Map(MapNum).Tile(X, Y).data1) = 0
        End If
        PlayerPlanet(PlanetNum).PlanetMap = Map(MapNum)
        PlayerMsg Index, "Tarefa iniciada com sucesso!", brightgreen
        SendSaibamans Index, PlanetNum
        SavePlayerPlanet PlanetNum
        MapCache_Create MapNum
        SendMap Index, MapNum
        ' Respawn NPCS
        Dim i As Long
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, MapNum)
        Next
        SendMapNpcsToMap MapNum
        StartConstructNPC = True
    End If
End Function

Function NucleoLevel(ByVal MapNum As Long)
    Dim i As Long
    For i = 1 To ResourceCache(MapNum).Resource_Count
        If Resource(ResourceCache(MapNum).ResourceData(i).ResourceNum).NucleoLevel > 0 Then
            NucleoLevel = Resource(ResourceCache(MapNum).ResourceData(i).ResourceNum).NucleoLevel
            Exit Function
        End If
    Next i
    Dim PlanetNum As Long
    PlanetNum = PlayerMapIndex(MapNum)
    If PlanetNum > 0 Then
    For i = 1 To PlayerPlanet(PlanetNum).TotalSaibamans
        If PlayerPlanet(PlanetNum).Saibaman(i).Working = 1 Then
            If PlayerPlanet(PlanetNum).Saibaman(i).TaskType = 1 Then
                If Resource(PlayerPlanet(PlanetNum).Saibaman(i).TaskResult).NucleoLevel > 0 Then
                    NucleoLevel = Resource(PlayerPlanet(PlanetNum).Saibaman(i).TaskResult).NucleoLevel
                    Exit Function
                End If
            End If
        End If
    Next i
    End If
End Function

Function ResourceCount(ByVal MapNum As Long, ResourceNum As Long) As Long
    Dim i As Long
    For i = 1 To ResourceCache(MapNum).Resource_Count
        If Trim$(LCase(Resource(ResourceCache(MapNum).ResourceData(i).ResourceNum).Name)) = Trim$(LCase(Resource(ResourceNum).Name)) Then
            ResourceCount = ResourceCount + 1
        End If
    Next i
    Dim PlanetNum As Long
    PlanetNum = PlayerMapIndex(MapNum)
    For i = 1 To PlayerPlanet(PlanetNum).TotalSaibamans
        If PlayerPlanet(PlanetNum).Saibaman(i).Working = 1 Then
            If PlayerPlanet(PlanetNum).Saibaman(i).TaskType = 1 Then
                If Trim$(LCase(Resource(PlayerPlanet(PlanetNum).Saibaman(i).TaskResult).Name)) = Trim$(LCase(Resource(ResourceNum).Name)) Then ResourceCount = ResourceCount + 1
            End If
        End If
    Next i
End Function

Function Alloc(ByVal PlanetNum As Long) As Long
    Dim X As Long, MapNum As Long
    MapNum = PlayerPlanet(PlanetNum).PlanetData.Map
    If MapNum > 0 Then
        For X = 1 To ResourceCache(MapNum).Resource_Count
            If ResourceCache(MapNum).ResourceData(X).ResourceNum > 0 Then
                If Resource(ResourceCache(MapNum).ResourceData(X).ResourceNum).ToolRequired = 3 Then
                    Alloc = Alloc + (10 * Resource(ResourceCache(MapNum).ResourceData(X).ResourceNum).ResourceLevel)
                End If
            End If
        Next X
    End If
End Function

Function Allocated(ByVal PlanetNum As Long) As Long
    Dim i As Long
    For i = 1 To 5
        Allocated = Allocated + PlayerPlanet(PlanetNum).Soldados(i).Fila + PlayerPlanet(PlanetNum).Soldados(i).Quant
    Next i
End Function

Function CapacidadeMaxima(ByVal Level As Long) As Long
    CapacidadeMaxima = Level * 1000
End Function

Function GetEspeciariaPrice(ByVal Number As Byte) As Long
    Dim i As Long
    Dim Total As Long
    Dim Base(1 To 3) As Long
    Dim Abundance(1 To 3) As Double
    Dim Calc(1 To 3) As Double
    Dim Norm(1 To 3) As Double
    
    Base(1) = 80
    Base(2) = 160
    Base(3) = 320
    
    Abundance(1) = 57.48
    Abundance(2) = 28.74
    Abundance(3) = 13.8
    
    For i = 1 To 3
        If EspAmount(i) = 0 Then
            GetEspeciariaPrice = Base(Number)
            Exit Function
        Else
            Total = Total + EspAmount(i)
        End If
    Next i

    Calc(Number) = 100 - EspAmount(Number) * (100 / Total)
    Norm(Number) = 100 - (Calc(Number) + Abundance(Number))
    
    GetEspeciariaPrice = (Base(Number) / Abundance(Number)) * (Abundance(Number) - Norm(Number))
    If GetEspeciariaPrice < Base(Number) * 0.75 Then GetEspeciariaPrice = Base(Number) * 0.25
    If GetEspeciariaPrice > Base(Number) * 3 Then GetEspeciariaPrice = Base(Number) * 3
    
End Function
Sub RemoveTemporaryItems(ByVal Index As Long)
    Dim i As Long
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, i) = TesouroItem Then
            TakeInvItem Index, TesouroItem, GetPlayerInvItemValue(Index, i)
        End If
    Next i
End Sub
Sub CompleteService(ByVal Index As Long)
    Dim Gold As Long
    Dim Exp As Long
    Dim Level As Long
    
    Level = Planets(TempPlayer(Index).PlanetService).Level
    
    Gold = Planets(TempPlayer(Index).PlanetService).Preco * 0.5
    Exp = Int(ExperienceBase(Level) * 0.5)
    
    GiveInvItem Index, MoedaZ, Gold
    GivePlayerEXP Index, Exp
    
    SendPlaySound Index, "pagamento conclido.mp3"
    SendServiceComplete Index, Gold, Exp
    PlanetInService(TempPlayer(Index).PlanetService) = False
    TempPlayer(Index).PlanetService = 0
    Player(Index).NumServices = Player(Index).NumServices + 1
    TempPlayer(Index).OnlineServices = TempPlayer(Index).OnlineServices + 1
    If TempPlayer(Index).OnlineServices >= 3 Then
        TempPlayer(Index).OnlineServices = 0
        PlayerMsg Index, "Você recebeu: 1 Caixa da sorte!", Yellow
        GiveInvItem Index, 141, 1
    Else
        PlayerMsg Index, "Você completou " & TempPlayer(Index).OnlineServices & " serviços consecutivos, complete mais " & (3 - TempPlayer(Index).OnlineServices) & " e ganhe uma caixa surpresa!", Yellow
    End If
    SendPlayerData Index
End Sub

Function InLevel(Index As Long, Level As Long) As Boolean
    Select Case GetPlayerMap(Index)
        Case 1: If Level <= 25 Then InLevel = True
        Case 53: If Level > 25 And Level <= 50 Then InLevel = True
        Case 54: If Level > 50 Then InLevel = True
    End Select
End Function
