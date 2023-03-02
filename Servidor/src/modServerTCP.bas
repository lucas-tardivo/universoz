Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    ' Update the form caption
    frmServer.Caption = "GoPlay Games - " & Options.Game_Name & " - " & TotalOnlinePlayers & "/" & Trim(STR(MAX_PLAYERS))
    
    ' Update form labels
    frmServer.lblPlayers = TotalOnlinePlayers & "/" & Trim(STR(MAX_PLAYERS))
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next

End Sub

Function IsConnected(ByVal Index As Long) As Boolean

    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean

    If IsConnected(Index) Then
        If LenB(Trim$(Player(Index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Function IsBanned(ByVal IP As String, Optional Name As String) As Boolean
    Dim filename As String
    Dim fIP As String
    Dim fName As String
    Dim F As Long
    filename = App.path & "\data\banlist.txt"

    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Input As #F

    Do While Not EOF(F)
        Input #F, fIP
        Input #F, fName

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #F
            Exit Function
        End If
        
        ' Is banned?
        If Trim$(LCase$(fName)) = Trim$(LCase$(Name)) Then
            IsBanned = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
        
        ' Add a packet to the packets/second number.
        PacketsOut = PacketsOut + 1
              
              On Error Resume Next
        frmServer.Socket(Index).SendData Buffer.ToArray()
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
            DoEvents
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                Call SendDataTo(i, Data)
                DoEvents
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
                DoEvents
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If i <> Index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If

    Next

End Sub

Sub SendDataToParty(ByVal PartyNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To Party(PartyNum).MemberCount
        If Party(PartyNum).Member(i) > 0 Then
            Call SendDataTo(Party(PartyNum).Member(i), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToAll Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, Buffer.ToArray
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Public Sub GuildMsg(ByVal GuildNum As Long, Msg As String, Color As Byte)
    Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Guild = GuildNum Then
                PlayerMsg i, "[GUILD] " & Msg, Color
            End If
        End If
    Next i
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataTo Index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToMap MapNum, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AlertMSG(ByVal Index As Long, ByVal Msg As String, Optional disconnect As Boolean = True)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    SendDataTo Index, Buffer.ToArray
    DoEvents
    If disconnect = True Then Call CloseSocket(Index)
    
    Set Buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal PartyNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim i As Long
    ' send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(PartyNum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(PartyNum).Member(i)) And IsPlaying(Party(PartyNum).Member(i)) Then
                PlayerMsg Party(PartyNum).Member(i), Msg, Color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)

    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMSG(Index, printf("Voce perdeu a conexão com %s.", Options.Game_Name))
    End If

End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (Index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub SocketConnected(ByVal Index As Long)
Dim i As Long

    If Index <> 0 Then
        ' make sure they're not banned
        If Not IsBanned(GetPlayerIP(Index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(Index) & ".", ChatSystem)
        Else
            Call AlertMSG(Index, "You have been banned from " & Options.Game_Name & ", and can no longer play.")
        End If
        ' re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        SendHighIndex
    End If
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    If GetPlayerAccess(Index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 1000 Then
            Exit Sub
        End If
    
        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 25 Then
            Exit Sub
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = GetTickCount + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(Index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(Index).Buffer.length >= 4 Then
        pLength = TempPlayer(Index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).Buffer.length - 4
        If pLength <= TempPlayer(Index).Buffer.length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).Buffer.ReadLong
            HandleData Index, TempPlayer(Index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(Index).Buffer.length >= 4 Then
            pLength = TempPlayer(Index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(Index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal Index As Long)

    If Index > 0 Then
        Call LeftGame(Index)
        Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.", ChatSystem)
        frmServer.Socket(Index).Close
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If

End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim X As Long
    Dim Y As Long
    Dim i As Long, z As Long, w As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong MapNum
    Buffer.WriteString Trim$(Map(MapNum).Name)
    Buffer.WriteString Trim$(Map(MapNum).Music)
    Buffer.WriteString Trim$(Map(MapNum).BGS)
    Buffer.WriteLong Map(MapNum).Revision
    Buffer.WriteByte Map(MapNum).Moral
    Buffer.WriteLong Map(MapNum).Up
    Buffer.WriteLong Map(MapNum).Down
    Buffer.WriteLong Map(MapNum).Left
    Buffer.WriteLong Map(MapNum).Right
    Buffer.WriteLong Map(MapNum).BootMap
    Buffer.WriteByte Map(MapNum).BootX
    Buffer.WriteByte Map(MapNum).BootY
    
    Buffer.WriteLong Map(MapNum).Weather
    Buffer.WriteLong Map(MapNum).WeatherIntensity
    
    Buffer.WriteLong Map(MapNum).Fog
    Buffer.WriteLong Map(MapNum).FogSpeed
    Buffer.WriteLong Map(MapNum).FogOpacity
    Buffer.WriteByte Map(MapNum).FogDir
    
    Buffer.WriteLong Map(MapNum).Red
    Buffer.WriteLong Map(MapNum).Green
    Buffer.WriteLong Map(MapNum).Blue
    Buffer.WriteLong Map(MapNum).Alpha
    
    Buffer.WriteByte Map(MapNum).MaxX
    Buffer.WriteByte Map(MapNum).MaxY

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY

            With Map(MapNum).Tile(X, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).X
                    Buffer.WriteLong .Layer(i).Y
                    Buffer.WriteLong .Layer(i).Tileset
                Next
                For z = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Autotile(z)
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .data1
                Buffer.WriteLong .data2
                Buffer.WriteLong .data3
                Buffer.WriteString .Data4
                Buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(MapNum).Npc(X)
        Buffer.WriteLong Map(MapNum).NpcSpawnType(X)
    Next
    
    Buffer.WriteLong Map(MapNum).Panorama
    
    Buffer.WriteByte Map(MapNum).Fly
    Buffer.WriteByte Map(MapNum).Ambiente
    
    MapCache(MapNum).Data = Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> Index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = printf("Não há jogadores online.")
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = printf("Existem %d jogadores online: ", Val(n)) & s & "."
    End If

    Call PlayerMsg(Index, s, WhoColor)
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
    Dim Buffer As clsBuffer, i As Long

    If Index > MAX_PLAYERS Then Exit Function
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerLevel(Index)
    Buffer.WriteLong GetPlayerPOINTS(Index)
    Buffer.WriteLong GetPlayerSprite(Index)
    Buffer.WriteLong GetPlayerMap(Index)
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteLong GetPlayerClass(Index)
    Buffer.WriteLong TempPlayer(Index).Trans
    Buffer.WriteLong Player(Index).PDL
    Buffer.WriteLong GetPlayerPDL(Index)
    Buffer.WriteLong Player(Index).EsoNum
    Buffer.WriteLong Player(Index).EsoTime
    Buffer.WriteByte TempPlayer(Index).Fly
    Buffer.WriteByte Player(Index).VIP
    Buffer.WriteByte Player(Index).Hair
    Buffer.WriteByte TempPlayer(Index).HairChange
    Buffer.WriteLong Player(Index).Titulo
    Buffer.WriteByte Player(Index).InTutorial
    Buffer.WriteLong Player(Index).Guild
    
    'AFK
    If isAFK(Index) Then
        Buffer.WriteByte 1
    Else
        Buffer.WriteByte 0
    End If
    
    Buffer.WriteLong TempPlayer(Index).Speed
    Buffer.WriteLong Player(Index).VIPExp
    Buffer.WriteLong TempPlayer(Index).PlanetService
    Buffer.WriteLong GetPlayerVipNextLevel(Index)
    Buffer.WriteLong TempPlayer(Index).Instance
    Buffer.WriteLong Player(Index).NumServices
    Buffer.WriteByte Player(Index).IsGod
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, i)
        Buffer.WriteLong GetPlayerStatPoints(Index, i)
    Next
    
    Buffer.WriteByte Player(Index).IsDead
    
    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing
End Function

Sub SendJoinMap(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> Index Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    SendDataTo Index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
    
    Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong Index
    SendDataToMapBut Index, MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(MapNum, i).PlayerName
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).X
        Buffer.WriteLong MapItem(MapNum, i).Y
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(MapNum, i).PlayerName
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).X
        Buffer.WriteLong MapItem(MapNum, i).Y
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal MapNum As Long, ByVal MapNPCNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcVitals
    Buffer.WriteLong MapNPCNum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNpc(MapNum).Npc(MapNPCNum).Vital(i)
    Next
    Buffer.WriteLong GetNpcMaxVital(MapNum, MapNPCNum, HP)

    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).X
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(HP)
        Buffer.WriteLong GetNpcMaxVital(MapNum, i, HP)
        Buffer.WriteLong MapNpc(MapNum).Npc(i).PDL
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).X
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(HP)
        Buffer.WriteLong GetNpcMaxVital(MapNum, i, HP)
        Buffer.WriteLong MapNpc(MapNum).Npc(i).PDL
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendItems(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If

    Next

End Sub

Sub SendAnimations(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(Index, i)
        End If

    Next

End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If

    Next

End Sub

Sub SendQuests(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS

        If LenB(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdatequestTo(Index, i)
        End If

    Next

End Sub

Sub SendGuilds(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_GUILDS

        If LenB(Trim$(Guild(i).Name)) > 0 Then
            Call SendUpdateGuildTo(Index, i)
        End If

    Next

End Sub

Sub SendResources(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(Index, i)
        End If

    Next

End Sub

Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(Index, i)
        Buffer.WriteLong GetPlayerInvItemValue(Index, i)
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal invSlot As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong invSlot
    Buffer.WriteLong GetPlayerInvItemNum(Index, invSlot)
    Buffer.WriteLong GetPlayerInvItemValue(Index, invSlot)
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerWornEq
    Buffer.WriteLong GetPlayerEquipment(Index, Armor)
    Buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(Index, helmet)
    Buffer.WriteLong GetPlayerEquipment(Index, shield)
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerEquipment(Index, Armor)
    Buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(Index, helmet)
    Buffer.WriteLong GetPlayerEquipment(Index, shield)
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Weapon)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, helmet)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, shield)
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals, Optional Receiver As Long = 0)
    Dim packet As String
    Dim Buffer As clsBuffer
    If Receiver = 0 Then Receiver = Index
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong Index
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
    End Select

    SendDataTo Receiver, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendEXP(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(Index)
    Buffer.WriteLong GetPlayerNextLevel(Index)
    Buffer.WriteLong GetPlayerLastLevel(Index)
    Buffer.WriteLong GetPlayerPDL(Index)
    Buffer.WriteLong GetPlayerLevel(Index)
    Buffer.WriteByte Player(Index).IsGod
    Buffer.WriteLong Player(Index).GodLevel
    Buffer.WriteLong Player(Index).GodExp
    Buffer.WriteLong GetPlayerGodNextLevel(Index)
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStats(ByVal Index As Long)
Dim i As Long
Dim packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, i)
        Buffer.WriteLong GetPlayerStatPoints(Index, i)
        Buffer.WriteLong GetPlayerStatNextLevel(Index, i)
        Buffer.WriteLong GetPlayerStatLastLevel(Index, i)
    Next
    Buffer.WriteLong Player(Index).Points
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendWelcome(ByVal Index As Long)
    Call PlayerMsg(Index, Options.MOTD, Yellow)
    Call PlayerMsg(Index, printf("Digite /ajuda para saber os comandos do chat!"), Yellow)
    'Call PlayerMsg(Index, "[EVENTO] Faça um vídeo do nosso jogo para ganhar chaves para seus amigos e concorrer á prêmios exclusivos! Para mais informações acesse: www.goplaygames.com.br", Yellow)
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClassesData
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).stat(q)
        Next
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).stat(q)
        Next
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLeftGame(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    SendDataToAllBut Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal Index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set Buffer = New clsBuffer
    NPCSize = LenB(Npc(NpcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(NpcNum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteBytes NPCData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set Buffer = New clsBuffer
    NPCSize = LenB(Npc(NpcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(NpcNum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteBytes NPCData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdatequestToAll(ByVal questnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim questSize As Long
    Dim questData() As Byte
    Set Buffer = New clsBuffer
    questSize = LenB(Quest(questnum))
    ReDim questData(questSize - 1)
    CopyMemory questData(0), ByVal VarPtr(Quest(questnum)), questSize
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong questnum
    Buffer.WriteBytes questData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdatequestTo(ByVal Index As Long, ByVal questnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim questSize As Long
    Dim questData() As Byte
    Set Buffer = New clsBuffer
    questSize = LenB(Quest(questnum))
    ReDim questData(questSize - 1)
    CopyMemory questData(0), ByVal VarPtr(Quest(questnum)), questSize
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong questnum
    Buffer.WriteBytes questData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateGuildToAll(ByVal GuildNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim GuildSize As Long
    Dim GuildData() As Byte
    Set Buffer = New clsBuffer
    GuildSize = LenB(Guild(GuildNum))
    ReDim GuildData(GuildSize - 1)
    CopyMemory GuildData(0), ByVal VarPtr(Guild(GuildNum)), GuildSize
    Buffer.WriteLong SUpdateGuild
    Buffer.WriteLong GuildNum
    Buffer.WriteBytes GuildData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateGuildTo(ByVal Index As Long, ByVal GuildNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim GuildSize As Long
    Dim GuildData() As Byte
    Set Buffer = New clsBuffer
    GuildSize = LenB(Guild(GuildNum))
    ReDim GuildData(GuildSize - 1)
    CopyMemory GuildData(0), ByVal VarPtr(Guild(GuildNum)), GuildSize
    Buffer.WriteLong SUpdateGuild
    Buffer.WriteLong GuildNum
    Buffer.WriteBytes GuildData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal Index As Long, ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShops(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If

    Next

End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If

    Next
    
    Call SendPlayerSpells(Index)

End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong GetPlayerSpell(Index, i)
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal Index As Long, Optional Resource_num As Long = 0)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong Resource_num
    
    If Resource_num = 0 Then
        Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).Resource_Count

        If ResourceCache(GetPlayerMap(Index)).Resource_Count > 0 Then
    
            For i = 1 To ResourceCache(GetPlayerMap(Index)).Resource_Count
                Buffer.WriteByte ResourceCache(GetPlayerMap(Index)).ResourceData(i).ResourceState
                Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).X
                Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).Y
                If Resource(ResourceCache(GetPlayerMap(Index)).ResourceData(i).ResourceNum).ResourceType = 4 Then   'Extrator
                    Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).ResourceNum
                Else
                    Buffer.WriteLong 0
                End If
            Next
    
        End If
    Else
        Buffer.WriteByte ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState
        Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).X
        Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).Y
        If Resource(ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceNum).ResourceType = 4 Then 'Extrator
            Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceNum
        Else
            Buffer.WriteLong 0
        End If
    End If

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal MapNum As Long, Optional Resource_num As Long = 0, Optional ForceUpdate As Boolean = False)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong Resource_num

    If Resource_num = 0 Then
        Buffer.WriteLong ResourceCache(MapNum).Resource_Count
        If ResourceCache(MapNum).Resource_Count > 0 Then
    
            For i = 1 To ResourceCache(MapNum).Resource_Count
                Buffer.WriteByte ResourceCache(MapNum).ResourceData(i).ResourceState
                Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).X
                Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).Y
                If Resource(ResourceCache(MapNum).ResourceData(i).ResourceNum).ResourceType = 4 Or ForceUpdate Then  'Extrator
                    Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).ResourceNum
                Else
                    Buffer.WriteLong 0
                End If
            Next
    
        End If
    Else
        Buffer.WriteByte ResourceCache(MapNum).ResourceData(Resource_num).ResourceState
        Buffer.WriteLong ResourceCache(MapNum).ResourceData(Resource_num).X
        Buffer.WriteLong ResourceCache(MapNum).ResourceData(Resource_num).Y
        If Resource(ResourceCache(MapNum).ResourceData(Resource_num).ResourceNum).ResourceType = 4 Or ForceUpdate Then  'Extrator
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(Resource_num).ResourceNum
        Else
            Buffer.WriteLong 0
        End If
    End If

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendActionMsg(ByVal MapNum As Long, ByVal message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal X As Long, ByVal Y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString message
    Buffer.WriteLong Color
    Buffer.WriteLong MsgType
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendBossMsg(ByVal MapNum As Long, ByVal message As String, ByVal Color As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString message
    Buffer.WriteLong Color
    Buffer.WriteLong 5
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendBlood(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendAnimation(ByVal MapNum As Long, ByVal Anim As Long, ByVal X As Long, ByVal Y As Long, ByVal Dir As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0, Optional ByVal OnlyTo As Long = 0, Optional IsLinear As Byte = 0, Optional LockToNpc As Byte = 0, Optional CastAnim As Byte = 0, Optional But As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    Buffer.WriteByte Dir
    Buffer.WriteByte IsLinear
    Buffer.WriteByte LockToNpc
    Buffer.WriteByte CastAnim
    
    If OnlyTo > 0 Then
        SendDataTo OnlyTo, Buffer.ToArray
    Else
        If But = 0 Then
            SendDataToMap MapNum, Buffer.ToArray()
        Else
            SendDataToMapBut But, MapNum, Buffer.ToArray()
        End If
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendCooldown(ByVal Index As Long, ByVal slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong slot
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    Buffer.WriteLong Index
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal MapNum As Long, ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub ResetShopAction(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendStunned(ByVal Index As Long, TargetTyp As Byte, Duration As Long, MapNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong Index
    Buffer.WriteByte TargetTyp
    Buffer.WriteLong Duration
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendBank(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Bank(Index).Item(i).Num
        Buffer.WriteLong Bank(Index).Item(i).Value
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal Index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopNum
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal Index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendTrade(ByVal Index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal Index As Long, ByVal dataType As Byte)
Dim Buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    
    tradeTarget = TempPlayer(Index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(Index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(Index).TradeOffer(i).Num).Type = ITEM_TYPE_CURRENCY Or Item(TempPlayer(Index).TradeOffer(i).Num).Stackable > 0 Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).Price * TempPlayer(Index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Stackable > 0 Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    Buffer.WriteLong totalWorth
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTarget(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    Buffer.WriteLong TempPlayer(Index).Target
    Buffer.WriteLong TempPlayer(Index).TargetType
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(Index).Hotbar(i).slot
        Buffer.WriteByte Player(Index).Hotbar(i).sType
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLoginOk(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong Index
    Buffer.WriteLong Player_HighIndex
    Buffer.WriteLong MAX_LEVELS
    Buffer.WriteLong MoedaZ
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInGame(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHighIndex()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighIndex
    Buffer.WriteLong Player_HighIndex
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal Index As Long, ByVal TargetPlayer As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(TargetPlayer).Name)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal PartyNum As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(PartyNum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(PartyNum).Member(i)
    Next
    Buffer.WriteLong Party(PartyNum).MemberCount
    SendDataToParty PartyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal Index As Long)
Dim Buffer As clsBuffer, i As Long, PartyNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    PartyNum = TempPlayer(Index).inParty
    If PartyNum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(PartyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(PartyNum).Member(i)
        Next
        Buffer.WriteLong Party(PartyNum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal PartyNum As Long, ByVal Index As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong Index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(Index, i)
        Buffer.WriteLong Player(Index).Vital(i)
    Next
    SendDataToParty PartyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal MapNum As Long, ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong Index
    Buffer.WriteString MapItem(MapNum, Index).PlayerName
    Buffer.WriteLong MapItem(MapNum, Index).Num
    Buffer.WriteLong MapItem(MapNum, Index).Value
    Buffer.WriteLong MapItem(MapNum, Index).X
    Buffer.WriteLong MapItem(MapNum, Index).Y
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendChatBubble(ByVal MapNum As Long, ByVal Target As Long, ByVal TargetType As Long, ByVal message As String, ByVal Colour As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatBubble
    Buffer.WriteLong Target
    Buffer.WriteLong TargetType
    Buffer.WriteString message
    Buffer.WriteLong Colour
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpecialEffect(ByVal Index As Long, EffectType As Long, Optional data1 As Long = 0, Optional data2 As Long = 0, Optional data3 As Long = 0, Optional Data4 As Long = 0)
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpecialEffect
    
    Select Case EffectType
        Case EFFECT_TYPE_FADEIN
            Buffer.WriteLong EffectType
        Case EFFECT_TYPE_FADEOUT
            Buffer.WriteLong EffectType
        Case EFFECT_TYPE_FLASH
            Buffer.WriteLong EffectType
        Case EFFECT_TYPE_FOG
            Buffer.WriteLong EffectType
            Buffer.WriteLong data1 'fognum
            Buffer.WriteLong data2 'fog movement speed
            Buffer.WriteLong data3 'opacity
        Case EFFECT_TYPE_WEATHER
            Buffer.WriteLong EffectType
            Buffer.WriteLong data1 'weather type
            Buffer.WriteLong data2 'weather intensity
        Case EFFECT_TYPE_TINT
            Buffer.WriteLong EffectType
            Buffer.WriteLong data1 'red
            Buffer.WriteLong data2 'green
            Buffer.WriteLong data3 'blue
            Buffer.WriteLong Data4 'alpha
    End Select
    
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing
End Sub


Sub SendAttack(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SAttack
    Buffer.WriteLong Index
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendFlash(ByVal Target As Long, MapNum As Long, isNpc As Boolean)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SFlash
    Buffer.WriteLong Target
    If isNpc Then
        Buffer.WriteByte 1
    Else
        Buffer.WriteByte 0
    End If
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSwitchesAndVariables(Index As Long, Optional everyone As Boolean = False)
Dim Buffer As clsBuffer, i As Long

Set Buffer = New clsBuffer
Buffer.WriteLong SSwitchesAndVariables

For i = 1 To MAX_SWITCHES
    Buffer.WriteString Switches(i)
Next

For i = 1 To MAX_VARIABLES
    Buffer.WriteString Variables(i)
Next

If everyone Then
    SendDataToAll Buffer.ToArray
Else
    SendDataTo Index, Buffer.ToArray
End If

Set Buffer = Nothing
End Sub
Sub SendProjectile(ByVal MapNum As Long, ByVal Attacker As Long, ByVal victim As Long, ByVal TargetType As Long, ByVal Graphic As Long, ByVal Rotate As Long, ByVal RotateSpeed As Long, ByVal NPCAttack As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Call Buffer.WriteLong(SCreateProjectile)
    Call Buffer.WriteLong(Attacker)
    Call Buffer.WriteLong(victim)
    Call Buffer.WriteLong(TargetType)
    Call Buffer.WriteLong(Graphic)
    Call Buffer.WriteLong(Rotate)
    Call Buffer.WriteLong(RotateSpeed)
    Call Buffer.WriteByte(NPCAttack)
    Call SendDataToMap(MapNum, Buffer.ToArray())
    
    Set Buffer = Nothing
End Sub

Public Sub Events_SendEventData(ByVal pIndex As Long, ByVal EIndex As Long)
    If pIndex <= 0 Or pIndex > MAX_PLAYERS Then Exit Sub
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Dim Buffer As clsBuffer
    Dim i As Long, D As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SEventData
    Buffer.WriteLong EIndex
    Buffer.WriteString Events(EIndex).Name
    Buffer.WriteByte Events(EIndex).chkSwitch
    Buffer.WriteByte Events(EIndex).chkVariable
    Buffer.WriteByte Events(EIndex).chkHasItem
    Buffer.WriteLong Events(EIndex).SwitchIndex
    Buffer.WriteByte Events(EIndex).SwitchCompare
    Buffer.WriteLong Events(EIndex).VariableIndex
    Buffer.WriteByte Events(EIndex).VariableCompare
    Buffer.WriteLong Events(EIndex).VariableCondition
    Buffer.WriteLong Events(EIndex).HasItemIndex
    If Events(EIndex).HasSubEvents Then
        Buffer.WriteLong UBound(Events(EIndex).SubEvents)
        For i = 1 To UBound(Events(EIndex).SubEvents)
            With Events(EIndex).SubEvents(i)
                Buffer.WriteLong .Type
                If .HasText Then
                    Buffer.WriteLong UBound(.Text)
                    For D = 1 To UBound(.Text)
                        Buffer.WriteString .Text(D)
                    Next D
                Else
                    Buffer.WriteLong 0
                End If
                If .HasData Then
                    Buffer.WriteLong UBound(.Data)
                    For D = 1 To UBound(.Data)
                        Buffer.WriteLong .Data(D)
                    Next D
                Else
                    Buffer.WriteLong 0
                End If
            End With
        Next i
    Else
        Buffer.WriteLong 0
    End If
    
    Buffer.WriteByte Events(EIndex).Trigger
    Buffer.WriteByte Events(EIndex).WalkThrought
    
    SendDataTo pIndex, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub Events_SendEventUpdate(ByVal pIndex As Long, ByVal EIndex As Long, ByVal SIndex As Long)
    If pIndex <= 0 Or pIndex > MAX_PLAYERS Then Exit Sub
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    If Not (Events(EIndex).HasSubEvents) Then Exit Sub
    If SIndex <= 0 Or SIndex > UBound(Events(EIndex).SubEvents) Then Exit Sub
    
    Dim Buffer As clsBuffer
    Dim D As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventUpdate
    Buffer.WriteLong SIndex
    With Events(EIndex).SubEvents(SIndex)
        Buffer.WriteLong .Type
        If .HasText Then
            Buffer.WriteLong UBound(.Text)
            For D = 1 To UBound(.Text)
                Buffer.WriteString .Text(D)
            Next D
        Else
            Buffer.WriteLong 0
        End If
        If .HasData Then
            Buffer.WriteLong UBound(.Data)
            For D = 1 To UBound(.Data)
                Buffer.WriteLong .Data(D)
            Next D
        Else
            Buffer.WriteLong 0
        End If
    End With
    
    SendDataTo pIndex, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub Events_SendEventQuit(ByVal pIndex As Long)
    If pIndex <= 0 Or pIndex > MAX_PLAYERS Then Exit Sub
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventUpdate
    Buffer.WriteLong 1          'Current Event
    Buffer.WriteLong Evt_Quit   'Quit Event Type
    Buffer.WriteLong 0          'Text Count
    Buffer.WriteLong 0          'Data Count
    
    SendDataTo pIndex, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Sub SendMapKey(ByVal Index As Long, ByVal Value As Byte, ByVal EventNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteByte Value
    Buffer.WriteLong EventNum
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal MapNum As Long, ByVal Value As Byte, ByVal EventNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteByte Value
    Buffer.WriteLong EventNum
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub
Sub SendMapKeyToAll(ByVal Value As Byte, ByVal EventNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteByte Value
    Buffer.WriteLong EventNum
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendEffects(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_EFFECTS

        If LenB(Trim$(Effect(i).Name)) > 0 Then
            Call SendUpdateEffectTo(Index, i)
        End If

    Next

End Sub

Sub SendUpdateEffectToAll(ByVal EffectNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim EffectSize As Long
    Dim EffectData() As Byte
    Set Buffer = New clsBuffer
    EffectSize = LenB(Effect(EffectNum))
    ReDim EffectData(EffectSize - 1)
    CopyMemory EffectData(0), ByVal VarPtr(Effect(EffectNum)), EffectSize
    Buffer.WriteLong SUpdateEffect
    Buffer.WriteLong EffectNum
    Buffer.WriteBytes EffectData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateEffectTo(ByVal Index As Long, ByVal EffectNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim EffectSize As Long
    Dim EffectData() As Byte
    Set Buffer = New clsBuffer
    EffectSize = LenB(Effect(EffectNum))
    ReDim EffectData(EffectSize - 1)
    CopyMemory EffectData(0), ByVal VarPtr(Effect(EffectNum)), EffectSize
    Buffer.WriteLong SUpdateEffect
    Buffer.WriteLong EffectNum
    Buffer.WriteBytes EffectData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendEffect(ByVal MapNum As Long, ByVal Effect As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal OnlyTo As Long = 0, Optional But As Long = 0)
    Dim Buffer As clsBuffer
    If Effect = 0 Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEffect
    Buffer.WriteLong Effect
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    If OnlyTo > 0 Then
        SendDataTo OnlyTo, Buffer.ToArray
    Else
        If But = 0 Then
            SendDataToMap MapNum, Buffer.ToArray()
        Else
            SendDataToMapBut But, MapNum, Buffer.ToArray()
        End If
    End If
    
    Set Buffer = Nothing
End Sub

Public Sub SendNews(ByVal Index As Long, openNews As Byte)
Dim News As String, F As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSendNews
    Buffer.WriteByte openNews
    F = FreeFile
    Open App.path & "\data\news.txt" For Input As #F
        Line Input #F, News
    Close #F
    Buffer.WriteString News
    
    SendDataTo Index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerFly(ByVal Index As Long)
    SendPlayerData Index
End Sub

Public Sub SendAction(ByVal Index As Long, Action As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSpecialAction
    Buffer.WriteString Action
    
    SendDataTo Index, Buffer.ToArray
    
    Set Buffer = Nothing
    TextAdd "Ação (" & Action & ") executada com sucesso no jogador " & GetPlayerName(Index) & "!", 4
End Sub

Public Sub SendSpellBuffer(ByVal Index As Long, ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSpellBuffer
    Buffer.WriteLong Index
    Buffer.WriteLong TempPlayer(Index).spellBuffer.Spell
    Buffer.WriteLong SpellNum
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub SendShenlong(ByVal MapNum As Long, ByVal Active As Byte, Animation As Byte, Optional Index As Long = 0, Optional X As Long, Optional Y As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SShenlong
    Buffer.WriteLong MapNum
    Buffer.WriteByte Active
    Buffer.WriteByte Animation
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    If Index = 0 Then
        SendDataToAll Buffer.ToArray
    Else
        SendDataTo Index, Buffer.ToArray
    End If
    
    Set Buffer = Nothing
End Sub

Public Sub SendTransporteCome(TransporteNum As Byte, Optional Anim As Byte = 1)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
        Buffer.WriteLong STransporte
        Buffer.WriteByte TransporteNum
        Buffer.WriteLong Transporte(TransporteNum).Map
        Buffer.WriteByte Anim
            
        SendDataToMap Transporte(TransporteNum).Map, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub SendTransporteComeTo(TransporteNum As Byte, ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
        Buffer.WriteLong STransporte
        Buffer.WriteByte TransporteNum
        Buffer.WriteLong Transporte(TransporteNum).Map
        Buffer.WriteByte 1
            
        SendDataTo Index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub SendPlaySound(ByVal Index As Long, Sound As String, Optional ToMap As Boolean = False)
    Dim Buffer As clsBuffer, i As Long
    Set Buffer = New clsBuffer
        Buffer.WriteLong SPlaySound
        Buffer.WriteString Trim$(Sound)
        If ToMap = False Then
            SendDataTo Index, Buffer.ToArray
        Else
            For i = 1 To Player_HighIndex
                If Player(i).Map = Player(Index).Map Then
                    If Player(i).X >= Player(Index).X - 4 And Player(i).X <= Player(Index).X + 4 Then
                        If Player(i).Y >= Player(Index).Y - 4 And Player(i).Y <= Player(Index).Y + 4 Then
                            SendDataTo i, Buffer.ToArray
                        End If
                    End If
                End If
            Next i
        End If
    Set Buffer = Nothing
End Sub

Sub SendMapNpcXY(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcDataXY
    
    Buffer.WriteLong Index
    Buffer.WriteLong MapNpc(MapNum).Npc(Index).X
    Buffer.WriteLong MapNpc(MapNum).Npc(Index).Y

    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerQuests(ByVal Index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerQuests
    
    For i = 1 To MAX_QUESTS
        Buffer.WriteByte Player(Index).QuestState(i).State
    Next i

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerQuest(ByVal Index As Long, ByVal questnum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerQuest
    Buffer.WriteLong questnum
    Buffer.WriteByte Player(Index).QuestState(questnum).State
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerDailyQuest(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.PlayerDaily
    Buffer.WriteString printf(DailyMission(Player(Index).Daily.MissionIndex).Description, STR(Player(Index).Daily.MissionObjective))
    Buffer.WriteLong Player(Index).Daily.MissionActual
    Buffer.WriteByte Player(Index).Daily.Completed
    Buffer.WriteByte Player(Index).Daily.DailyBonus
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerAFK(ByVal Index As Long, ByVal AFK As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.AFK
    Buffer.WriteLong Index
    Buffer.WriteByte AFK
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendFishingTime(ByVal Index As Long)
    Dim Tick As Long
    
    Tick = rand(15000, 45000)
    TempPlayer(Index).NextFish = GetTickCount + Tick
    
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.Fish
    Buffer.WriteLong Tick
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendOpenRefine(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenRefine
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendGravidade(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.Gravidade
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDialogue(ByVal Index As Long, ByVal Title As String, ByVal Msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.GravityOk
    Buffer.WriteString Title
    Buffer.WriteString Msg
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendProvacaoState(ByVal Index As Long, ByVal State As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.ProvacaoInit
    Buffer.WriteByte State
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendOpenGuildMaster(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.OpenGuild
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendGuildInvite(ByVal Index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.GuildInvite
    Buffer.WriteString Msg
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendOpenArena(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.OpenArena
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendArenaChallenge(ByVal Index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.ArenaChallenging
    Buffer.WriteString Msg
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAntiHack(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.AntiHackData
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendFabrica(ByVal Index As Long, ByVal PlanetNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.FabricaData
    Dim i As Long
    For i = 1 To 5
        Buffer.WriteLong PlayerPlanet(PlanetNum).Sementes(i).Quant
        Buffer.WriteLong PlayerPlanet(PlanetNum).Sementes(i).Fila
    Next i
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSoldados(ByVal Index As Long, ByVal PlanetNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.ExercitoData
    Dim i As Long
    For i = 1 To 5
        Buffer.WriteLong PlayerPlanet(PlanetNum).Soldados(i).Quant
        Buffer.WriteLong PlayerPlanet(PlanetNum).Soldados(i).Fila
    Next i
    For i = 1 To 5
        Buffer.WriteLong PlayerPlanet(PlanetNum).Sementes(i).Quant
    Next i
    Buffer.WriteLong Alloc(PlanetNum)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendConfirmation(ByVal Index As Long, ByVal Msg As String, ByVal ConfirmIndex As Long)
    Dim Buffer As clsBuffer
    TempPlayer(Index).Confirmation = ConfirmIndex
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.Confirmation
    Buffer.WriteString Msg
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendConquistas(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SConquistas

    If Not IsConquistaEmpty Then
        Dim i As Long
        Dim n As Long
    
        Buffer.WriteLong UBound(Conquistas)
        For i = 1 To UBound(Conquistas)
            Buffer.WriteString Conquistas(i).Name
            Buffer.WriteString Conquistas(i).Desc
            Buffer.WriteLong Conquistas(i).Exp
            Buffer.WriteLong Conquistas(i).Progress
            For n = 1 To 5
                Buffer.WriteLong Conquistas(i).Reward(n).Num
                Buffer.WriteLong Conquistas(i).Reward(n).Value
            Next n
        Next i
    Else
        Buffer.WriteLong 0
    End If
    
    SendDataTo Index, Buffer.ToArray()
        
    Set Buffer = Nothing
End Sub

Sub SendPlayerConquistas(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.ConquistasInfo
    
    If Not IsConquistaEmpty Then
        Buffer.WriteLong UBound(Conquistas)
        For i = 1 To UBound(Conquistas)
            Buffer.WriteByte Player(Index).Conquistas(i)
            Buffer.WriteLong Player(Index).ConquistaProgress(i)
        Next i
    Else
        Buffer.WriteLong 0
    End If
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerConquista(ByVal Index As Long, ByVal Conquista As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.ConquistaInfo
    Buffer.WriteLong Conquista
    Buffer.WriteByte Player(Index).Conquistas(Conquista)
    Buffer.WriteLong Player(Index).ConquistaProgress(Conquista)
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendOpenTroca(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.OpenTroca
    
    For i = 1 To 3
        Buffer.WriteLong EspAmount(i)
        Buffer.WriteLong GetEspeciariaPrice(i)
    Next i
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendServiceComplete(ByVal Index As Long, ByVal Gold As Long, ByVal Exp As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerInfo
    Buffer.WriteByte PlayerInfoType.ServiceFeedback
    Buffer.WriteLong Gold
    Buffer.WriteLong Exp
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSupportMsg(ByVal Index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSupport
    Buffer.WriteString Msg
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
