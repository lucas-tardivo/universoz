Attribute VB_Name = "modEvent"
Public LastEventTick As Long

Public MagicWord As String

Public InEvent As Byte

Public Const MaxAventuras As Byte = 1

Public AventuraNum As Long

'Recompensas
Public Recompensa(1 To AutomaticEvents.Count - 1) As RecompensasRec

'Config
Public AventuraConfig(1 To MaxAventuras) As AventuraRec

Type AventuraRec
    Nome As String
    MapNum As Long
    EventNum As Long
    X As Long
    Y As Long
End Type

Type RecompensasRec
    Reward(1 To 20) As PlayerInvRec
    TotalRewards As Byte
End Type

Sub LoadEventsRewards()
Dim Evento As Long

    'recompensas
    Evento = AutomaticEvents.PalavraMagica
    
    Call AddReward(Evento, 1, 10)
    
    'configs
    Call AddAventuraConfig(2, "Aventura do nada", 20, 19, 5)
    
End Sub

Sub AddReward(Evento As Long, ItemNum As Long, ItemValue As Long)
    Dim i As Long
    For i = 1 To 20
        If Recompensa(Evento).Reward(i).Num = 0 Then
            Recompensa(Evento).Reward(i).Num = ItemNum
            Recompensa(Evento).Reward(i).Value = ItemValue
            Recompensa(Evento).TotalRewards = i
            Exit Sub
        End If
    Next i
End Sub

Sub AddAventuraConfig(MapNum As Long, Nome As String, EventNum As Long, Optional X As Long = -1, Optional Y As Long = -1)
    Dim i As Long
    For i = 1 To MaxAventuras
        If AventuraConfig(i).MapNum = 0 Then
            AventuraConfig(i).MapNum = MapNum
            AventuraConfig(i).Nome = Nome
            AventuraConfig(i).EventNum = EventNum
            If X <> -1 Then
                AventuraConfig(i).X = X
            Else
                AventuraConfig(i).X = Map(MapNum).MaxX / 2
            End If
            If Y <> -1 Then
                AventuraConfig(i).Y = Y
            Else
                AventuraConfig(i).Y = Map(MapNum).MaxY / 2
            End If
            Exit Sub
        End If
    Next i
End Sub

Sub StartEvent()
    Dim EventType As Byte, i As Long, EventMsg As String
    
    EventType = rand(1, AutomaticEvents.Count - 1)
    
    EventMsg = "[EVENTO AUTOMÁTICO] "
    
    If EventType = AutomaticEvents.PalavraMagica Then  'palavra mágica
        MagicWord = vbNullString
        
        For i = 1 To rand(10, 20)
            MagicWord = MagicWord & Chr(rand(65, 90))
        Next i
        
        EventMsg = EventMsg & "Palavra mágica! o primeiro a digitar " & MagicWord & " NO GLOBAL vencerá o evento!"
    End If
    
    If EventType = AutomaticEvents.Aventura Then 'aventura
        AventuraNum = rand(1, MaxAventuras)
        
        EventMsg = EventMsg & "Aventura (" & AventuraConfig(AventuraNum).Nome & ") digite ENTRAR no global e participe!"
    End If
    
    Call GlobalMsg(EventMsg, Yellow)
    InEvent = EventType
End Sub

Sub GiveReward(ByVal Index As Long)
    Dim RewardNum As Long
    If InEvent = AutomaticEvents.PalavraMagica Then
        RewardNum = rand(1, Recompensa(InEvent).TotalRewards)
        GiveInvItem Index, Recompensa(InEvent).Reward(RewardNum).Num, Recompensa(InEvent).Reward(RewardNum).Value, True
        Call GlobalMsg(GetPlayerName(Index) & " venceu o evento da palavra mágica e recebeu " & Recompensa(InEvent).Reward(RewardNum).Value & " " & Trim(Item(Recompensa(InEvent).Reward(RewardNum).Num).Name) & "(s)!", Yellow)
    End If
End Sub

Sub StartGlobalEvent(EventNum As Long)
    EventGlobalType = EventNum
    If TotalOnlinePlayers < 5 Then Exit Sub
    
    frmServer.cmdIniciarEvent.Enabled = False
    
    If EventGlobalType = 1 Then
        'Freezer
        Dim NpcNum As Long
        Dim i As Long
        Dim ConsideredPlayers As Long
        NpcNum = 92
        
        Npc(NpcNum).HP = 0
        Npc(NpcNum).Damage = 0
        ConsideredPlayers = 0
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerAccess(i) <= 1 Then
                    Npc(NpcNum).HP = Npc(NpcNum).HP + GetPlayerDamage(i)
                    Npc(NpcNum).Damage = Npc(NpcNum).Damage + GetPlayerMaxVital(i, HP)
                    Npc(NpcNum).Level = Npc(NpcNum).Level + GetPlayerPDL(i)
                    ConsideredPlayers = ConsideredPlayers + 1
                End If
            End If
        Next i
        If ConsideredPlayers < 1 Then ConsideredPlayers = 1
        Npc(NpcNum).HP = Npc(NpcNum).HP * 150
        Npc(NpcNum).Damage = Int(Npc(NpcNum).Damage / ConsideredPlayers) / 100
        Npc(NpcNum).Level = Npc(NpcNum).Level / ConsideredPlayers
        If Npc(NpcNum).Damage < 0 Then Npc(NpcNum).Damage = 1
        
        Map(2).Npc(15) = NpcNum
        SpawnNpc 15, 2, True
        
        GlobalMsg "[EVENTO GLOBAL] O perverso Freezer, o Rei Gelado, está invadindo o planeta Vegeta! Ajude a conter esta ameaça!", Yellow
    End If
    
    If EventGlobalType = 2 Then
        For i = 15 To MAX_MAP_NPCS
            Map(2).Npc(i) = rand(93, 94)
            SpawnNpc i, 2, True
        Next i
        
        GlobalMsg "[EVENTO GLOBAL] Piratas espaciais estão tentando invadir o planeta Vegeta! Mostre para eles quem é que manda!", Yellow
    End If
    
    If EventGlobalType = 3 Then
        GlobalMsg "[EVENTO GLOBAL] O planeta Trash está liberado para conquistar suas riquezas! Vá e batalhe por elas!", Yellow
        CacheResources 49
    End If
End Sub

Sub DesactivateEvent()
    Dim i As Long

    frmServer.cmdIniciarEvent.Enabled = True
    
    If EventGlobalType = 1 Then
        DespawnNPC 2, 15
        Map(2).Npc(15) = 0
        GlobalMsg "[EVENTO GLOBAL] Parabéns! A invasão de Freezer foi contida!", Yellow
    End If
    
    If EventGlobalType = 2 Then
        For i = 15 To MAX_MAP_NPCS
            DespawnNPC 2, i
            Map(2).Npc(i) = 0
        Next i
        GlobalMsg "[EVENTO GLOBAL] Parabéns! A invasão dos piratas foi contida!", Yellow
    End If
    
    If EventGlobalType = 3 Then
        GlobalMsg "[EVENTO GLOBAL] O planeta Trash não está mais dando suas riquezas!", Yellow
    End If
    
    EventGlobalType = 0
End Sub
