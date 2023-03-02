Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            GetPlayerMaxVital = 100 + ((GetPlayerLevel(Index) * (1 + (GetPlayerStat(Index, Endurance, True) * 0.1))) + (GetPlayerStat(Index, Endurance, True) * 3))
        Case MP
            GetPlayerMaxVital = 100 + ((GetPlayerLevel(Index) * 1) + Int(GetPlayerStat(Index, Intelligence, True) * 2))
    End Select
    
    If Player(Index).IsGod > 0 Then
        GetPlayerMaxVital = (GetPlayerMaxVital / 100) * (100 + (Player(Index).IsGod * 20))
    End If
End Function

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If
    
    If TempPlayer(Index).Trans > 0 Then
        If Spell(TempPlayer(Index).Trans).TransVital(Vital) > 0 Then
            i = 0
            Exit Function
        End If
    End If

    Select Case Vital
        Case HP
            i = (GetPlayerMaxVital(Index, HP)) * 0.1
        Case MP
            i = (GetPlayerMaxVital(Index, MP)) * 0.05
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
    If GetPlayerVitalRegen = 0 Then GetPlayerVitalRegen = 1
    
    If Player(Index).IsDead = 1 Then GetPlayerVitalRegen = 0
End Function

Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(Index, Weapon)
        If Item(weaponNum).data3 > 0 Then
            GetPlayerDamage = (GetPlayerStat(Index, Strength)) + 8
        Else
            GetPlayerDamage = (GetPlayerStat(Index, Strength)) + (Item(weaponNum).data2) + 8
        End If
    Else
        GetPlayerDamage = (GetPlayerStat(Index, Strength)) + 8
    End If
    
    If Player(Index).IsGod > 0 Then
        GetPlayerDamage = (GetPlayerDamage / 100) * (100 + (Player(Index).IsGod * 20))
    End If


End Function
Function GetPlayerDefence(ByVal Index As Long) As Long
    Dim DefNum As Long
    Dim Def As Long
    Dim i As Long
    
    GetPlayerDefence = 0
    Def = 0
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerEquipment(Index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(Index, Armor)
        Def = Def + Item(DefNum).data2
    End If
    
    If GetPlayerEquipment(Index, helmet) > 0 Then
        DefNum = GetPlayerEquipment(Index, helmet)
        Def = Def + Item(DefNum).data2
    End If
    
    If GetPlayerEquipment(Index, shield) > 0 Then
        DefNum = GetPlayerEquipment(Index, shield)
        Def = Def + Item(DefNum).data2
    End If
    
   If Not GetPlayerEquipment(Index, Armor) > 0 And Not GetPlayerEquipment(Index, helmet) > 0 And Not GetPlayerEquipment(Index, shield) > 0 Then
        GetPlayerDefence = 0
    Else
        GetPlayerDefence = (Def * 0.5) + ((Def / 100) * rand(0, 50)) '(GetPlayerStat(Index, Endurance) * 4) + ((Def * GetPlayerStat(Index, Endurance)) * 0.8)
    End If
    
    If Player(Index).IsGod > 0 Then
        GetPlayerDefence = (GetPlayerDefence / 100) * (100 + (Player(Index).IsGod * 20))
    End If
End Function

Function GetNpcMaxVital(ByVal MapNum As Long, ByVal MapNPCNum As Long, ByVal Vital As Vitals) As Long
    Dim X As Long

    ' Prevent subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If
    
    ' Prevent subscript out of range
    If MapNpc(MapNum).Npc(MapNPCNum).Num <= 0 Or MapNpc(MapNum).Npc(MapNPCNum).Num > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).ND = 0 Or Not UZ Then
                GetNpcMaxVital = Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).HP
            Else
                GetNpcMaxVital = (NPCBase(MapNpc(MapNum).Npc(MapNPCNum).Level).HP / 100) * Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).HP
            End If
        Case MP
            GetNpcMaxVital = 30 + (Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).stat(Intelligence) * 10) + 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (Npc(NpcNum).stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (Npc(NpcNum).stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal MapNum As Long, ByVal MapNPCNum As Long) As Long
    If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).ND = 0 Or Not UZ Then
        GetNpcDamage = Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Damage
    Else
        GetNpcDamage = (NPCBase(MapNpc(MapNum).Npc(MapNPCNum).Level).Damage / 100) * Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Damage
    End If
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(Index, agility) / 52.08
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodgePlayer(ByVal victim As Long, ByVal Attacker As Long) As Boolean
Dim accuracyrate As Long, evasionrate As Long

    CanPlayerDodgePlayer = False

    'automatic
    If rand(1, 100) <= 25 Then
        Exit Function
    End If
    
    accuracyrate = (GetPlayerAccuracy(Attacker) / 100) * rand(1, 100)
    evasionrate = (GetPlayerEvasion(victim) / 100) * rand(1, 100)
    
    If accuracyrate >= evasionrate Then
        CanPlayerDodgePlayer = False
    Else
        CanPlayerDodgePlayer = True
    End If
End Function

Public Function CanPlayerDodgeNPC(ByVal victim As Long, ByVal NpcNum As Long) As Boolean
Dim accuracyrate As Long, evasionrate As Long

    CanPlayerDodgeNPC = False

    'automatic
    If rand(1, 100) <= 25 Then
        Exit Function
    End If
    
    If Npc(MapNpc(GetPlayerMap(victim)).Npc(NpcNum).Num).ND = 0 Or Not UZ Then
        accuracyrate = (Npc(MapNpc(GetPlayerMap(victim)).Npc(NpcNum).Num).stat(Stats.agility))
    Else
        accuracyrate = (NPCBase(MapNpc(GetPlayerMap(victim)).Npc(NpcNum).Level).Acc / 100) * (100 + Npc(MapNpc(GetPlayerMap(victim)).Npc(NpcNum).Num).stat(Stats.agility))
    End If
    
    accuracyrate = (accuracyrate / 100) * rand(1, 100)
    evasionrate = (GetPlayerEvasion(victim) / 100) * rand(1, 100)
    
    If accuracyrate >= evasionrate Then
        CanPlayerDodgeNPC = False
    Else
        CanPlayerDodgeNPC = True
    End If
End Function

Public Function CanNpcBlock(ByVal NpcNum As Long) As Boolean
Dim rate As Long
Dim stat As Long
Dim rndNum As Long

    CanNpcBlock = False
    
    stat = Npc(NpcNum).stat(Stats.agility) / 5  'guessed shield agility
    rate = stat / 12.08
    
    rndNum = rand(1, 100)
    
    If rndNum <= rate Then
        CanNpcBlock = True
    End If
    
End Function

Public Function CanNpcCrit(ByVal NpcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = Npc(NpcNum).stat(Stats.agility) / 52.08
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodgePlayer(ByVal NpcNum As Long, ByVal Attacker As Long) As Boolean
Dim accuracyrate As Long, evasionrate As Long

    CanNpcDodgePlayer = False

    'automatic
    If rand(1, 100) <= 25 Then
        Exit Function
    End If
    
    accuracyrate = (GetPlayerAccuracy(Attacker) / 100) * rand(1, 100)
    
    If Npc(MapNpc(GetPlayerMap(Attacker)).Npc(NpcNum).Num).ND = 0 Or Not UZ Then
        evasionrate = (Npc(MapNpc(GetPlayerMap(Attacker)).Npc(NpcNum).Num).stat(Stats.Willpower) / 100)
    Else
        evasionrate = (NPCBase(MapNpc(Player(Attacker).Map).Npc(NpcNum).Level).Esq / 100) * (100 + Npc(MapNpc(Player(Attacker).Map).Npc(NpcNum).Num).stat(Stats.Willpower))
    End If
    
    evasionrate = (evasionrate / 100) * rand(1, 100)
    
    If accuracyrate >= evasionrate Then
        CanNpcDodgePlayer = False
    Else
        CanNpcDodgePlayer = True
    End If
End Function

Public Function CanNpcParry(ByVal NpcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = Npc(NpcNum).stat(Stats.Strength) * 0.25
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal MapNPCNum As Long)
Dim blockAmount As Long
Dim NpcNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, MapNPCNum) Then
    
        MapNum = GetPlayerMap(Index)
        NpcNum = MapNpc(MapNum).Npc(MapNPCNum).Num
        
        If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_TREINO Or Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_TREINOHOUSE Then
            If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_TREINOHOUSE Then
                If Not IsHouseValid(Index) Then
                    PlayerMsg Index, printf("Sua casa expirou! renove ela entrando em nosso site!"), brightred
                    Exit Sub
                End If
            End If
            GivePlayerEXP Index, Npc(NpcNum).Exp
            Exit Sub
        End If
    
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = TARGET_TYPE_PLAYER
        MapNpc(MapNum).Npc(MapNPCNum).Target = Index
    
        ' check if NPC can avoid the attack
        If CanNpcDodgePlayer(MapNPCNum, Index) Then
            SendActionMsg MapNum, actionf("Esquivou!"), Pink, 1, (MapNpc(MapNum).Npc(MapNPCNum).X * 32), (MapNpc(MapNum).Npc(MapNPCNum).Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        'blockAmount = CanNpcBlock(MapNpcNum)
        'Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (Npc(NpcNum).stat(Stats.Endurance)))
        ' randomise from 1 to max hit
        Damage = (Damage / 100) * rand(75, 100)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, actionf("Critico!"), BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
        
        If Damage <= 0 Then Damage = 1
        
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, MapNPCNum, Damage)
        End If
    End If
End Sub
Public Sub TryNpcShootPlayer(MapNum As Long, MapNPCNum As Long, Index As Long)
Dim blockAmount As Long
Dim NpcNum As Long
Dim Damage As Long
Dim Buffer As clsBuffer

        Damage = 0
        NpcNum = MapNpc(MapNum).Npc(MapNPCNum).Num
        
        If MapNpc(MapNum).TempNpc(MapNPCNum).StunDuration > 0 Then Exit Sub
        
        Call CreateProjectile(MapNum, Index, MapNPCNum, TARGET_TYPE_PLAYER, Npc(NpcNum).ArrowAnim, 0, 1)
        
        ' check if NPC can avoid the attack
        If CanPlayerDodgeNPC(Index, MapNPCNum) Then
            SendActionMsg MapNum, actionf("Esquivou!"), Pink, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            
            ' Send this packet so they can see the npc attacking
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcAttack
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Index
            Buffer.WriteByte 1
            Buffer.WriteByte 1
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
            Exit Sub
        End If
        
        ' Get the damage we can do
        Damage = (GetNpcDamage(MapNum, MapNPCNum) / 100) * Npc(NpcNum).ArrowDamage
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(Index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (GetPlayerDefence(Index) * 2))
        ' randomise from 1 to max hit
        Damage = (Damage / 100) * rand(75, 100)
        
        If Damage <= 0 Then Damage = 1
        
        If Damage > 0 Then
            Call NpcAttackPlayer(MapNPCNum, Index, Damage, , True)
            Call SendAnimation(MapNum, Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).ArrowAnimation, GetPlayerX(Index), GetPlayerY(Index), MapNpc(MapNum).Npc(MapNPCNum).Dir, TARGET_TYPE_PLAYER, Index)
            
            ' Send this packet so they can see the npc attacking
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcAttack
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Index
            Buffer.WriteByte 1
            Buffer.WriteByte 1
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        End If
End Sub

Public Sub TryPlayerShootNpc(ByVal Index As Long, ByVal MapNPCNum As Long)
Dim blockAmount As Long
Dim NpcNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerShootNpc(Index, MapNPCNum) Then
        
        MapNum = GetPlayerMap(Index)
        NpcNum = MapNpc(MapNum).Npc(MapNPCNum).Num
        
        Call CreateProjectile(MapNum, Index, MapNPCNum, TARGET_TYPE_NPC, Item(GetPlayerEquipment(Index, Weapon)).Projectile, Item(GetPlayerEquipment(Index, Weapon)).Rotation)
        
        If Npc(NpcNum).Behaviour = NPC_BEHAVIOUR_TREINO Then
            GivePlayerEXP Index, Npc(NpcNum).Exp
            Exit Sub
        End If
        
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = TARGET_TYPE_PLAYER
        MapNpc(MapNum).Npc(MapNPCNum).Target = Index
        
        ' check if NPC can avoid the attack
        If CanNpcDodgePlayer(MapNPCNum, Index) Then
            SendActionMsg MapNum, actionf("Esquivou!"), Pink, 1, (MapNpc(MapNum).Npc(MapNPCNum).X * 32), (MapNpc(MapNum).Npc(MapNPCNum).Y * 32)
            Exit Sub
        End If
        
        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(MapNPCNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        'Damage = Damage - rand(1, (Npc(npcnum).stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        Damage = (Damage / 100) * rand(75, 100)
        
        If Damage <= 0 Then Damage = 1
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, actionf("Critico!"), BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
        
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, MapNPCNum, Damage)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal MapNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).Npc(MapNPCNum).Num <= 0 Then
        Exit Function
    End If
    
    Dim PlanetNum As Long
    If UZ Then
    PlanetNum = PlayerMapIndex(GetPlayerMap(Attacker))
    If PlanetNum > 0 Then
        If Trim$(LCase(PlayerPlanet(PlanetNum).PlanetData.Owner)) = Trim$(LCase(GetPlayerName(Attacker))) Then Exit Function
    End If
    End If

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNPCNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
    
        ' exit out early
        If IsSpell Then
             If NpcNum > 0 Then
                If Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).Speed
        Else
            attackspeed = 500
        End If
        
        attackspeed = attackspeed - (GetPlayerStat(Attacker, agility) * 5)
        
        If attackspeed < 250 Then attackspeed = 250

        If NpcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)

                Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT

                    NpcX = MapNpc(MapNum).Npc(MapNPCNum).X

                    NpcY = MapNpc(MapNum).Npc(MapNPCNum).Y + 1

                Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT

                    NpcX = MapNpc(MapNum).Npc(MapNPCNum).X

                    NpcY = MapNpc(MapNum).Npc(MapNPCNum).Y - 1

                Case DIR_LEFT

                    NpcX = MapNpc(MapNum).Npc(MapNPCNum).X + 1

                    NpcY = MapNpc(MapNum).Npc(MapNPCNum).Y

                Case DIR_RIGHT

                    NpcX = MapNpc(MapNum).Npc(MapNPCNum).X - 1

                    NpcY = MapNpc(MapNum).Npc(MapNPCNum).Y

            End Select

            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then
                    If Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        CanPlayerAttackNpc = True
                    Else
                        If Npc(NpcNum).Event > 0 Then
                            InitEvent Attacker, Npc(NpcNum).Event
                            Exit Function
                        End If
                        If Len(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                            Call SendChatBubble(MapNum, MapNPCNum, TARGET_TYPE_NPC, Trim$(Npc(NpcNum).AttackSay), DarkBrown)
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function CanPlayerShootNpc(ByVal Attacker As Long, ByVal MapNPCNum As Long) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).Npc(MapNPCNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNPCNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then

        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If isInRange(Item(GetPlayerEquipment(Attacker, Weapon)).Range, GetPlayerX(Attacker), GetPlayerY(Attacker), MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y) = False Then
                Exit Function
            End If
            attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).Speed
        Else
            attackspeed = 500
        End If
        
        attackspeed = attackspeed - (GetPlayerStat(Attacker, agility) * 5)
        
        If attackspeed < 250 Then attackspeed = 250

        If NpcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + attackspeed Then
            If Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then CanPlayerShootNpc = True
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal MapNPCNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim Def As Long
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNPCNum).Num
    Name = Trim$(Npc(NpcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    ' Check for a weapon and say damage
    SendActionMsg MapNum, "-" & Damage, brightred, 1, (MapNpc(MapNum).Npc(MapNPCNum).X * 32), (MapNpc(MapNum).Npc(MapNPCNum).Y * 32)
    SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y
    
    ' send animation
    If n > 0 Then
        If Not overTime Then
            If SpellNum = 0 Then
                Call SendEffect(MapNum, Item(n).Effect, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, , Attacker)
                Call SendAnimation(MapNum, Item(n).Animation, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, GetPlayerDir(Attacker), , , , , , , Attacker)
                SendMapSound Attacker, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, SoundEntity.seItem, n
            Else
                If Spell(SpellNum).CastPlayerAnim = 5 Then SendAttack Attacker
            End If
        End If
    Else
        If SpellNum = 0 Then Call SendAnimation(MapNum, PlayerAttackAnim, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, GetPlayerDir(Attacker), , , , , , , Attacker)
    End If
    'If SpellNum > 0 Then
    '    If Spell(SpellNum).SpellAnim > 0 Then
    '        Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, MapNpc(MapNum).Npc(MapNpcNum).X, MapNpc(MapNum).Npc(MapNpcNum).Y, GetPlayerDir(attacker))
    '        SendMapSound attacker, MapNpc(MapNum).Npc(MapNpcNum).X, MapNpc(MapNum).Npc(MapNpcNum).Y, SoundEntity.seSpell, SpellNum
    '    End If
    'End If
    
    Call SendFlash(MapNPCNum, MapNum, True)
    
    If Damage >= MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) Then
    
        If UZ Then
            If getProvação(MapNum) = 0 Then
                If TempPlayer(Attacker).MatchIndex > 0 Then
                    MatchData(TempPlayer(Attacker).MatchIndex).TotalNpcs = MatchData(TempPlayer(Attacker).MatchIndex).TotalNpcs - 1
                    If Planets(MatchData(TempPlayer(Attacker).MatchIndex).Planet).Type = 0 Then MatchData(TempPlayer(Attacker).MatchIndex).Points = MatchData(TempPlayer(Attacker).MatchIndex).Points + 1 'MapNpc(MapNum).Npc(MapNPCNum).Points
                    If Planets(MatchData(TempPlayer(Attacker).MatchIndex).Planet).Type = 4 Then
                        Dim TesouroQuant As Long
                        TesouroQuant = 1 + (Int(MatchData(TempPlayer(Attacker).MatchIndex).WaveNum) / 10)
                        If rand(0, 100) >= 50 Then Call SpawnItem(TesouroItem, TesouroQuant, MapNum, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y)
                        If rand(0, 100) < 10 Or IsElite(NpcNum) Then
                            Dim DropX As Long, DropY As Long, Drops As Long, rX As Long, rY As Long
                            Drops = rand(1, 5)
                            rX = MapNpc(MapNum).Npc(MapNPCNum).X
                            rY = MapNpc(MapNum).Npc(MapNPCNum).Y
                            For n = 1 To Drops
                                DropX = rand(rX - 1, rX + 1)
                                DropY = rand(rY - 1, rY + 1)
                                If DropX > 0 And DropX < Map(MapNum).MaxX Then
                                    If DropY > 0 And DropY < Map(MapNum).MaxY Then
                                        If DropX <> rX Or DropY <> rY Then
                                            If Map(MapNum).Tile(DropX, DropY).Type = TileType.TILE_TYPE_WALKABLE Then
                                                Call SpawnItem(TesouroItem, 1, MapNum, DropX, DropY)
                                            End If
                                        End If
                                    End If
                                End If
                            Next n
                        End If
                        If NpcCount(MapNum) <= 1 Then
                            MatchData(TempPlayer(Attacker).MatchIndex).WaveTick = 0
                        End If
                    End If
                    If IsDaily(Attacker, KillCreatures) Then UpdateDaily Attacker
                    Select Case Planets(MatchData(TempPlayer(Attacker).MatchIndex).Planet).Especie
                        Case 1: If IsDaily(Attacker, KillHumanCreatures) Then UpdateDaily Attacker
                        Case 2: If IsDaily(Attacker, KillInsectCreatures) Then UpdateDaily Attacker
                        Case 3: If IsDaily(Attacker, KillFeralCreatures) Then UpdateDaily Attacker
                    End Select
                    If IsElite(NpcNum) Then
                        MatchData(TempPlayer(Attacker).MatchIndex).Stars = MatchData(TempPlayer(Attacker).MatchIndex).Stars + 1
                        If IsDaily(Attacker, KillElite) Then UpdateDaily Attacker
                        Select Case Planets(MatchData(TempPlayer(Attacker).MatchIndex).Planet).Especie
                            Case 1: If IsDaily(Attacker, KillHumanElite) Then UpdateDaily Attacker
                            Case 2: If IsDaily(Attacker, KillInsectElite) Then UpdateDaily Attacker
                            Case 3: If IsDaily(Attacker, KillFeralElite) Then UpdateDaily Attacker
                        End Select
                    End If
                    SendMatchData TempPlayer(Attacker).MatchIndex
                End If
                
                
                Dim PlanetNum As Long
                If Not IsPlayerMap(MapNum) And MapNum >= PlanetStart And MapNum <= PlanetStart + MAX_PLANETS + 1 Then
                    PlanetNum = GetPlanetNum(MapNum)
                    If Planets(PlanetNum).Type = 1 Then
                        SendBossMsg MapNum, "Parabéns! Você concluiu a missão! Este planeta explodirá em 10 segundos!", Yellow
                        Npc(NpcNum).Name = vbNullString
                        GivePlayerVIPExp Attacker, Planets(PlanetNum).Level
                        If Player(Attacker).Guild > 0 Then GiveGuildExp Attacker, Planets(PlanetNum).Level, Planets(PlanetNum).Level
                        Planets(PlanetNum).Owner = GetPlayerName(Attacker)
                        Planets(PlanetNum).State = 2
                        Planets(PlanetNum).TimeToExplode = GetTickCount + 10000
                        If TempPlayer(Attacker).PlanetService = PlanetNum Then
                            CompleteService Attacker
                        End If
                    Else
                        HealPlayer Attacker, GetPlayerMaxVital(Attacker, HP) * 0.02
                        SendActionMsg GetPlayerMap(Attacker), "+2% HP", brightgreen, 1, GetPlayerX(Attacker) * 32, GetPlayerY(Attacker) * 32
                        HealPlayerMP Attacker, GetPlayerMaxVital(Attacker, MP) * 0.02
                        SendActionMsg GetPlayerMap(Attacker), "+2% KI", brightblue, 1, GetPlayerX(Attacker) * 32, GetPlayerY(Attacker) * 32
                    End If
                    If Planets(PlanetNum).Type = 5 Then
                        If NpcCount(MapNum) <= 1 Then
                            GiveInvItem Attacker, MoedaZ, Int(Planets(PlanetNum).Preco)
                            SendBossMsg GetPlayerMap(Attacker), "Parabéns! Você protegeu este planeta de ser saqueado por piratas!", Yellow, Attacker
                            PlayerMsg Attacker, "-- Recompensa pela missão --", brightgreen
                            PlayerMsg Attacker, "Você recebeu: " & Int(Planets(PlanetNum).Preco) & "z", brightgreen
                            GivePlayerEXP Attacker, Int(ExperienceBase(Planets(PlanetNum).Level))
                            PlayerMsg Attacker, "Você recebeu: " & Int(ExperienceBase(Planets(PlanetNum).Level)) & " exp", brightgreen
                            GivePlayerVIPExp Attacker, Planets(PlanetNum).Level
                            If Player(Attacker).Guild > 0 Then GiveGuildExp Attacker, Planets(PlanetNum).Level, Planets(PlanetNum).Level
                            Planets(PlanetNum).Owner = GetPlayerName(Attacker)
                            Planets(PlanetNum).State = 2
                            Planets(PlanetNum).TimeToExplode = GetTickCount + 10000
                            If TempPlayer(Attacker).PlanetService = PlanetNum Then
                                CompleteService Attacker
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        ' Calculate exp to give attacker
        If Not UZ Then
            Exp = Npc(NpcNum).Exp
        Else
            If Npc(NpcNum).ND = 0 Then
                Exp = Npc(NpcNum).Exp
            Else
                Exp = Experience(MapNpc(MapNum).Npc(MapNPCNum).Level) / 3000
            End If
        End If

        ' Make sure we dont get less then 0
        If Exp <= 0 Then
            Exp = 1
        End If
        
        If GetPlayerLevel(Attacker) > Npc(NpcNum).ND Then
            If GetPlayerLevel(Attacker) - Npc(NpcNum).ND > 10 Then
                Exp = Exp * 1.1
            Else
                Exp = (Exp / 100) * ((GetPlayerLevel(Attacker) - Npc(NpcNum).ND) + 100)
            End If
        End If

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, GetPlayerMap(Attacker)
            If Npc(NpcNum).Event > 0 Then
                Party_ShareEvent TempPlayer(Attacker).inParty, Npc(NpcNum).Event, GetPlayerMap(Attacker)
            End If
        Else
            ' no party - keep exp for self
            GivePlayerEXP Attacker, Exp
            'If Exp > 0 Then PlayerMsg Attacker, "Você recebeu " & Exp & " exp", Yellow
            If Npc(NpcNum).Event > 0 Then
                InitEvent Attacker, Npc(NpcNum).Event
            End If
        End If
        
        'Drop the goods if they get it
        For i = 1 To 10
            If Npc(NpcNum).Drop(i).Num > 0 Then
                n = Int(Rnd * Npc(NpcNum).Drop(i).Chance) + 1
        
                If n = 1 Then
                    Call SpawnItem(Npc(NpcNum).Drop(i).Num, Int(Npc(NpcNum).Drop(i).Value * Options.DropFactor), MapNum, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y)
                End If
            End If
        Next i

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(MapNPCNum).Num = 0
        MapNpc(MapNum).TempNpc(MapNPCNum).SpawnWait = GetTickCount
        MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) = 0
        UpdateMapBlock MapNum, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, False
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).TempNpc(MapNPCNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).TempNpc(MapNPCNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong MapNPCNum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
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
        
        If DeathEffect > 0 Then Call SendEffect(MapNum, DeathEffect, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y)
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) = MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) - Damage

        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = 1 ' player
        MapNpc(MapNum).Npc(MapNPCNum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(MapNPCNum).Num Then
                    MapNpc(MapNum).Npc(i).Target = Attacker
                    MapNpc(MapNum).Npc(i).TargetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(MapNum).TempNpc(MapNPCNum).stopRegen = True
        MapNpc(MapNum).TempNpc(MapNPCNum).stopRegenTimer = GetTickCount
        
        'If TempPlayer(Attacker).Target = 0 Then
            TempPlayer(Attacker).Target = MapNPCNum
            TempPlayer(Attacker).TargetType = TARGET_TYPE_NPC
            SendTarget Attacker
        'End If
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunNPC MapNPCNum, MapNum, SpellNum
        End If
        
        SendMapNpcVitals MapNum, MapNPCNum
    End If

    If SpellNum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = GetTickCount
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal MapNPCNum As Long, ByVal Index As Long)
Dim MapNum As Long, NpcNum As Long, blockAmount As Long, Damage As Long, Buffer As clsBuffer

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(MapNPCNum, Index) Then
        MapNum = GetPlayerMap(Index)
        NpcNum = MapNpc(MapNum).Npc(MapNPCNum).Num

    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodgeNPC(Index, MapNPCNum) Then
            SendActionMsg MapNum, actionf("Esquivou!"), Pink, 1, (Player(Index).X * 32), (Player(Index).Y * 32)
            
            ' Send this packet so they can see the npc attacking
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcAttack
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Index
            Buffer.WriteByte 2
            Buffer.WriteByte 0
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(MapNum, MapNPCNum)
        
        ' if the player blocks, take away the block amount
        'blockAmount = CanPlayerBlock(Index)
        'Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - GetPlayerDefence(Index)
        
        ' randomise for up to 10% lower than max hit
        Damage = (Damage / 100) * rand(75, 100)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, actionf("Critico!"), BrightCyan, 1, (MapNpc(MapNum).Npc(MapNPCNum).X * 32), (MapNpc(MapNum).Npc(MapNPCNum).Y * 32)
        End If
        
        If Damage <= 0 Then Damage = 1

        If Damage > 0 Then
        
            ' Send this packet so they can see the npc attacking
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcAttack
            Buffer.WriteLong MapNPCNum
            Buffer.WriteLong Index
            Buffer.WriteByte 1
            Buffer.WriteByte 0
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
            
            Call NpcAttackPlayer(MapNPCNum, Index, Damage)
            Exit Sub
        End If
        
        ' Send this packet so they can see the npc attacking
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcAttack
        Buffer.WriteLong MapNPCNum
        Buffer.WriteLong Index
        Buffer.WriteByte 0
        Buffer.WriteByte 0
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
    End If
End Sub

Function CanNpcAttackPlayer(ByVal MapNPCNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).Npc(MapNPCNum).Num <= 0 Then
        Exit Function
    End If
    
    If MapNpc(GetPlayerMap(Index)).TempNpc(MapNPCNum).StunTimer + (MapNpc(GetPlayerMap(Index)).TempNpc(MapNPCNum).StunDuration * 1000) > GetTickCount Then Exit Function
    
    If Player(Index).IsDead = 1 Then Exit Function

    MapNum = GetPlayerMap(Index)
    NpcNum = MapNpc(MapNum).Npc(MapNPCNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    attackspeed = Npc(NpcNum).attackspeed
    
    If attackspeed < 100 Then attackspeed = 500

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum).TempNpc(MapNPCNum).AttackTimer + attackspeed Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(MapNum).TempNpc(MapNPCNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NpcNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum).Npc(MapNPCNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(MapNPCNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(MapNum).Npc(MapNPCNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(MapNPCNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum).Npc(MapNPCNum).Y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum).Npc(MapNPCNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum).Npc(MapNPCNum).Y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum).Npc(MapNPCNum).X) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
            
            If CanNpcAttackPlayer = False Then
                If Npc(NpcNum).Ranged = 1 Then
                    If isInRange(Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Range, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, GetPlayerX(Index), GetPlayerY(Index)) Then
                        TryNpcShootPlayer MapNum, MapNPCNum, Index
                    End If
                End If
                If Npc(NpcNum).IA(NPCIA.Shunppo).Data(1) = 1 Then
                    If isInRange(Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Range, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, GetPlayerX(Index), GetPlayerY(Index)) Then
                        If rand(1, 100) <= Npc(NpcNum).IA(NPCIA.Shunppo).Data(2) Then
                            NpcShunppo Index, MapNPCNum
                            If TempPlayer(Index).Target = 0 Then
                                TempPlayer(Index).Target = MapNPCNum
                                TempPlayer(Index).TargetType = TARGET_TYPE_NPC
                                SendTarget Index
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNPCNum As Long, ByVal victim As Long, ByVal Damage As Long, Optional DoAnimation As Boolean = True, Optional IsRangeAttack As Boolean = False)
    Dim Name As String
    Dim Exp As Long
    Dim MapNum As Long
    Dim i As Long

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).Npc(MapNPCNum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(victim)
    Name = Trim$(Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Name)
    
    'Damage = Damage - GetPlayerDefence(victim)
    
    If Damage <= 0 Then Damage = 1
    
    ' set the regen timer
    MapNpc(MapNum).TempNpc(MapNPCNum).stopRegen = True
    MapNpc(MapNum).TempNpc(MapNPCNum).stopRegenTimer = GetTickCount
    
    ' Say damage
    SendActionMsg GetPlayerMap(victim), "-" & Damage, brightred, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
    
    If MapNpc(MapNum).Npc(MapNPCNum).Num > 0 And GetPlayerMap(victim) > 0 Then
        If DoAnimation = True Then
            Call SendAnimation(MapNum, Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Animation, GetPlayerX(victim), GetPlayerY(victim), MapNpc(MapNum).Npc(MapNPCNum).Dir, TARGET_TYPE_PLAYER, victim)
            If MapNpc(MapNum).Npc(MapNPCNum).Num > 0 Then Call SendEffect(MapNum, Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Effect, GetPlayerX(victim), GetPlayerY(victim))
        End If
    End If
    ' send the sound
    SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNPCNum).Num
    
    Call SendFlash(victim, MapNum, False)
    
    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' kill player
        KillPlayer victim
        Call CheckProvacoesMap(victim, False)
        
        ' Player is dead
        Call PlayerMsg(victim, printf("Você foi morto por %s", Name), brightred)
        
        If DeathEffect > 0 Then Call SendEffect(MapNum, DeathEffect, GetPlayerX(victim), GetPlayerY(victim))

        ' Set NPC target to 0
        MapNpc(MapNum).Npc(MapNPCNum).Target = 0
        MapNpc(MapNum).Npc(MapNPCNum).TargetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        If IsRangeAttack = False Then
            If MapNpc(MapNum).Npc(MapNPCNum).Num > 0 Then
            'Inteligencia artificial
            If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Spawn).Data(1) = 1 And MapNpc(MapNum).Npc(MapNPCNum).Spawned = 0 Then
                If rand(1, 100) <= Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Spawn).Data(2) Then
                    Call InitEvent(victim, Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Spawn).Data(3))
                    Call SendAnimation(MapNum, Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Spawn).Data(4), MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, MapNpc(MapNum).Npc(MapNPCNum).Dir, TARGET_TYPE_NPC, MapNPCNum)
                    MapNpc(MapNum).Npc(MapNPCNum).Spawned = 1
                End If
            End If
            
            If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Stun).Data(1) = 1 Then
                If rand(1, 100) <= Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Stun).Data(2) Then
                    If TempPlayer(victim).StunDuration = 0 Then 'Não stunar mais de uma vez
                        Call NpcStunPlayer(victim, Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Stun).Data(3))
                        Call NpcMakeImpact(victim, Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Stun).Data(4), MapNum, MapNPCNum)
                        Call SendAnimation(MapNum, Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Stun).Data(5), Player(victim).X, Player(victim).Y, MapNpc(MapNum).Npc(MapNPCNum).Dir, TARGET_TYPE_PLAYER, victim)
                    End If
                End If
            End If
            
            If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Storm).Data(1) = 1 Then
                If rand(1, 100) <= Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Storm).Data(2) Then
                    If MapNpc(MapNum).TempNpc(MapNPCNum).LastStorm + (Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Storm).Data(4) * 1000) < GetTickCount Then
                        SendActionMsg MapNum, Trim$(Spell(Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Storm).Data(3)).Name) & "!", brightred, ActionMsgType.ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(MapNPCNum).X * 32, MapNpc(MapNum).Npc(MapNPCNum).Y * 32
                        RegisterNPCAOE MapNPCNum, MapNum, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Storm).Data(3)
                        MapNpc(MapNum).TempNpc(MapNPCNum).LastStorm = GetTickCount
                    End If
                End If
            End If
            
            If Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Shunppo).Data(1) = 1 Then
                If isInRange(Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).Range, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, GetPlayerX(victim), GetPlayerY(victim)) Then
                    If rand(1, 100) <= Npc(MapNpc(MapNum).Npc(MapNPCNum).Num).IA(NPCIA.Shunppo).Data(2) Then
                        NpcShunppo victim, MapNPCNum
                    End If
                End If
            End If
            End If
        End If
        
        If TempPlayer(victim).Target = 0 Then
            TempPlayer(victim).Target = MapNPCNum
            TempPlayer(victim).TargetType = TARGET_TYPE_NPC
            SendTarget victim
        End If
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
    End If

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim NpcNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, victim) Then
    
        MapNum = GetPlayerMap(Attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodgePlayer(victim, Attacker) Then
            SendActionMsg MapNum, actionf("Esquivou!"), Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If Item(GetPlayerEquipment(Attacker, Weapon)).Projectile > 0 Then
                If isInRange(Item(GetPlayerEquipment(Attacker, Weapon)).Range, GetPlayerX(Attacker), GetPlayerY(Attacker), GetPlayerX(victim), GetPlayerY(victim)) Then
                    Call CreateProjectile(MapNum, Attacker, victim, TARGET_TYPE_PLAYER, Item(GetPlayerEquipment(Attacker, Weapon)).Projectile, Item(GetPlayerEquipment(Attacker, Weapon)).Rotation)
                Else
                    Call PlayerMsg(Attacker, printf("Fora do alcance."), brightred)
                End If
            End If
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - GetPlayerDefence(victim)
        
        ' randomise for up to 10% lower than max hit
        Damage = (Damage / 100) * rand(75, 100)
        
        If Damage <= 0 Then Damage = 1
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, actionf("Critico!"), BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, victim, Damage)
        End If
    End If
End Sub
Public Sub TryPlayerShootPlayer(ByVal Attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim NpcNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, victim, True) Then
    
        MapNum = GetPlayerMap(Attacker)
        
        Call CreateProjectile(MapNum, Attacker, victim, TARGET_TYPE_PLAYER, Item(GetPlayerEquipment(Attacker, Weapon)).Projectile, Item(GetPlayerEquipment(Attacker, Weapon)).Rotation)
        
        ' check if NPC can avoid the attack
        If CanPlayerDodgePlayer(victim, Attacker) Then
            SendActionMsg MapNum, actionf("Esquivou!"), Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - GetPlayerDefence(victim)
        
        ' randomise for up to 10% lower than max hit
        Damage = (Damage / 100) * rand(75, 100)
        
        If Damage <= 0 Then Damage = 1
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, actionf("Critico!"), BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, victim, Damage)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim attackspeed As Long
    
    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            attackspeed = (Item(GetPlayerEquipment(Attacker, Weapon)).Speed - (GetPlayerStat(Attacker, agility) * 10))
        Else
            attackspeed = (1000 - (GetPlayerStat(Attacker, agility) * 10))
        End If
        
        If attackspeed < 200 Then attackspeed = 200
        
        If GetTickCount < TempPlayer(Attacker).AttackTimer + attackspeed Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).LastMove + AFKTime < GetTickCount Then Exit Function
    
    ' Jogador vivo?
    If Player(victim).IsDead = 1 Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)

            Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT

   

                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) = GetPlayerX(Attacker))) Then Exit Function

            Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT

   

                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(victim) = GetPlayerX(Attacker))) Then Exit Function

            Case DIR_LEFT

   

                If Not ((GetPlayerY(victim) = GetPlayerY(Attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(Attacker))) Then Exit Function

            Case DIR_RIGHT

   

                If Not ((GetPlayerY(victim) = GetPlayerY(Attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(Attacker))) Then Exit Function

            Case Else

                Exit Function

        End Select
    End If

    ' Check if map is attackable
    If Not GetPlayerMap(Attacker) = ARENA_MAP Then
        If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE And Not Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Type = TILE_TYPE_ARENA Then
            If GetPlayerPK(victim) = NO Then
                Call PlayerMsg(Attacker, printf("Essa é uma zona segura!"), brightred)
                Exit Function
            End If
        End If
        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_OWNER Then
            If TempPlayer(Attacker).Instance <> TempPlayer(victim).Instance Then
                Exit Function
            End If
        End If
    End If
    
    

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure the victim isn't an admin
    If Not GetPlayerMap(Attacker) = ARENA_MAP Then
    'If GetPlayerAccess(victim) > ADMIN_MONITOR And Not Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Type = TILE_TYPE_ARENA Then
    '    Call PlayerMsg(Attacker, printf("Você não pode atacar %s!", GetPlayerName(victim)), brightred)
    '    Exit Function
    'End If
    End If

    ' Make sure attacker is high enough level
    'If GetPlayerLevel(Attacker) < SECURELEVEL Then
    '    Call PlayerMsg(Attacker, printf("Você está abaixo do level %d, você não pode atacar ninguém!", Val(SECURELEVEL)), brightred)
    '    Exit Function
    'End If

    ' Make sure victim is high enough level
    'If GetPlayerLevel(victim) < SECURELEVEL Then
    '    Call PlayerMsg(Attacker, printf("%s está abaixo do level %d e não pode ser atacado!", GetPlayerName(victim) & "," & SECURELEVEL), brightred)
    '    Exit Function
    'End If
    
    If Not GetPlayerMap(Attacker) = ARENA_MAP Then
    If TempPlayer(Attacker).inParty > 0 Then
        If TempPlayer(Attacker).inParty = TempPlayer(victim).inParty Then Exit Function
    End If
    End If

    CanPlayerAttackPlayer = True
End Function

Function CanPlayerShootPlayer(ByVal Attacker As Long, ByVal victim As Long) As Boolean
    
    Dim attackspeed As Long
    
    ' Check attack timer
    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).Speed - (GetPlayerStat(Attacker, agility) * 10)
    Else
        attackspeed = 500 - (GetPlayerStat(Attacker, agility) * 10)
    End If
    
    If attackspeed < 200 Then attackspeed = 200
    
    If GetTickCount < TempPlayer(Attacker).AttackTimer + attackspeed Then Exit Function
    
    If isInRange(Item(GetPlayerEquipment(Attacker, Weapon)).Range, GetPlayerX(Attacker), GetPlayerY(Attacker), GetPlayerX(victim), GetPlayerY(victim)) = False Then
        Exit Function
    End If
    

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).LastMove + AFKTime < GetTickCount Then Exit Function
    
    ' Jogador vivo?
    If Player(victim).IsDead = 1 Then Exit Function

    ' Check if map is attackable
    If Not GetPlayerMap(Attacker) = ARENA_MAP Then
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE And Not Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Type = TILE_TYPE_ARENA Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(Attacker, printf("Essa é uma zona segura!"), brightred)
            Exit Function
        End If
    End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure the victim isn't an admin
    If Not GetPlayerMap(Attacker) = ARENA_MAP Then
    If GetPlayerAccess(victim) > ADMIN_MONITOR And Not Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(victim), GetPlayerY(victim)).Type = TILE_TYPE_ARENA Then
        Call PlayerMsg(Attacker, printf("Você não pode atacar %s!", GetPlayerName(victim)), brightred)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < SECURELEVEL Then
        Call PlayerMsg(Attacker, printf("Você está abaixo do level %d e não pode atacar ninguém!", Val(SECURELEVEL)), brightred)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < SECURELEVEL Then
        Call PlayerMsg(Attacker, printf("%s está abaixo do level %d e não pode ser atacado!", GetPlayerName(victim) & "," & SECURELEVEL), brightred)
        Exit Function
    End If
    
    If TempPlayer(Attacker).inParty > 0 Then
        If TempPlayer(Attacker).inParty = TempPlayer(victim).inParty Then Exit Function
    End If
    End If

    CanPlayerShootPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount
    
    If TempPlayer(Attacker).Target = 0 Then
        TempPlayer(Attacker).Target = victim
        TempPlayer(Attacker).TargetType = TARGET_TYPE_PLAYER
        SendTarget Attacker
    End If
    
    'Damage = Damage - GetPlayerDefence(victim)
    If Damage <= 0 Then Damage = 1
    
    SendActionMsg GetPlayerMap(victim), "-" & Damage, brightred, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
    
    ' send animation
    If n > 0 Then
        If SpellNum = 0 Then
            Call SendEffect(GetPlayerMap(victim), Item(n).Effect, GetPlayerX(victim), GetPlayerY(victim), , Attacker)
            Call SendAnimation(GetPlayerMap(victim), Item(n).Animation, GetPlayerX(victim), GetPlayerY(victim), GetPlayerDir(Attacker), , , , , , , Attacker)
            SendMapSound Attacker, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seItem, n
        Else
            If Spell(SpellNum).CastPlayerAnim = 5 Then SendAttack Attacker
        End If
    Else
        Call SendAnimation(GetPlayerMap(victim), PlayerAttackAnim, GetPlayerX(victim), GetPlayerY(victim), GetPlayerDir(Attacker), , , , , , , Attacker)
    End If
    
    Call SendFlash(victim, GetPlayerMap(victim), False)
    
    If Damage >= GetPlayerVital(victim, Vitals.HP) And Not GetPlayerMap(Attacker) = SalaDoTempo Then
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " foi morto por " & GetPlayerName(Attacker), brightred)
        ' Calculate exp to give attacker
        Exp = 0 '(GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If

        If Exp > 0 Then
            Call SetPlayerExp(victim, GetPlayerExp(victim) - Exp)
            SendEXP victim
            Call PlayerMsg(victim, printf("Você perdeu %d exp.", Val(Exp)), brightred)
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, Exp, Attacker, GetPlayerMap(Attacker)
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, Exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        'If GetPlayerPK(victim) = NO Then
        '    If GetPlayerPK(Attacker) = NO Then
        '        Call SetPlayerPK(Attacker, Yes)
        '        Call SendPlayerData(Attacker)
        '        Call GlobalMsg(GetPlayerName(Attacker) & " se tornou um !!!", BrightRed)
        '    End If

        'Else
        '    Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        'End If
        
        Call CheckItemDrop(victim)
        If Player(victim).IsDead = 0 And TempPlayer(victim).inDevSuite = 0 Then Player(victim).IsDead = 1
        Call SendPlayerData(victim)
        'Call OnDeath(victim)
        If DeathEffect > 0 Then Call SendEffect(GetPlayerMap(victim), DeathEffect, GetPlayerX(victim), GetPlayerY(victim))
    Else
        ' Player not dead, just do the damage
        If GetPlayerMap(Attacker) <> SalaDoTempo Then
            Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
            Call SendVital(victim, Vitals.HP)
            Call SendVital(victim, Vitals.HP, Attacker)
        Else
            TempPlayer(Attacker).DamageAmount = TempPlayer(Attacker).DamageAmount + Damage
            If TempPlayer(Attacker).DamageAmount > 10000 Then
                TempPlayer(Attacker).DamageAmount = 0
                GiveInvItem Attacker, 205, 1
                PlayerMsg Attacker, "Parabéns! Sua Furia Sayajin aumentou!", Yellow
            End If
        End If
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
        
        'If TempPlayer(victim).Target = 0 Then
            TempPlayer(victim).Target = Attacker
            TempPlayer(victim).TargetType = TARGET_TYPE_PLAYER
            SendTarget victim
        'End If
        
        'if a stunning spell, stun the player
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunPlayer victim, SpellNum
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal Index As Long, ByVal spellslot As Long)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim TargetType As Byte
    Dim Target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    If (GetPlayerMap(Index) = ViagemMap And UZ) Or Player(Index).GravityHours > 0 Then Exit Sub
    
    SpellNum = GetPlayerSpell(Index, spellslot)
    MapNum = GetPlayerMap(Index)
    
    If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, printf("Tecnica em espera!"), brightred
        Exit Sub
    End If

    MPCost = Spell(SpellNum).MPCost
    
    If UZ Then
        If Spell(SpellNum).Type = SPELL_TYPE_TRANS And Spell(SpellNum).HairChange = 5 Then
            Dim PlanetNum As Long
            PlanetNum = GetPlanetNum(GetPlayerMap(Index))
            If PlanetNum > 0 And PlanetNum <= MAX_PLANETS Then
                If Planets(PlanetNum).MoonData.Pic > 0 Then MPCost = 0
            End If
        End If
    End If

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, printf("Sem KI necessário!"), brightred)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, printf("Você precisa ser nivel %d para usar esta tecnica.", Val(LevelReq)), brightred)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", brightred)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, printf("Apenas %s podem usar essa tecnica.", CheckGrammar(Trim$(Class(ClassReq).Name))), brightred)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    If Spell(SpellNum).Type = SpellType.SPELL_TYPE_VOAR And Map(GetPlayerMap(Index)).Fly = 1 Then
        PlayerMsg Index, printf("Você não pode voar aqui!"), brightred
        Exit Sub
    End If
    
    If Spell(SpellNum).Type = SpellType.SPELL_TYPE_TRANS Then
        If TempPlayer(Index).Fly = 1 Then
            If Spell(SpellNum).HairChange = 5 Then
                PlayerMsg Index, "Você não pode se transformar em Oozaru enquanto voa!", brightred
                Exit Sub
            End If
        End If
        If TempPlayer(Index).Trans = SpellNum Then
            TransPlayer Index, 0
            PlayerMsg Index, printf("Você se destransformou!"), White
            Exit Sub
        End If
    End If
    
    If TempPlayer(Index).HairChange = 5 Then
        PlayerMsg Index, "Você não pode usar habilidades enquanto está em forma de Oozaru!", brightred
        Exit Sub
    End If
    
    If Spell(SpellNum).Type = SpellType.SPELL_TYPE_LINEAR Or Spell(SpellNum).Type = SpellType.SPELL_TYPE_TRANS Or Spell(SpellNum).Type = SpellType.SPELL_TYPE_VOAR Then SpellCastType = 0
    
    TargetType = TempPlayer(Index).TargetType
    Target = TempPlayer(Index).Target
    Range = Spell(SpellNum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                PlayerMsg Index, printf("Você não tem um alvo."), brightred
            End If
            If TargetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg Index, printf("Fora do alcance."), brightred
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
                If Spell(SpellNum).Type = SPELL_TYPE_SHUNPPO And Target = Index Then HasBuffered = False
            ElseIf TargetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(MapNum).Npc(Target).X, MapNpc(MapNum).Npc(Target).Y) Then
                    PlayerMsg Index, printf("Fora do alcance."), brightred
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(SpellNum).CastAnim, GetPlayerX(Index), GetPlayerY(Index), GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index, , , , 1
        'SendActionMsg MapNum, "Casting " & Trim$(Spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        TempPlayer(Index).spellBuffer.Spell = spellslot
        TempPlayer(Index).spellBuffer.Timer = GetTickCount
        TempPlayer(Index).spellBuffer.Target = TempPlayer(Index).Target
        TempPlayer(Index).spellBuffer.tType = TempPlayer(Index).TargetType
        SendSpellBuffer Index, SpellNum
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal TargetType As Byte)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim X As Long, Y As Long, n As Long
   
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    SpellNum = GetPlayerSpell(Index, spellslot)
    MapNum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, printf("Sem KI necessário!"), brightred)
        Exit Sub
        Else
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPCost)
        Call SendVital(Index, Vitals.MP)
    End If
   
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, printf("Você precisa ser nivel %d para usar esta tecnica.", Val(LevelReq)), brightred)
        Exit Sub
    End If
   
    AccessReq = Spell(SpellNum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", brightred)
        Exit Sub
    End If
   
    ClassReq = Spell(SpellNum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, printf("Apenas %s podem usar essa tecnica.", CheckGrammar(Trim$(Class(ClassReq).Name))), brightred)
            Exit Sub
        End If
    End If
   
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    If Spell(SpellNum).Type = SPELL_TYPE_LINEAR Or Spell(SpellNum).Type = SPELL_TYPE_TRANS Or Spell(SpellNum).Type = SPELL_TYPE_VOAR Then SpellCastType = 0
   
    ' set the vital
    Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_HEALHP Or SPELL_TYPE_HEALMP
            Vital = GetPlayerStat(Index, Intelligence) + Spell(SpellNum).Vital
        Case SPELL_TYPE_DAMAGEHP Or SPELL_TYPE_LINEAR
            Vital = GetPlayerSkillDamage(Index, Spell(SpellNum).Vital)
        Case Else
            Vital = GetPlayerSkillDamage(Index, Spell(SpellNum).Vital)
    End Select
    
    If Player(Index).IsGod > 0 Then
        Vital = (Vital / 100) * (100 + (Player(Index).IsGod * 20))
    End If
    
    AoE = Spell(SpellNum).AoE
    Range = Spell(SpellNum).Range
   
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, Index, Vital, SpellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, Index, Vital, SpellNum
                    DidCast = True
                Case SPELL_TYPE_LINEAR
                
                    X = GetPlayerX(Index)
                    Y = GetPlayerY(Index)
                    
                    SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                    TempPlayer(Index).SpellCD(spellslot) = GetTickCount + (Spell(SpellNum).CDTime * 1000)
                    Call SendCooldown(Index, spellslot)
                    
                    TempPlayer(Index).spellBuffer.Spell = 0
                    TempPlayer(Index).spellBuffer.Target = 0
                    TempPlayer(Index).spellBuffer.Timer = 0
                    TempPlayer(Index).spellBuffer.tType = 0
                    SendSpellBuffer Index, 0
        
                    DidCast = True
                    
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP: X = X - Spell(SpellNum).LinearRange - 1
                        Case DIR_DOWN: X = X - Spell(SpellNum).LinearRange - 1
                        Case DIR_LEFT: Y = Y - Spell(SpellNum).LinearRange - 1
                        Case DIR_RIGHT: Y = Y - Spell(SpellNum).LinearRange - 1
                    End Select
                    
                    Dim j As Long
                    For j = 1 To 1 + (Spell(SpellNum).LinearRange * 2)
                        Select Case GetPlayerDir(Index)
                            Case DIR_UP
                                X = X + 1
                                Y = GetPlayerY(Index)
                            Case DIR_DOWN
                                X = X + 1
                                Y = GetPlayerY(Index)
                            Case DIR_LEFT
                                Y = Y + 1
                                X = GetPlayerX(Index)
                            Case DIR_RIGHT
                                Y = Y + 1
                                X = GetPlayerX(Index)
                        End Select
                        For i = 1 To Range
                            
                            Select Case GetPlayerDir(Index)
                                Case DIR_UP: Y = Y - 1
                                Case DIR_DOWN: Y = Y + 1
                                Case DIR_LEFT: X = X - 1
                                Case DIR_RIGHT: X = X + 1
                            End Select
                            
                            If Y < 0 Or Y > Map(MapNum).MaxY Or X < 0 Or X > Map(MapNum).MaxX Then Exit Sub
                            CheckAttack Index, X, Y, Vital, SpellNum
                            
                            If i = 1 Then n = 1
                            If i = Range Then n = 3
                            If n = 0 Then n = 2
                            
                            If GetPlayerDir(Index) = DIR_UP Or GetPlayerDir(Index) = DIR_DOWN Then
                                If X = GetPlayerX(Index) Then
                                    SendAnimation MapNum, Spell(SpellNum).SpellLinearAnim(n), X, Y, GetPlayerDir(Index), , , , 1
                                End If
                            End If
                            
                            If GetPlayerDir(Index) = DIR_LEFT Or GetPlayerDir(Index) = DIR_RIGHT Then
                                If Y = GetPlayerY(Index) Then
                                    SendAnimation MapNum, Spell(SpellNum).SpellLinearAnim(n), X, Y, GetPlayerDir(Index), , , , 1
                                End If
                            End If
                            
                            n = 0
                            
                            SendEffect MapNum, Spell(SpellNum).Effect, X, Y
                        Next i
                    Next j
                Case SPELL_TYPE_TRANS
                    If TempPlayer(Index).Trans = SpellNum Then
                        TransPlayer Index, 0
                    Else
                        TransPlayer Index, SpellNum
                    End If
                    If Spell(SpellNum).SpellAnim > 0 And TempPlayer(Index).Trans > 0 Then SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, GetPlayerX(Index), GetPlayerY(Index), GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index
                    SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                    DidCast = True
                Case SPELL_TYPE_VOAR
                    If TempPlayer(Index).Fly = 1 Then
                        If Not IsTileValid(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type) Then
                            PlayerMsg Index, printf("Você não pode pousar aqui!"), brightred
                        Else
                            TempPlayer(Index).Fly = 0
                            SendPlayerFly Index
                            If Spell(SpellNum).Sound <> "" Then SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                        End If
                    Else
                        TempPlayer(Index).Fly = 1
                        SendPlayerFly Index
                        If Spell(SpellNum).Sound <> "" Then SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                    End If
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                X = GetPlayerX(Index)
                Y = GetPlayerY(Index)
            ElseIf SpellCastType = 3 Then
                If TargetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
               
                If TargetType = TARGET_TYPE_PLAYER Then
                    X = GetPlayerX(Target)
                    Y = GetPlayerY(Target)
                Else
                    X = MapNpc(MapNum).Npc(Target).X
                    Y = MapNpc(MapNum).Npc(Target).Y
                End If
               
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                    PlayerMsg Index, printf("Fora do alcance."), brightred
                    SendClearSpellBuffer Index
                End If
                
                If TargetType = TARGET_TYPE_PLAYER Then
                    If Spell(SpellNum).Projectile > 0 Then SendAnimation MapNum, Spell(SpellNum).SpellAnim, GetPlayerX(Target), GetPlayerY(Target), GetPlayerDir(Index), TARGET_TYPE_PLAYER, Target
                Else
                    If Spell(SpellNum).Projectile > 0 Then SendAnimation MapNum, Spell(SpellNum).SpellAnim, MapNpc(MapNum).Npc(Target).X, MapNpc(MapNum).Npc(Target).Y, GetPlayerDir(Index), TARGET_TYPE_NPC, Target
                End If
            End If
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    If Spell(SpellNum).AoEDuration = 0 Then
                        CreateProjectile GetPlayerMap(Index), Index, Target, TargetType, Spell(SpellNum).Projectile, Spell(SpellNum).RotateSpeed
                        Dim AnimX As Long, AnimY As Long
                        If Spell(SpellNum).Projectile = 0 Then
                            For AnimX = GetPlayerX(Index) - AoE To GetPlayerX(Index) + AoE
                                For AnimY = GetPlayerY(Index) - AoE To GetPlayerY(Index) + AoE
                                    If AnimX >= 0 And AnimX <= Map(GetPlayerMap(Index)).MaxX Then
                                        If AnimY >= 0 And AnimY <= Map(GetPlayerMap(Index)).MaxY Then
                                            SendAnimation MapNum, Spell(SpellNum).SpellAnim, AnimX, AnimY, GetPlayerDir(Index)
                                            'Resources
                                            CheckResource Index, AnimX, AnimY, Vital
                                        End If
                                    End If
                                Next AnimY
                            Next AnimX
                        End If
                        For i = 1 To Player_HighIndex
                            If IsPlaying(i) Then
                                If i <> Index Then
                                    If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                        If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                            If CanPlayerAttackPlayer(Index, i, True) Then
                                                'If Spell(SpellNum).Projectile = 0 Then SendAnimation MapNum, Spell(SpellNum).SpellAnim, GetPlayerX(i), GetPlayerY(i), GetPlayerDir(Index), TARGET_TYPE_PLAYER, i
                                                SendEffect MapNum, Spell(SpellNum).Effect, GetPlayerX(i), GetPlayerY(i)
                                                'CreateProjectile GetPlayerMap(Index), Index, i, TARGET_TYPE_PLAYER, Spell(SpellNum).Projectile, Spell(SpellNum).RotateSpeed
                                                If Not CanPlayerDodgePlayer(i, Index) Then
                                                    PlayerAttackPlayer Index, i, Vital, SpellNum
                                                Else
                                                    SendActionMsg GetPlayerMap(Index), actionf("Esquivou!"), Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        For i = 1 To MAX_MAP_NPCS
                            If MapNpc(MapNum).Npc(i).Num > 0 Then
                                If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                    If isInRange(AoE, X, Y, MapNpc(MapNum).Npc(i).X, MapNpc(MapNum).Npc(i).Y) Then
                                        If CanPlayerAttackNpc(Index, i, True) Then
                                            'If Spell(SpellNum).Projectile = 0 Then SendAnimation MapNum, Spell(SpellNum).SpellAnim, MapNpc(MapNum).Npc(i).X, MapNpc(MapNum).Npc(i).X, GetPlayerDir(Index), TARGET_TYPE_NPC, i
                                            SendEffect MapNum, Spell(SpellNum).Effect, MapNpc(MapNum).Npc(i).X, MapNpc(MapNum).Npc(i).X
                                            'CreateProjectile GetPlayerMap(Index), Index, i, TARGET_TYPE_NPC, Spell(SpellNum).Projectile, Spell(SpellNum).RotateSpeed
                                            If Not CanNpcDodgePlayer(i, Index) Then
                                                PlayerAttackNpc Index, i, Vital, SpellNum
                                            Else
                                                SendActionMsg GetPlayerMap(Index), actionf("Esquivou!"), Pink, 1, MapNpc(MapNum).Npc(i).X, MapNpc(MapNum).Npc(i).Y
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Else
                        RegisterAOE Index, X, Y, SpellNum
                    End If
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> Index Then
                                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                    If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                        SpellPlayer_Effect VitalType, increment, i, Vital, SpellNum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, MapNpc(MapNum).Npc(i).X, MapNpc(MapNum).Npc(i).Y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, SpellNum, MapNum
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If TargetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
           
            If TargetType = TARGET_TYPE_PLAYER Then
                X = GetPlayerX(Target)
                Y = GetPlayerY(Target)
            Else
                X = MapNpc(MapNum).Npc(Target).X
                Y = MapNpc(MapNum).Npc(Target).Y
            End If
               
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                PlayerMsg Index, printf("Fora de alcance."), brightred
                SendClearSpellBuffer Index
                Exit Sub
            End If
           
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, GetPlayerX(Target), GetPlayerY(Target), GetPlayerDir(Index), TARGET_TYPE_PLAYER, Target
                                SendEffect MapNum, Spell(SpellNum).Effect, GetPlayerX(Target), GetPlayerY(Target)
                                CreateProjectile GetPlayerMap(Index), Index, Target, TARGET_TYPE_PLAYER, Spell(SpellNum).Projectile, Spell(SpellNum).RotateSpeed
                                If Not CanPlayerDodgePlayer(Target, Index) Then
                                    PlayerAttackPlayer Index, Target, Vital, SpellNum
                                    If Spell(SpellNum).Impact > 0 Then MakeImpact Target, Spell(SpellNum).Impact, TargetType, , Index
                                Else
                                    SendActionMsg GetPlayerMap(Index), actionf("Esquivou!"), Pink, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32)
                                End If
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, MapNpc(MapNum).Npc(Target).X, MapNpc(MapNum).Npc(Target).Y, GetPlayerDir(Index), TARGET_TYPE_NPC, Target
                                SendEffect MapNum, Spell(SpellNum).Effect, MapNpc(MapNum).Npc(Target).X, MapNpc(MapNum).Npc(Target).Y
                                CreateProjectile GetPlayerMap(Index), Index, Target, TARGET_TYPE_NPC, Spell(SpellNum).Projectile, Spell(SpellNum).RotateSpeed
                                
                                If Not CanNpcDodgePlayer(Target, Index) Then
                                    PlayerAttackNpc Index, Target, Vital, SpellNum
                                    If Spell(SpellNum).Impact > 0 Then MakeImpact Target, Spell(SpellNum).Impact, TargetType, MapNum, Index
                                Else
                                    SendActionMsg GetPlayerMap(Index), actionf("Esquivou!"), Pink, 1, MapNpc(MapNum).Npc(Target).X, MapNpc(MapNum).Npc(Target).Y
                                End If
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    End If
                   
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(Index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, SpellNum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, SpellNum
                        End If
                    Else
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(Index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, SpellNum, MapNum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, SpellNum, MapNum
                        End If
                    End If
                    
                Case SPELL_TYPE_SHUNPPO
                
                    DidCast = True
                    If TargetType = TARGET_TYPE_PLAYER And Target = Index Then Exit Sub
                    
                    If DidCast = True Then
                        'DidCast = False
                        Call SendAnimation(Player(Index).Map, Spell(SpellNum).CastAnim, Player(Index).X, Player(Index).Y, Player(Index).Dir)
                        
                        If TargetType = TARGET_TYPE_PLAYER Then
                                Select Case GetPlayerDir(Target)
                                    Case DIR_UP: Y = Y + 1
                                    Case DIR_DOWN: Y = Y - 1
                                    Case DIR_LEFT: X = X + 1
                                    Case DIR_RIGHT: X = X - 1
                                End Select
                        Else
                                Select Case MapNpc(Player(Index).Map).Npc(Target).Dir
                                    Case DIR_UP: Y = Y + 1
                                    Case DIR_DOWN: Y = Y - 1
                                    Case DIR_LEFT: X = X + 1
                                    Case DIR_RIGHT: X = X - 1
                                End Select
                        End If
                        
                        Call SendAnimation(Player(Index).Map, Spell(SpellNum).SpellAnim, X, Y, Player(Index).Dir)
                        
                        If X > 0 And X <= Map(Player(Index).Map).MaxX Then
                        If Y > 0 And Y <= Map(Player(Index).Map).MaxY Then
                        If IsTileValid(Map(Player(Index).Map).Tile(X, Y).Type) Then
                            UpdateMapBlock Player(Index).Map, Player(Index).X, Player(Index).Y, False
                            Player(Index).X = X
                            Player(Index).Y = Y
                            UpdateMapBlock Player(Index).Map, Player(Index).X, Player(Index).Y, True
                            If TargetType = TARGET_TYPE_PLAYER Then
                                Call SetPlayerDir(Index, GetPlayerDir(Target))
                            Else
                                Call SetPlayerDir(Index, MapNpc(Player(Index).Map).Npc(Target).Dir)
                            End If
                            Call SendPlayerXYToMap(Index)
                            If Spell(SpellNum).StunDuration > 0 Then
                                If TargetType = TARGET_TYPE_PLAYER Then StunPlayer Target, SpellNum
                                If TargetType = TARGET_TYPE_NPC Then StunNPC Target, Player(Index).Map, SpellNum
                            End If
                            DidCast = True
                        Else
                            Call PlayerMsg(Index, printf("Impossível se mover para local"), brightred)
                            DidCast = True
                        End If
                        End If
                        End If
                    Else
                    
                    End If
                    
            End Select
    End Select
   
    If DidCast Then
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
       
        TempPlayer(Index).SpellCD(spellslot) = GetTickCount + (Spell(SpellNum).CDTime * 1000)
        Call SendCooldown(Index, spellslot)
        
        TempPlayer(Index).spellBuffer.Spell = 0
        TempPlayer(Index).spellBuffer.Target = 0
        TempPlayer(Index).spellBuffer.Timer = 0
        TempPlayer(Index).spellBuffer.tType = 0
        SendSpellBuffer Index, 0
        'SendActionMsg MapNum, Trim$(Spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = brightgreen
            If Vital = Vitals.MP Then Colour = brightblue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, GetPlayerX(Index), GetPlayerY(Index), GetPlayerDir(Index), TARGET_TYPE_PLAYER, Index
        SendEffect GetPlayerMap(Index), Spell(SpellNum).Effect, GetPlayerX(Index), GetPlayerY(Index)
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        
        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
        
        If increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
        ElseIf Not increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
        End If
        SendVital Index, Vital
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal MapNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = brightgreen
            If Vital = Vitals.MP Then Colour = brightblue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation MapNum, Spell(SpellNum).SpellAnim, MapNpc(MapNum).Npc(Index).X, MapNpc(MapNum).Npc(Index).Y, GetPlayerDir(Index), TARGET_TYPE_NPC, Index
        SendEffect MapNum, Spell(SpellNum).Effect, MapNpc(MapNum).Npc(Index).X, MapNpc(MapNum).Npc(Index).Y
        SendActionMsg MapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(Index).X * 32, MapNpc(MapNum).Npc(Index).Y * 32
        
        ' send the sound
        SendMapSound Index, MapNpc(MapNum).Npc(Index).X, MapNpc(MapNum).Npc(Index).Y, SoundEntity.seSpell, SpellNum
        
        If increment Then
            MapNpc(MapNum).Npc(Index).Vital(Vital) = MapNpc(MapNum).Npc(Index).Vital(Vital) + Damage
        ElseIf Not increment Then
            MapNpc(MapNum).Npc(Index).Vital(Vital) = MapNpc(MapNum).Npc(Index).Vital(Vital) - Damage
        End If
        ' send update
        SendMapNpcVitals MapNum, Index
    End If
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal Index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).TempNpc(Index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).TempNpc(Index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal Index As Long, ByVal dotNum As Long)
End Sub

Public Sub HandleHoT_Player(ByVal Index As Long, ByVal hotNum As Long)
End Sub

Public Sub HandleDoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal dotNum As Long)
End Sub

Public Sub HandleHoT_Npc(ByVal MapNum As Long, ByVal Index As Long, ByVal hotNum As Long)
End Sub

Public Sub NpcStunPlayer(ByVal Index As Long, ByVal Duration As Long)
    ' check if it's a stunning spell
    If Duration > 0 And TempPlayer(Index).StunDuration = 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Duration
        TempPlayer(Index).StunTimer = GetTickCount
        
        If TempPlayer(Index).spellBuffer.Spell > 0 Then TempPlayer(Index).spellBuffer.Spell = 0
        ' send it to the index
        SendStunned Index, TargetType.TARGET_TYPE_PLAYER, TempPlayer(Index).StunDuration, GetPlayerMap(Index)
        ' tell him he's stunned
        PlayerMsg Index, printf("Você foi atordoado!"), brightred
    End If
End Sub

Public Sub ServerStunPlayer(ByVal Index As Long, ByVal Duration As Long)
        TempPlayer(Index).StunDuration = Duration
        TempPlayer(Index).StunTimer = GetTickCount
        SendStunned Index, TargetType.TARGET_TYPE_PLAYER, TempPlayer(Index).StunDuration, GetPlayerMap(Index)
End Sub

Public Sub StunPlayer(ByVal Index As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Spell(SpellNum).StunDuration
        TempPlayer(Index).StunTimer = GetTickCount
        ' cancelar skills
        If TempPlayer(Index).spellBuffer.Spell > 0 Then TempPlayer(Index).spellBuffer.Spell = 0
        ' send it to the index
        SendStunned Index, TargetType.TARGET_TYPE_PLAYER, TempPlayer(Index).StunDuration, GetPlayerMap(Index)
        ' tell him he's stunned
        PlayerMsg Index, printf("Você foi atordoado!"), brightred
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal MapNum As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(MapNum).TempNpc(Index).StunDuration = Spell(SpellNum).StunDuration
        MapNpc(MapNum).TempNpc(Index).StunTimer = GetTickCount
        SendStunned Index, TargetType.TARGET_TYPE_NPC, MapNpc(MapNum).TempNpc(Index).StunDuration, MapNum
    End If
End Sub

Sub CheckAttack(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Vital As Long, SpellNum As Long, Optional NPCCaster As Boolean = False, Optional MapIndex As Long = 0)
Dim i As Long, MapNum As Long

If NPCCaster Then
    MapNum = MapIndex
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                    If CanNpcAttackPlayer(Index, i) Then
                        If Not CanPlayerDodgeNPC(i, Index) Then
                            If TempPlayer(i).ImpactedBy <> Index Or TempPlayer(i).ImpactedTick < GetTickCount Then
                                NpcAttackPlayer Index, i, Vital, , True
                                If Spell(SpellNum).Impact > 0 Then Call MakeImpact(i, Spell(SpellNum).Impact, TARGET_TYPE_PLAYER, , Index)
                                'If SpellNum > 0 Then
                                '    If Spell(SpellNum).SpellAnim > 0 Then
                                '        Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, GetPlayerX(i), GetPlayerY(i), GetPlayerDir(Index))
                                '        SendMapSound Index, GetPlayerX(i), GetPlayerY(i), SoundEntity.seSpell, SpellNum
                                '    End If
                                'End If
                            End If
                        Else
                            SendActionMsg MapNum, actionf("Esquivou!"), Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                        End If
                    End If
                End If
            End If
        End If
    Next i
Else
    MapNum = GetPlayerMap(Index)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                    If CanPlayerAttackPlayer(Index, i, True) Then
                        If Not CanPlayerDodgePlayer(i, Index) Then
                            If TempPlayer(i).ImpactedBy <> Index Or TempPlayer(i).ImpactedTick < GetTickCount Then
                                PlayerAttackPlayer Index, i, Vital, SpellNum
                                If Spell(SpellNum).Impact > 0 Then Call MakeImpact(i, Spell(SpellNum).Impact, TARGET_TYPE_PLAYER, , Index)
                                If SpellNum > 0 Then
                                    If Spell(SpellNum).SpellAnim > 0 Then
                                        Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, GetPlayerX(i), GetPlayerY(i), GetPlayerDir(Index))
                                        SendMapSound Index, GetPlayerX(i), GetPlayerY(i), SoundEntity.seSpell, SpellNum
                                    End If
                                End If
                            End If
                        Else
                            SendActionMsg MapNum, actionf("Esquivou!"), Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(MapNum).Npc(i).Num > 0 Then
            If MapNpc(MapNum).Npc(i).X = X And MapNpc(MapNum).Npc(i).Y = Y Then
                If CanPlayerAttackNpc(Index, i, True) Then
                    If Not CanNpcDodgePlayer(i, Index) Then
                        If MapNpc(MapNum).TempNpc(i).ImpactedBy <> Index Or MapNpc(MapNum).TempNpc(i).ImpactedTick < GetTickCount Then
                            PlayerAttackNpc Index, i, Vital, SpellNum
                            If Spell(SpellNum).Impact > 0 Then Call MakeImpact(i, Spell(SpellNum).Impact, TARGET_TYPE_NPC, MapNum, Index)
                            If SpellNum > 0 Then
                                If Spell(SpellNum).SpellAnim > 0 Then
                                    Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, MapNpc(MapNum).Npc(i).X, MapNpc(MapNum).Npc(i).Y, GetPlayerDir(Index))
                                    SendMapSound Index, MapNpc(MapNum).Npc(i).X, MapNpc(MapNum).Npc(i).Y, SoundEntity.seSpell, SpellNum
                                End If
                            End If
                        End If
                    Else
                        SendActionMsg MapNum, actionf("Esquivou!"), Pink, 1, (MapNpc(MapNum).Npc(i).X * 32), (MapNpc(MapNum).Npc(i).Y * 32)
                    End If
                End If
            End If
        End If
    Next i
    
    If UZ And TempPlayer(Index).MatchIndex > 0 Then CheckResource Index, X, Y, Vital * 2
End If
End Sub
Sub NpcMakeImpact(ByVal Index As Long, ByVal ImpactValue As Byte, ByVal MapNum As Long, ByVal MapNPCNum As Long)
Dim i As Long, X As Long, Y As Long, Dir As Byte
Dim XDif, YDif As Long

    X = Player(Index).X
    Y = Player(Index).Y
    
    XDif = X - MapNpc(MapNum).Npc(MapNPCNum).X
    YDif = Y - MapNpc(MapNum).Npc(MapNPCNum).Y
    
    If XDif = 0 Then
        If YDif < 0 Then Dir = DIR_UP
        If YDif > 0 Then Dir = DIR_DOWN
    Else
        If XDif < 0 Then Dir = DIR_LEFT
        If XDif > 0 Then Dir = DIR_RIGHT
    End If
    
    For i = 1 To ImpactValue
        Select Case Dir
            Case DIR_UP: Y = Y - 1
            Case DIR_DOWN: Y = Y + 1
            Case DIR_LEFT: X = X - 1
            Case DIR_RIGHT: X = X + 1
        End Select
        
        If X > 0 And X < Map(Player(Index).Map).MaxX Then
            If Y > 0 And Y < Map(Player(Index).Map).MaxY Then
                If Map(Player(Index).Map).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE Then
                    Player(Index).X = X
                    Player(Index).Y = Y
                Else
                    Exit For
                End If
            End If
        End If
    Next i
    
    TempPlayer(Index).ImpactedBy = MapNPCNum
    TempPlayer(Index).ImpactedTick = GetTickCount + 100
    SendPlayerXYToMap Index

End Sub
Sub MakeImpact(ByVal Index As Long, ByVal ImpactValue As Byte, ByVal TargetType As Byte, Optional MapNum As Long, Optional Attacker As Long)
Dim i As Long, X As Long, Y As Long, Dir As Byte
Dim XDif, YDif As Long

If TargetType = TARGET_TYPE_PLAYER Then
    X = Player(Index).X
    Y = Player(Index).Y
    
    XDif = X - Player(Attacker).X
    YDif = Y - Player(Attacker).Y
    
    If XDif = 0 Then
        If YDif < 0 Then Dir = DIR_UP
        If YDif > 0 Then Dir = DIR_DOWN
    Else
        If XDif < 0 Then Dir = DIR_LEFT
        If XDif > 0 Then Dir = DIR_RIGHT
    End If
    
    For i = 1 To ImpactValue
        Select Case Dir
            Case DIR_UP: Y = Y - 1
            Case DIR_DOWN: Y = Y + 1
            Case DIR_LEFT: X = X - 1
            Case DIR_RIGHT: X = X + 1
        End Select
        
        If X > 0 And X < Map(Player(Index).Map).MaxX Then
        If Y > 0 And Y < Map(Player(Index).Map).MaxY Then
        If Map(Player(Index).Map).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE Then
            Player(Index).X = X
            Player(Index).Y = Y
        Else
            Exit For
        End If
        End If
        End If
    Next i
    
    TempPlayer(Index).ImpactedBy = Attacker
    TempPlayer(Index).ImpactedTick = GetTickCount + 100
    SendPlayerXYToMap Index
End If

If TargetType = TARGET_TYPE_NPC Then
    X = MapNpc(MapNum).Npc(Index).X
    Y = MapNpc(MapNum).Npc(Index).Y
    
    XDif = X - Player(Attacker).X
    YDif = Y - Player(Attacker).Y
    
    If XDif = 0 Then
        If YDif < 0 Then Dir = DIR_UP
        If YDif > 0 Then Dir = DIR_DOWN
    Else
        If XDif < 0 Then Dir = DIR_LEFT
        If XDif > 0 Then Dir = DIR_RIGHT
    End If
    
    For i = 1 To ImpactValue
        Select Case Dir
            Case DIR_UP: Y = Y - 1
            Case DIR_DOWN: Y = Y + 1
            Case DIR_LEFT: X = X - 1
            Case DIR_RIGHT: X = X + 1
        End Select
        
        If MapNum > 0 Then
        If X > 0 And X < Map(MapNum).MaxX Then
            If Y > 0 And Y < Map(MapNum).MaxY Then
                If Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE Then
                    MapNpc(MapNum).Npc(Index).X = X
                    MapNpc(MapNum).Npc(Index).Y = Y
                Else
                    Exit For
                End If
            End If
        End If
        End If
    Next i
    
    MapNpc(MapNum).TempNpc(Index).ImpactedBy = Attacker
    MapNpc(MapNum).TempNpc(Index).ImpactedTick = GetTickCount + 100
    SendMapNpcXY Index, MapNum
End If

End Sub
Sub NpcShunppo(Index As Long, MapNPCNum As Long)
    Dim MapNum As Long, NpcNum As Long
    Dim X As Long, Y As Long
    
    MapNum = Player(Index).Map
    NpcNum = MapNpc(MapNum).Npc(MapNPCNum).Num
    
    If NpcNum > 0 Then
    
    X = Player(Index).X
    Y = Player(Index).Y
    
    Call SendAnimation(Player(Index).Map, Npc(NpcNum).IA(NPCIA.Shunppo).Data(4), MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, Player(Index).Dir)
             
    Select Case GetPlayerDir(Index)
        Case DIR_UP: Y = Y + 1
        Case DIR_DOWN: Y = Y - 1
        Case DIR_LEFT: X = X + 1
        Case DIR_RIGHT: X = X - 1
    End Select

    Call SendAnimation(Player(Index).Map, Npc(NpcNum).IA(NPCIA.Shunppo).Data(4), X, Y, Player(Index).Dir)
    
    If X >= 0 And X <= Map(Player(Index).Map).MaxX Then
        If Y >= 0 And Y <= Map(Player(Index).Map).MaxY Then
            If Map(Player(Index).Map).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE Then
                UpdateMapBlock Player(Index).Map, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, False
                MapNpc(MapNum).Npc(MapNPCNum).X = X
                MapNpc(MapNum).Npc(MapNPCNum).Y = Y
                UpdateMapBlock Player(Index).Map, MapNpc(MapNum).Npc(MapNPCNum).X, MapNpc(MapNum).Npc(MapNPCNum).Y, True
                MapNpc(MapNum).Npc(MapNPCNum).Dir = GetPlayerDir(Index)
                Call SendMapNpcXY(MapNPCNum, MapNum)
                If Npc(NpcNum).IA(NPCIA.Shunppo).Data(3) > 0 Then
                    NpcStunPlayer Index, Npc(NpcNum).IA(NPCIA.Shunppo).Data(3)
                End If
            End If
        End If
    End If
    End If
    
End Sub
Function IsTileValid(TileNum As Byte) As Boolean
    IsTileValid = True
    Select Case TileNum
        Case TileType.TILE_TYPE_BANK: IsTileValid = False
        Case TileType.tile_type_blocked: IsTileValid = False
        Case TileType.TILE_TYPE_EVENT: IsTileValid = False
        Case TileType.TILE_TYPE_RESOURCE: IsTileValid = False
        Case TileType.TILE_TYPE_SHOP: IsTileValid = False
        Case TileType.TILE_TYPE_SLIDE: IsTileValid = False
        Case TileType.TILE_TYPE_TRAP: IsTileValid = False
        Case TileType.tile_type_warp: IsTileValid = False
    End Select
End Function
