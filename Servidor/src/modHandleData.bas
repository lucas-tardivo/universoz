Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CRequestSwitchesAndVariables) = GetAddress(AddressOf HandleRequestSwitchesAndVariables)
    HandleDataSub(CSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(CSaveEventData) = GetAddress(AddressOf Events_HandleSaveEventData)
    HandleDataSub(CRequestEventData) = GetAddress(AddressOf Events_HandleRequestEventData)
    HandleDataSub(CRequestEventsData) = GetAddress(AddressOf Events_HandleRequestEventsData)
    HandleDataSub(CRequestEditEvents) = GetAddress(AddressOf Events_HandleRequestEditEvents)
    HandleDataSub(CChooseEventOption) = GetAddress(AddressOf Events_HandleChooseEventOption)
    HandleDataSub(CRequestEditEffect) = GetAddress(AddressOf HandleRequestEditEffect)
    HandleDataSub(CSaveEffect) = GetAddress(AddressOf HandleSaveEffect)
    HandleDataSub(CRequestEffects) = GetAddress(AddressOf HandleRequestEffects)
    HandleDataSub(CTarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(CEditNews) = GetAddress(AddressOf HandleEditNews)
    HandleDataSub(CRequestEditNews) = GetAddress(AddressOf HandleRequestEditNews)
    HandleDataSub(CRequestNews) = GetAddress(AddressOf HandleRequestNews)
    HandleDataSub(CDevSuite) = GetAddress(AddressOf HandleDevSuite)
    HandleDataSub(COnDeath) = GetAddress(AddressOf HandleOnDeath)
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditquest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSavequest)
    HandleDataSub(CQuestInfo) = GetAddress(AddressOf HandleQuestInfo)
    HandleDataSub(CUpgrade) = GetAddress(AddressOf HandleUpgrade)
    HandleDataSub(CSellPlanet) = GetAddress(AddressOf HandleSellPlanet)
    HandleDataSub(CEnterGravity) = GetAddress(AddressOf HandleEnterGravity)
    HandleDataSub(CCompleteTutorial) = GetAddress(AddressOf HandleCompleteTutorial)
    HandleDataSub(CFeedback) = GetAddress(AddressOf HandleFeedback)
    HandleDataSub(CCreateGuild) = GetAddress(AddressOf HandleCreateGuild)
    HandleDataSub(CGuildAction) = GetAddress(AddressOf HandleGuildAction)
    HandleDataSub(CChallengeArena) = GetAddress(AddressOf HandleChallengeArena)
    HandleDataSub(CAntiHackData) = GetAddress(AddressOf HandleAntiHackData)
    HandleDataSub(CPlanetChange) = GetAddress(AddressOf HandlePlanetChange)
    HandleDataSub(CConfirmation) = GetAddress(AddressOf HandleConfirmation)
    HandleDataSub(CSellEsp) = GetAddress(AddressOf HandleSellEspeciaria)
    HandleDataSub(CSupport) = GetAddress(AddressOf HandleSupport)
End Sub

Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    ' Add one to the incoming packet number.
    PacketsIn = PacketsIn + 1
    
    CallWindowProc HandleDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMSG(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Then
                Call AlertMSG(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(n) Then
                    Call AlertMSG(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.", ChatPlayer)
                
                ' Load the player
                Call LoadPlayer(Index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(Player(Index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar Index
                Else
                    ' send new char shit
                    If Not IsPlaying(Index) Then
                        Call SendNewCharClasses(Index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", ChatPlayer)
            Else
                Call AlertMSG(Index, "Sorry, that account name is already taken!")
            End If
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMSG(Index, "The name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMSG(Index, "Essa conta não existe.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMSG(Index, "Senha incorreta.")
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(Index, Name)

            If LenB(Trim$(Player(Index).Name)) > 0 Then
                Call DeleteName(Player(Index).Name)
            End If

            Call ClearPlayer(Index)
            ' Everything went ok
            Call Kill(App.path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AlertMSG(Index, "Sua conta foi deletada.")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Trim(Buffer.ReadString)
            Password = Buffer.ReadString
            
            ' Check versions
            If Buffer.ReadLong < App.Major Or Buffer.ReadLong < App.Minor Or Buffer.ReadLong < App.Revision Then
                Call AlertMSG(Index, printf("Versão desatualizada, abra o jogo novamente na nova versão!"))
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMSG(Index, printf("O Servidor está desligando ou entrando em manutenção."))
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMSG(Index, printf("Seu nome e senha devem ter no minimo 3 caracteres"))
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMSG(Index, printf("Essa conta não existe."))
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMSG(Index, printf("Sua senha está incorreta."))
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMSG(Index, printf("Esta conta já está online."))
                Exit Sub
            End If
            
            If IsBanned(GetPlayerIP(Index), Name) Then
                Call AlertMSG(Index, printf("Sua conta foi banida!."))
                Exit Sub
            End If
            
            ' Load the player
            Call LoadPlayer(Index, Name)
            
            If frmServer.cfkClosed.Value = 1 Then
                If Player(Index).Access = 0 Then
                    Call AlertMSG(Index, printf("Servidor fechado para administradores"))
                    Call ClearPlayer(Index)
                    Exit Sub
                End If
            End If
            
            ClearBank Index
            LoadBank Index, Name
            ' Check if character data has been created
            If LenB(Trim$(Player(Index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar Index
            Else
                ' send new char shit
                If Not IsPlaying(Index) Then
                    Call SendNewCharClasses(Index)
                End If
            End If
            
            ' Show the player up on the socket status
            Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", ChatPlayer)
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim Hair As Byte
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        Sprite = Buffer.ReadLong
        Hair = Buffer.ReadByte

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMSG(Index, printf("O nome de seu personagem deve ter no mínimo 3 letras."), False)
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMSG(Index, printf("Nome inválido, apenas letras, numeros, espaços e _ são válidos"), False)
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(Index) Then
            Call AlertMSG(Index, printf("Este personagem ja existe!"))
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMSG(Index, printf("Desculpe, este nome já está em uso!"), False)
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(Index, Name, Sex, Class, Sprite, Hair)
        ' log them in!!
        HandleUseChar Index
        
        Set Buffer = Nothing
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next
    
    If Mid(LCase(Msg), 1, 9) = "/reportar" Then
        Call Report(Index, Msg)
        Exit Sub
    End If
    
    If LCase(Msg) = printf("/sair") And Not (GetPlayerMap(Index) > 5 And GetPlayerMap(Index) <= 15) Then
        If Player(Index).GravityHours > 0 Then
            Player(Index).GravityInit = vbNullString
            Player(Index).GravityHours = 0
            PlayerWarp Index, START_MAP, Int(Map(START_MAP).MaxX / 2), Int(Map(START_MAP).MaxY / 2)
            PlayerMsg Index, "Você cancelou seu treinamento!", brightred
        End If
        Exit Sub
    End If
    
    Call TextAdd("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", ChatMap)
    Call SayMsg_Map(GetPlayerMap(Index), Index, Msg, QBColor(White))
    Call SendChatBubble(GetPlayerMap(Index), Index, TARGET_TYPE_PLAYER, Msg, White)
    
    If LCase(Msg) = printf("saia dai shenlong") Then
        Call CheckShenlong(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    End If
    
    If Mid(LCase(Msg), 1, Len(printf("eu desejo"))) = printf("eu desejo") Then
        If ShenlongActive = 1 Then
            If ShenlongTick + 20000 < GetTickCount Then
                Call DoWish(Index, Msg)
            End If
        End If
    End If
    
    
    Set Buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next
    
    If Player(Index).Guild = 0 Then
        PlayerMsg Index, "Você não está em nenhuma guild!", brightred
        Exit Sub
    End If
    
    Call TextAdd("Guild #" & Player(Index).Guild & ": " & GetPlayerName(Index) & " " & Msg, ChatEmote)
    Call GuildMsg(Player(Index).Guild, GetPlayerName(Index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    'If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_PRISON Then
    '    PlayerMsg Index, printf("Você não pode falar no global dentro de uma prisão!"), brightred
    '    Exit Sub
    'End If
    
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 And AscW(Mid$(Msg, i, 1)) < 192 Then
            Exit Sub
        End If

    Next

    s = "[Global]" & GetPlayerName(Index) & ": " & Msg
    Call SayMsg_Global(Index, Msg, QBColor(White))
    Call TextAdd(s, ChatGlobal)
    
    If InEvent = AutomaticEvents.PalavraMagica Then
        If LCase(Msg) = LCase(MagicWord) Then
            Call GiveReward(Index)
            InEvent = AutomaticEvents.None
        End If
    End If
    
    If InEvent = AutomaticEvents.Aventura Then
        If LCase(Msg) = "entrar" Then
            Call PlayerWarp(Index, AventuraConfig(AventuraNum).MapNum, AventuraConfig(AventuraNum).X, AventuraConfig(AventuraNum).Y)
            Call InitEvent(Index, AventuraConfig(AventuraNum).EventNum)
        End If
    End If
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> Index Then
        If MsgTo > 0 Then
            Call TextAdd(GetPlayerName(Index) & " says " & GetPlayerName(MsgTo) & ", " & Msg & "'", ChatPlayer)
            Call PlayerMsg(MsgTo, printf("%s diz á voce, '%s'", GetPlayerName(Index) & "," & Msg), TellColor)
            Call PlayerMsg(Index, printf("Você diz %s, '%s'", GetPlayerName(MsgTo) & "," & Msg), TellColor)
        Else
            Call PlayerMsg(Index, printf("Jogador não está online."), White)
        End If

    Else
        Call PlayerMsg(Index, printf("Você não pode mandar uma mensagem para si mesmo."), brightred)
    End If
    
    Set Buffer = Nothing

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    movement = Buffer.ReadLong 'CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Set Buffer = Nothing
    
    TempPlayer(Index).LastMove = GetTickCount

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(Index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(Index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(Index).StunDuration > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(Index).InShop > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(Index) <> tmpX Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    If GetPlayerY(Index) <> tmpY Then
        SendPlayerXY (Index)
        Exit Sub
    End If
    
    ' cant move if chatting
    If TempPlayer(Index).CurrentEvent > 0 Then
        TempPlayer(Index).CurrentEvent = -1
        Call Events_SendEventQuit(Index)
    End If

    Call PlayerMove(Index, Dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim InvNum As Long
Dim Buffer As clsBuffer
    
    ' get inventory slot number
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    InvNum = Buffer.ReadLong
    Set Buffer = Nothing

    UseItem Index, InvNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim X As Long, Y As Long
    Dim Shoot As Boolean
    
    ' can't attack whilst casting
    If TempPlayer(Index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub
    
    If GetPlayerMap(Index) = ViagemMap And UZ Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack Index
    
    If Shoot = True Then
        Select Case TempPlayer(Index).TargetType
            Case TARGET_TYPE_NPC: TryPlayerShootNpc Index, TempPlayer(Index).Target
            Case TARGET_TYPE_PLAYER: TryPlayerShootPlayer Index, TempPlayer(Index).Target
        End Select
        Exit Sub
    End If
    
    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i
        ' Make sure we dont try to attack ourselves
        If TempIndex <> Index Then
            TryPlayerAttackPlayer Index, i
        End If
    Next
    
    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc Index, i
    Next

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
    
    CheckResource Index, X, Y
    CheckEvent Index, X, Y
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim Buffer As clsBuffer
Dim sMes As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PointType = Buffer.ReadByte 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType >= Stats.Stat_Count) Then
        Exit Sub
    End If
    
    If TempPlayer(Index).Trans > 0 Then
        PlayerMsg Index, printf("Você não pode distribuir pontos enquanto estiver transformado!"), brightred
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(Index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerRawStat(Index, PointType) >= MAX_STAT_LEVELS Then
            PlayerMsg Index, printf("Este stat chegou ao nível máximo!"), brightred
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)
        
        Dim PointsAdd As Long
        PointsAdd = 1 'Int((GetPlayerStatNextLevel(Index, PointType) - GetPlayerStatLastLevel(Index, PointType)) * 0.05)
        'If PointsAdd < 1 Then PointsAdd = 1
        'If PointsAdd > GetPlayerPOINTS(Index) Then PointsAdd = GetPlayerPOINTS(Index) - 1
        'If PointsAdd = 0 Then Exit Sub

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                'Call SetPlayerStatPoints(Index, Stats.Strength, GetPlayerStatPoints(Index, Stats.Strength) + PointsAdd)
                Call SetPlayerStat(Index, Strength, GetPlayerRawStat(Index, Strength) + PointsAdd)
                sMes = printf("Força")
            Case Stats.Endurance
                'Call SetPlayerStatPoints(Index, Stats.Endurance, GetPlayerStatPoints(Index, Stats.Endurance) + PointsAdd)
                Call SetPlayerStat(Index, Endurance, GetPlayerRawStat(Index, Endurance) + PointsAdd)
                Call SendVital(Index, HP)
                sMes = printf("Constituição")
            Case Stats.Intelligence
                'Call SetPlayerStatPoints(Index, Stats.Intelligence, GetPlayerStatPoints(Index, Stats.Intelligence) + PointsAdd)
                Call SetPlayerStat(Index, Intelligence, GetPlayerRawStat(Index, Intelligence) + PointsAdd)
                Call SendVital(Index, MP)
                sMes = "KI"
            Case Stats.agility
                'Call SetPlayerStatPoints(Index, Stats.agility, GetPlayerStatPoints(Index, Stats.agility) + PointsAdd)
                Call SetPlayerStat(Index, agility, Player(Index).stat(agility) + PointsAdd)
                sMes = printf("Destreza")
            Case Stats.Willpower
                'Call SetPlayerStatPoints(Index, Stats.Willpower, GetPlayerStatPoints(Index, Stats.Willpower) + PointsAdd)
                Call SetPlayerStat(Index, Willpower, GetPlayerRawStat(Index, Willpower) + PointsAdd)
                sMes = printf("Técnica")
        End Select
        
        SendActionMsg GetPlayerMap(Index), actionf("+%d pontos em %s", PointsAdd & "," & sMes), White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        'SendActionMsg GetPlayerMap(Index), "+6 PDL", brightgreen, 1, (GetPlayerX(Index) * 32) + 32, (GetPlayerY(Index) * 32) + 32
        
    Else
        Exit Sub
    End If

    ' Send the update
    'Call SendStats(Index)
    SendStats Index
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call TextAdd(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ChatSystem)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp to yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Call TextAdd(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(Index) & ".", ChatSystem)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp yourself to yourself!", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
    Call PlayerMsg(Index, "You have been warped to map #" & n, brightblue)
    Call TextAdd(GetPlayerName(Index) & " warped to map #" & n & ".", ChatSystem)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The sprite
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    Call SetPlayerSprite(Index, n)
    Call SendPlayerData(Index)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim MapNum As Long
    Dim X As Long
    Dim Y As Long, z As Long, w As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Index)
    i = Map(MapNum).Revision + 1
    Call ClearMap(MapNum)
    
    Map(MapNum).Name = Buffer.ReadString
    Map(MapNum).Music = Buffer.ReadString
    Map(MapNum).BGS = Buffer.ReadString
    Map(MapNum).Revision = i
    Map(MapNum).Moral = Buffer.ReadByte
    Map(MapNum).Up = Buffer.ReadLong
    Map(MapNum).Down = Buffer.ReadLong
    Map(MapNum).Left = Buffer.ReadLong
    Map(MapNum).Right = Buffer.ReadLong
    Map(MapNum).BootMap = Buffer.ReadLong
    Map(MapNum).BootX = Buffer.ReadByte
    Map(MapNum).BootY = Buffer.ReadByte
    
    Map(MapNum).Weather = Buffer.ReadLong
    Map(MapNum).WeatherIntensity = Buffer.ReadLong
    
    Map(MapNum).Fog = Buffer.ReadLong
    Map(MapNum).FogSpeed = Buffer.ReadLong
    Map(MapNum).FogOpacity = Buffer.ReadLong
    Map(MapNum).FogDir = Buffer.ReadLong
    
    Map(MapNum).Red = Buffer.ReadLong
    Map(MapNum).Green = Buffer.ReadLong
    Map(MapNum).Blue = Buffer.ReadLong
    Map(MapNum).Alpha = Buffer.ReadLong
    
    Map(MapNum).MaxX = Buffer.ReadByte
    Map(MapNum).MaxY = Buffer.ReadByte
    
    Map(MapNum).Fly = Buffer.ReadByte
    Map(MapNum).Ambiente = Buffer.ReadByte
    
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).Tile(X, Y).Layer(i).X = Buffer.ReadLong
                Map(MapNum).Tile(X, Y).Layer(i).Y = Buffer.ReadLong
                Map(MapNum).Tile(X, Y).Layer(i).Tileset = Buffer.ReadLong
            Next
            For z = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).Tile(X, Y).Autotile(z) = Buffer.ReadLong
            Next
            Map(MapNum).Tile(X, Y).Type = Buffer.ReadByte
            Map(MapNum).Tile(X, Y).data1 = Buffer.ReadLong
            Map(MapNum).Tile(X, Y).data2 = Buffer.ReadLong
            Map(MapNum).Tile(X, Y).data3 = Buffer.ReadLong
            Map(MapNum).Tile(X, Y).Data4 = Buffer.ReadString
            Map(MapNum).Tile(X, Y).DirBlock = Buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(X) = Buffer.ReadLong
        Map(MapNum).NpcSpawnType(X) = Buffer.ReadLong
        Call ClearMapNpc(X, MapNum)
    Next
    Map(MapNum).Panorama = Buffer.ReadLong
    
    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    ' Save the map
    Call SaveMap(MapNum)
    Call MapCache_Create(MapNum)
    Call CacheResources(MapNum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i
    
    Call CacheMapBlocks(MapNum)

    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Get yes/no value
    s = Buffer.ReadLong 'Parse(1)
    Set Buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Or (UZ And (GetPlayerMap(Index) >= PlanetStart And GetPlayerMap(Index) <= PlanetStart + MAX_PLANET_BASE + GetMaxPlayerPlanets)) Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If

    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)

    If GetPlayerMap(Index) > 0 Then
    'send Resource cache
    'For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index
    'Next
    End If

    TempPlayer(Index).GettingMap = NO
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapDone
    SendDataTo Index, Buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(Index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim InvNum As Long
    Dim Amount As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    InvNum = Buffer.ReadLong 'CLng(Parse(1))
    Amount = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing
    
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(Index, InvNum) < 1 Or GetPlayerInvItemNum(Index, InvNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(Index, InvNum)).CantDrop = 1 Then
        If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ItemType.ITEM_TYPE_TITULO Then
            PlayerMsg Index, printf("Este item não pode ser dropado!"), brightred
            Exit Sub
        End If
        TempPlayer(Index).ConfirmationVar = InvNum
        SendConfirmation Index, "Este item não pode ser deixado no chão, deseja destruí-lo? (NOTE: SE HÁ MAIS DE UM ITEM, TODOS SERÃO DESTRUÍDOS)", ConfirmType.DestroyItem
        Exit Sub
    End If
    
    If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, InvNum)).Stackable > 0 Then
        If Amount < 1 Or Amount > GetPlayerInvItemValue(Index, InvNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(Index, InvNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next

    CacheResources GetPlayerMap(Index)
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call TextAdd(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ChatSystem)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapReport
    
    For i = 1 To MAX_MAPS
        Buffer.WriteString Trim$(Map(i).Name)
    Next
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(Index) & "!", White)
                Call TextAdd(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ChatSystem)
                Call AlertMSG(n, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot kick yourself!", White)
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim F As Long
    Dim s As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    F = FreeFile
    Open App.path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        Call PlayerMsg(Index, n & ": Banned IP " & s & " by " & Name, White)
        n = n + 1
    Loop

    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String
    Dim File As Long
    Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    Kill filename
    Call PlayerMsg(Index, "Ban list destroyed.", White)
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n, Index)
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot ban yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEditMap
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SItemEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call TextAdd(GetPlayerName(Index) & " saved item #" & n & ".", ChatSystem)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimationEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call TextAdd(GetPlayerName(Index) & " saved Animation #" & n & ".", ChatSystem)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim NpcNum As Long
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    NpcNum = Buffer.ReadLong

    ' Prevent hacking
    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(Npc(NpcNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = Buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(Npc(NpcNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(NpcNum)
    Call SaveNpc(NpcNum)
    Call TextAdd(GetPlayerName(Index) & " saved Npc #" & NpcNum & ".", ChatSystem)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call TextAdd(GetPlayerName(Index) & " saved Resource #" & ResourceNum & ".", ChatSystem)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SShopEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    shopNum = Buffer.ReadLong

    ' Prevent hacking
    If shopNum < 0 Or shopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set Buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(shopNum)
    Call SaveShop(shopNum)
    Call TextAdd(GetPlayerName(Index) & " saving shop #" & shopNum & ".", ChatSystem)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    SpellNum = Buffer.ReadLong

    ' Prevent hacking
    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)
    Call TextAdd(GetPlayerName(Index) & " saved Spell #" & SpellNum & ".", ChatSystem)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The access
    i = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
                Call PlayerMsg(Index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", brightblue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call TextAdd(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ChatSystem)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Invalid access level.", Red)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(Buffer.ReadString) 'Parse(1))
    SaveOptions
    Set Buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call TextAdd(GetPlayerName(Index) & " changed MOTD to: " & Options.MOTD, ChatSystem)
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(Index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    ' set the spell buffer before castin
    
    If (GetPlayerMap(Index) = ViagemMap And UZ) Or Player(Index).GravityHours > 0 Then Exit Sub
    Call BufferSpell(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(Index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchInvSlots Index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        PlayerMsg Index, printf("Você não pode trocar a spell de lugar enquanto conjura."), brightred
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(Index).SpellCD(n) > GetTickCount Then
            PlayerMsg Index, printf("Você não pode trocar uma spell que está em intervalo."), brightred
            Exit Sub
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    If oldSlot < 1 Or oldSlot > MAX_PLAYER_SPELLS Or newSlot < 1 Or newSlot > MAX_PLAYER_SPELLS Then Exit Sub
    PlayerSwitchSpellSlots Index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendPing
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleUnequip(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerUnequipItem Index, Buffer.ReadLong
    Set Buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData Index
End Sub

Sub HandleRequestItems(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems Index
End Sub

Sub HandleRequestAnimations(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations Index
End Sub

Sub HandleRequestNPCS(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs Index
End Sub

Sub HandleRequestResources(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources Index
End Sub

Sub HandleRequestSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells Index
End Sub

Sub HandleRequestShops(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops Index
End Sub

Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    Dim ItemTo As String
    Dim Msg As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' item
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong
    ItemTo = Buffer.ReadString
    Msg = Buffer.ReadByte
        
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub
    
    If ItemTo = "Drop" Then
        SpawnItem tmpItem, tmpAmount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index)
    Else
        Dim i As Long
        i = FindPlayer(ItemTo)
        
        If i > 0 Then
            GiveInvItem i, tmpItem, tmpAmount, True
            SavePlayer i
            If Msg = 0 Then
                GlobalMsg GetPlayerName(Index) & " presenteou " & GetPlayerName(i) & " com " & tmpAmount & " " & Trim$(Item(tmpItem).Name) & "(s)", Yellow
            Else
                PlayerMsg Index, "Você presenteou este jogador com sucesso!", White
            End If
        Else
            PlayerMsg Index, printf("Jogador não está online."), White
        End If
    End If
    Set Buffer = Nothing
End Sub

Sub HandleForgetSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim spellslot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellslot = Buffer.ReadLong
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, printf("Você não pode esquecer uma magia que está no intervalo!"), brightred
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(Index).spellBuffer.Spell = spellslot Then
        PlayerMsg Index, printf("Você não pode esquecer uma magia que está sendo conjurada!"), brightred
        Exit Sub
    End If
    
    Player(Index).Spell(spellslot) = 0
    SendPlayerSpells Index
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(Index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim itemamount As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopslot = Buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(Index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
            
        ' check has the cost item
        For i = 1 To 5
            If .costitem(i) > 0 Then
                Dim CheckType As Byte, Success As Boolean
                Success = True
                If Item(.costitem(i)).Type <> ItemType.ITEM_TYPE_TITULO Then
                    CheckType = 0
                Else
                    Dim TitleLevel As Long
                    TitleLevel = Item(.costitem(i)).LevelReq
                    If TitleLevel = 0 Then
                        CheckType = 0
                    Else
                        CheckType = 1
                    End If
                End If
                If CheckType = 0 Then
                    itemamount = HasItem(Index, .costitem(i))
                    If itemamount = 0 Or itemamount < .costvalue(i) Then
                        PlayerMsg Index, printf("Você não tem os requisitos para comprar este item."), brightred
                        Success = False
                    End If
                Else
                    If Player(Index).Titulo > 0 Then
                        If Item(Player(Index).Titulo).LevelReq < TitleLevel Then
                            PlayerMsg Index, printf("O título que você está usando não é o suficiente para comprar este item!."), brightred
                            Success = False
                        End If
                    Else
                        PlayerMsg Index, printf("Você não tem o título necessário para comprar este item! Lembre-se de equipar ele!."), brightred
                        Success = False
                    End If
                End If
                If Item(.Item).Type = ItemType.ITEM_TYPE_SPELL Then
                    If HasSpell(Index, Item(.Item).data1) Then
                        PlayerMsg Index, "Você já tem essa magia!", brightred
                        Success = False
                    End If
                End If
                If Success = False Then
                    ResetShopAction Index
                    Exit Sub
                End If
            End If
        Next i
            
        For i = 1 To 5
            If .costitem(i) > 0 Then
                ' it's fine, let's go ahead
                If Item(.costitem(i)).Type <> ItemType.ITEM_TYPE_TITULO Then
                    TakeInvItem Index, .costitem(i), .costvalue(i)
                End If
            End If
        Next i
        GiveInvItem Index, .Item, .ItemValue
    End With
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, printf("Troca efetuada com sucesso."), brightgreen
    SendPlaySound Index, "pagamento conclido.mp3"
    ResetShopAction Index
    
    Set Buffer = Nothing
End Sub

Sub HandleSellItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim ItemNum As Long
    Dim Price As Long
    Dim multiplier As Double
    Dim Amount As Long
    
    If TempPlayer(Index).InShop < 1 Or TempPlayer(Index).InShop > MAX_SHOPS Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(Index, invSlot) < 1 Or GetPlayerInvItemNum(Index, invSlot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    ItemNum = GetPlayerInvItemNum(Index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(Index).InShop).BuyRate / 100
    Price = Item(ItemNum).Price * multiplier
    
    ' item has cost?
    If Price <= 0 Then
        PlayerMsg Index, printf("Este item não pode ser vendido!."), brightred
        ResetShopAction Index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem Index, ItemNum, 1
    GiveInvItem Index, 26, Price
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, printf("Troca efetuada com sucesso."), brightgreen
    SendPlaySound Index, "pagamento conclido.mp3"
    
    ResetShopAction Index
    
    Set Buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    
    PlayerSwitchBankSlots Index, oldSlot, newSlot
    
    Set Buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim BankSlot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    TakeBankItem Index, BankSlot, Amount
    
    Set Buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    GiveBankItem Index, invSlot, Amount
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SaveBank Index
    SavePlayer Index
    
    TempPlayer(Index).InBank = False
    
    Set Buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs
    If TempPlayer(Index).TargetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(Index).Target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = Index Then
        PlayerMsg Index, printf("Você não pode efetuar uma troca com si mesmo."), brightred
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(Index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).X
    tY = Player(tradeTarget).Y
    sX = Player(Index).X
    sY = Player(Index).Y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg Index, printf("Você precisa estar próximo do jogador para efetuar uma troca."), brightred
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg Index, printf("Você precisa estar próximo do jogador para efetuar uma troca."), brightred
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg Index, printf("Jogador ocupado."), brightred
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = Index
    SendTradeRequest tradeTarget, Index
End Sub

Sub HandleAcceptTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim i As Long

    If TempPlayer(Index).InTrade > 0 Then
        TempPlayer(Index).TradeRequest = 0
    Else
        tradeTarget = TempPlayer(Index).TradeRequest
        ' let them know they're trading
        PlayerMsg Index, printf("Você aceitou a troca de %s.", Trim$(GetPlayerName(tradeTarget))), brightgreen
        PlayerMsg tradeTarget, printf("%s aceitou sua troca.", Trim$(GetPlayerName(Index))), brightgreen
        ' clear the tradeRequest server-side
        TempPlayer(Index).TradeRequest = 0
        TempPlayer(tradeTarget).TradeRequest = 0
        ' set that they're trading with each other
        TempPlayer(Index).InTrade = tradeTarget
        TempPlayer(tradeTarget).InTrade = Index
        ' clear out their trade offers
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).Num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next
        ' Used to init the trade window clientside
        SendTrade Index, tradeTarget
        SendTrade tradeTarget, Index
        ' Send the offer data - Used to clear their client
        SendTradeUpdate Index, 0
        SendTradeUpdate Index, 1
        SendTradeUpdate tradeTarget, 0
        SendTradeUpdate tradeTarget, 1
    End If
End Sub

Sub HandleDeclineTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(Index).TradeRequest, printf("%s negou seu pedido de troca.", GetPlayerName(Index)), brightred
    PlayerMsg Index, printf("Você negou o pedido de troca."), brightred
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim ItemNum As Long
    
    TempPlayer(Index).AcceptTrade = True
    
    tradeTarget = TempPlayer(Index).InTrade
        
    If tradeTarget > 0 Then
    
        ' if not both of them accept, then exit
        If Not TempPlayer(tradeTarget).AcceptTrade Then
            SendTradeStatus Index, 2
            SendTradeStatus tradeTarget, 1
            Exit Sub
        End If
    
        ' take their items
        For i = 1 To MAX_INV
            ' player
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                ItemNum = Player(Index).Inv(TempPlayer(Index).TradeOffer(i).Num).Num
                If ItemNum > 0 Then
                    ' store temp
                    tmpTradeItem(i).Num = ItemNum
                    tmpTradeItem(i).Value = TempPlayer(Index).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot Index, TempPlayer(Index).TradeOffer(i).Num, tmpTradeItem(i).Value
                End If
            End If
            ' target
            If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
                ItemNum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                If ItemNum > 0 Then
                    ' store temp
                    tmpTradeItem2(i).Num = ItemNum
                    tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, tmpTradeItem2(i).Value
                End If
            End If
        Next
    
        ' taken all items. now they can't not get items because of no inventory space.
        For i = 1 To MAX_INV
            ' player
            If tmpTradeItem2(i).Num > 0 Then
                ' give away!
                GiveInvItem Index, tmpTradeItem2(i).Num, tmpTradeItem2(i).Value, False
            End If
            ' target
            If tmpTradeItem(i).Num > 0 Then
                ' give away!
                GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).Value, False
            End If
        Next
    
        SendInventory Index
        SendInventory tradeTarget
    
        ' they now have all the items. Clear out values + let them out of the trade.
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).Num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(Index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg Index, printf("Troca completa."), brightgreen
        PlayerMsg tradeTarget, printf("Troca completa."), brightgreen
    
        SendCloseTrade Index
        SendCloseTrade tradeTarget
            
    End If
End Sub

Sub HandleDeclineTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(Index).InTrade
    
    If tradeTarget > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).Num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(Index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg Index, printf("Você negou a troca."), brightred
        PlayerMsg tradeTarget, printf("%s negou a troca.", GetPlayerName(Index)), brightred
    
        SendCloseTrade Index
        SendCloseTrade tradeTarget
    End If
End Sub

Sub HandleTradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim ItemNum As Long
    Dim i As Long
    
    If (TempPlayer(Index).InTrade > 0) Then
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    ItemNum = GetPlayerInvItemNum(Index, invSlot)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invSlot) Then
        Exit Sub
    End If
    
    If Item(ItemNum).CantDrop = 1 Then
        PlayerMsg Index, printf("Você não pode trocar este item!"), brightred
        Exit Sub
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable > 0 Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invSlot Then
                ' add amount
                TempPlayer(Index).TradeOffer(i).Value = TempPlayer(Index).TradeOffer(i).Value + Amount
                ' clamp to limits
                If TempPlayer(Index).TradeOffer(i).Value > GetPlayerInvItemValue(Index, invSlot) Then
                    TempPlayer(Index).TradeOffer(i).Value = GetPlayerInvItemValue(Index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(Index).AcceptTrade = False
                TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
                
                SendTradeStatus Index, 0
                SendTradeStatus TempPlayer(Index).InTrade, 0
                
                SendTradeUpdate Index, 0
                SendTradeUpdate TempPlayer(Index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invSlot Then
                PlayerMsg Index, printf("Você já ofereceu esse item."), brightred
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(Index).TradeOffer(EmptySlot).Num = invSlot
    TempPlayer(Index).TradeOffer(EmptySlot).Value = Amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(Index).AcceptTrade = False
    TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
    End If

End Sub

Sub HandleUntradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeSlot = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(Index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(Index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(Index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(Index).AcceptTrade Then TempPlayer(Index).AcceptTrade = False
    If TempPlayer(TempPlayer(Index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim sType As Long
    Dim slot As Long
    Dim hotbarNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    sType = Buffer.ReadLong
    slot = Buffer.ReadLong
    hotbarNum = Buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(Index).Hotbar(hotbarNum).slot = 0
            Player(Index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If slot > 0 And slot <= MAX_INV Then
                If Player(Index).Inv(slot).Num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(Index, slot)).Name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum).slot = Player(Index).Inv(slot).Num
                        Player(Index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If slot > 0 And slot <= MAX_PLAYER_SPELLS Then
                If Player(Index).Spell(slot) > 0 Then
                    If Len(Trim$(Spell(Player(Index).Spell(slot)).Name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum).slot = Player(Index).Spell(slot)
                        Player(Index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar Index
    
    Set Buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim slot As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    slot = Buffer.ReadLong
    
    Select Case Player(Index).Hotbar(slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(Index).Inv(i).Num > 0 Then
                    If Player(Index).Inv(i).Num = Player(Index).Hotbar(slot).slot Then
                        If Item(Player(Index).Inv(i).Num).Type = ITEM_TYPE_CONSUME Then
                            Player(Index).Hotbar(slot).slot = 0
                            Player(Index).Hotbar(slot).sType = 0
                            SendHotbar Index
                        End If
                        UseItem Index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(Index).Spell(i) > 0 Then
                    If Player(Index).Spell(i) = Player(Index).Hotbar(slot).slot Then
                        BufferSpell Index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Set Buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' make sure it's a valid target
    If TempPlayer(Index).TargetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(Index).Target = Index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(Index).Target) Or Not IsPlaying(TempPlayer(Index).Target) Then Exit Sub
    If GetPlayerMap(TempPlayer(Index).Target) <> GetPlayerMap(Index) Then Exit Sub
    
    ' init the request
    Party_Invite Index, TempPlayer(Index).Target
End Sub

Sub HandleAcceptParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteAccept TempPlayer(Index).partyInvite, Index
End Sub

Sub HandleDeclineParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(Index).partyInvite, Index
End Sub

Sub HandlePartyLeave(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave Index
End Sub
Sub HandleRequestSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSwitchesAndVariables (Index)
End Sub

Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = Buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = Buffer.ReadString
    Next
    
    SaveSwitches
    SaveVariables
    
    Set Buffer = Nothing
    
    SendSwitchesAndVariables 0, True
End Sub

Public Sub Events_HandleChooseEventOption(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, Opt As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    
    Opt = Buffer.ReadLong
    Call DoEventLogic(Index, Opt)
    
    Set Buffer = Nothing
End Sub

Public Sub Events_HandleSaveEventData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim EIndex As Long, s As Long, SCount As Long, D As Long, DCount As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

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
        For s = 1 To SCount
            With Events(EIndex).SubEvents(s)
                .Type = Buffer.ReadLong
                'Textz
                DCount = Buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Text(1 To DCount)
                    .HasText = True
                    For D = 1 To DCount
                        .Text(D) = Buffer.ReadString
                    Next D
                Else
                    Erase .Text
                    .HasText = False
                End If
                'Dataz
                DCount = Buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Data(1 To DCount)
                    .HasData = True
                    For D = 1 To DCount
                        .Data(D) = Buffer.ReadLong
                    Next D
                Else
                    Erase .Data
                    .HasData = False
                End If
            End With
        Next s
    Else
        Events(EIndex).HasSubEvents = False
        Erase Events(EIndex).SubEvents
    End If
    
    Events(EIndex).Trigger = Buffer.ReadByte
    Events(EIndex).WalkThrought = Buffer.ReadByte
    
    Call SaveEvent(EIndex)
    
    Set Buffer = Nothing
End Sub

Public Sub Events_HandleRequestEventData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim EIndex As Long, s As Long, SCount As Long, D As Long, DCount As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    EIndex = Buffer.ReadLong
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Call Events_SendEventData(Index, EIndex)
    
    Set Buffer = Nothing
End Sub

Public Sub Events_HandleRequestEventsData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    For i = 1 To MAX_EVENTS
        Call Events_SendEventData(Index, i)
    Next i
End Sub

Public Sub Events_HandleRequestEditEvents(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventEditor
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Effect packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditEffect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEffectEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Effect packet ::
' ::::::::::::::::::::::
Sub HandleSaveEffect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim EffectSize As Long
    Dim EffectData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_EFFECTS Then
        Exit Sub
    End If

    ' Update the Effect
    EffectSize = LenB(Effect(n))
    ReDim EffectData(EffectSize - 1)
    EffectData = Buffer.ReadBytes(EffectSize)
    CopyMemory ByVal VarPtr(Effect(n)), ByVal VarPtr(EffectData(0)), EffectSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateEffectToAll(n)
    Call SaveEffect(n)
End Sub

Sub HandleRequestEffects(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendEffects Index
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, Target As Long, TargetType As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    Target = Buffer.ReadLong
    TargetType = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' set player's target - no need to send, it's client side
    If TempPlayer(Index).Target <> Target Then
        TempPlayer(Index).Target = Target
        TempPlayer(Index).TargetType = TargetType
        If TargetType = TARGET_TYPE_PLAYER Then SendVital Target, HP, Index
    Else
        TempPlayer(Index).Target = 0
        TempPlayer(Index).TargetType = TARGET_TYPE_NONE
    End If
End Sub

Private Sub HandleEditNews(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Text As String
Dim F As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    Text = Buffer.ReadString
    Set Buffer = Nothing
    
    F = FreeFile
    ' Delete file.
    Kill App.path & "\data\news.txt"
    
    ' Recreate, and then write to it.
    Open App.path & "\data\news.txt" For Append As #F
        Print #F, Trim$(Text)
    Close #F
End Sub

Public Sub HandleRequestEditNews(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
    
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewsEditor
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing
End Sub

Public Sub HandleRequestNews(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNews Index, 0
End Sub
Private Sub HandleDevSuite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim DevSuite As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    DevSuite = Buffer.ReadByte
    Set Buffer = Nothing
    
    If DevSuite = 1 Then
        If GetPlayerAccess(Index) < ADMIN_MONITOR Then
            Call AlertMSG(Index, "Your access level is not enought to access Developer suite")
        End If
    End If
    TempPlayer(Index).inDevSuite = DevSuite
End Sub

Private Sub HandleOnDeath(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(Index).IsDead = 1 Then OnDeath Index
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit quest packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditquest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save quest packet ::
' :::::::::::::::::::::
Private Sub HandleSavequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim questnum As Long
    Dim Buffer As clsBuffer
    Dim questSize As Long
    Dim questData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    questnum = Buffer.ReadLong

    ' Prevent hacking
    If questnum < 0 Or questnum > MAX_QUESTS Then
        Exit Sub
    End If

    questSize = LenB(Quest(questnum))
    ReDim questData(questSize - 1)
    questData = Buffer.ReadBytes(questSize)
    CopyMemory ByVal VarPtr(Quest(questnum)), ByVal VarPtr(questData(0)), questSize
    ' Save it
    Call SendUpdatequestToAll(questnum)
    Call SaveQuest(questnum)
    Call TextAdd(GetPlayerName(Index) & " saved quest #" & questnum & ".", ChatSystem)
End Sub

Public Sub HandleQuestInfo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim questnum As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        questnum = Buffer.ReadLong
        
        If questnum = 0 Then Exit Sub
        
        If Quest(questnum).EventNum = 0 Then
            Call PlayerMsg(Index, printf("Esta quest não tem informações adicionais"), brightred)
            Exit Sub
        End If
        
        If Quest(questnum).EventNum > 0 Then
            InitEvent Index, Quest(questnum).EventNum
        Else
            Call PlayerMsg(Index, printf("Progresso da missão: %d", Val(Player(Index).Variables(-Quest(questnum).EventNum))), Yellow)
        End If
        
    Set Buffer = Nothing
End Sub

Public Sub HandleUpgrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim slot As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        slot = Buffer.ReadLong
        Call UpgradeSpell(Index, slot)
        
    Set Buffer = Nothing
End Sub

Public Sub HandleSellPlanet(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Answer As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        Answer = Buffer.ReadByte
        
        Dim PlanetNum As Long
        PlanetNum = TempPlayer(Index).PlanetCaptured
        
        If PlanetNum > 0 Then
            If Answer = 0 Then
                'SendAnimation GetPlayerMap(Index), 38, Planets(PlanetNum).X, Planets(PlanetNum).Y, GetPlayerDir(Index)
                'SendBossMsg GetPlayerMap(Index), GetPlayerName(Index) & " acabou de explodir o planeta " & Trim$(Planets(PlanetNum).Name) & " (Level " & Planets(PlanetNum).Level & ")", yellow
                Planets(PlanetNum).Owner = GetPlayerName(Index)
                Planets(PlanetNum).State = 2
            Else
                Dim Preco As Long
                Preco = Planets(PlanetNum).Preco
                If Player(Index).EsoBonus > 0 Then Preco = (Preco / 100) * (100 + Player(Index).EsoBonus)
                If Options.GoldFactor > 1 Then Preco = Preco * Options.GoldFactor
                'If MatchData(TempPlayer(Index).MatchIndex).PriceBonus > 0 Then Preco = (Preco / 100) * (100 + MatchData(TempPlayer(Index).MatchIndex).PriceBonus)
                GiveInvItem Index, MoedaZ, Preco, True
                SendBossMsg GetPlayerMap(Index), GetPlayerName(Index) & " vendeu o planeta " & Trim$(Planets(PlanetNum).Name) & " por " & Preco & "$ (Level " & Planets(PlanetNum).Level & ")", Yellow
                SendPlaySound Index, "pagamento conclido.mp3"
                ClearPlanet PlanetNum
                CreatePlanet PlanetNum
                CreatePlanetMap PlanetNum
                MapCache_Create Planets(PlanetNum).Map
            End If
            
            TempPlayer(Index).PlanetCaptured = 0
            SendPlanetToAll PlanetNum
        End If
        
    Set Buffer = Nothing
End Sub
Public Sub HandleEnterGravity(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Gravity As Long, Hours As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        Gravity = Buffer.ReadLong
        Hours = Buffer.ReadLong
        
        If Gravity < 10 Then Gravity = 10
        If Gravity > 1500 Then Gravity = 1500
        If Hours < 1 Then Hours = 1
        If Hours > 6 Then Hours = 6
        
        Dim Price As Long
        Price = GravityValue(Index, Gravity) * Hours
        
        If HasItem(Index, MoedaZ) >= Price Then
            TakeInvItem Index, MoedaZ, Price
            Player(Index).GravityHours = Hours
            Player(Index).GravityInit = Now
            Player(Index).GravityValue = Gravity
            PlayerWarp Index, GravityMap, 14, 14
            SendDialogue Index, "Aviso", "Bem-vindo á sala de gravidade, você permanecerá aqui pelas próximas horas e ao fim receberá experiência pelo seu treinamento. Durante o treinamento VOCÊ NÃO PRECISA PERMANECER ONLINE, por isso é recomendado utilizar este treinamento apenas quando você for ficar ausente."
            SavePlayer Index
        Else
            PlayerMsg Index, "Você não tem o suficiente para entrar na sala de gravidade nestas configurações (Valor necessário: " & Price & "z)", brightred
        End If
        
    Set Buffer = Nothing
End Sub

Public Sub HandleCompleteTutorial(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(Index).InTutorial = 0 Then
        Player(Index).InTutorial = 1
        SendPlayerData Index
        SavePlayer Index
    End If
End Sub

Public Sub HandleFeedback(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Tipo As Long, Text As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        Tipo = Buffer.ReadLong
        Text = Buffer.ReadString
        
        addFeedback Index, Tipo, Text
        
    Set Buffer = Nothing
End Sub

Public Sub HandleCreateGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        Dim i As Long
        
        Name = Buffer.ReadString
        For i = 1 To MAX_GUILDS
            If Trim$(LCase(Guild(i).Name)) = Trim$(LCase(Name)) Then
                PlayerMsg Index, "O nome desta guild já está em uso!", brightred
                Exit Sub
            End If
        Next i
        
        If Player(Index).Guild > 0 Then
            PlayerMsg Index, "Você já está em um guild!", brightred
            Exit Sub
        End If
        
        If GetPlayerLevel(Index) < 20 Then
            PlayerMsg Index, "É necessário ser ao menos level 20 para criar uma guild!", Yellow
            Exit Sub
        End If
        
        If HasItem(Index, MoedaZ) < 1000000 Then
            PlayerMsg Index, "É necessário ter 1 milhão de Moedas Z para criar uma guild!", brightred
            Exit Sub
        End If
        
        i = FindOpenGuildSlot
        
        If i >= 1 And i <= MAX_GUILDS Then
            TakeInvItem Index, MoedaZ, 1000000
            Guild(i).Name = Name
            Guild(i).Level = 1
            Guild(i).TNL = Experience(1)
            Dim n As Long
            For n = 1 To 25
                Guild(i).IconColor(n) = Buffer.ReadByte
            Next n
            GlobalMsg "A guild " & Trim$(Guild(i).Name) & " acabou de ser criada pelo seu mestre " & GetPlayerName(Index), Yellow
            AddMember i, Index, GuildRank.Mestre
        Else
            PlayerMsg Index, "Sem espaço para novas guilds!", brightred
        End If
        
    Set Buffer = Nothing
End Sub
Public Sub HandleGuildAction(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Action As Byte
    Dim Buffer As clsBuffer
    Dim AccessRequired As Byte
    Dim GuildNum As Long
    Dim mString As String, mLong As Long, mByte As Byte
    Dim PlayerIndex As Long
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        GuildNum = Player(Index).Guild
        
        Action = Buffer.ReadByte
        
        If GuildNum = 0 Then
            If Action = 3 Then
                GuildNum = TempPlayer(Index).GuildInvite
                If GuildNum > 0 Then
                    Dim Answer As Byte
                    Answer = Buffer.ReadByte
                    
                    If Answer = 1 Then
                        If Not IsGuildFull(GuildNum) Then
                            AddMember GuildNum, Index, GuildRank.Member
                            GuildMsg GuildNum, GetPlayerName(Index) & " entrou na guild!", Yellow
                            Exit Sub
                        Else
                            PlayerMsg Index, "A guild está atualmente cheia!", brightred
                            Exit Sub
                        End If
                    Else
                        PlayerIndex = TempPlayer(Index).GuildInviteIndex
                        TempPlayer(Index).GuildInviteIndex = 0
                        TempPlayer(Index).GuildInvite = 0
                        If IsPlaying(PlayerIndex) Then
                            PlayerMsg PlayerIndex, GetPlayerName(Index) & " recusou seu convite para ingressar a guild!", brightred
                            Exit Sub
                        End If
                    End If
                End If
            Else
                Exit Sub
            End If
        End If
        
        Select Case Action
            Case 1: AccessRequired = GuildRank.Mestre 'Trocar MOTD
            Case 2: AccessRequired = GuildRank.Capitao 'Convidar
            Case 4: AccessRequired = GuildRank.Major 'Rebaixar
            Case 5: AccessRequired = GuildRank.Major 'Promover
            Case 6: AccessRequired = GuildRank.Major 'Expulsar
            Case 9: AccessRequired = GuildRank.Mestre  'UpBlock
        End Select
        
        If GetPlayerGuildRank(Index) < AccessRequired Then
            PlayerMsg Index, "Você não tem autorização da guild para executar este comando! (Acesso requerido: " & RankName(AccessRequired) & ")", brightred
            Exit Sub
        End If
        
        Select Case Action
            Case 1 'Trocar MOTD
                Guild(GuildNum).MOTD = Buffer.ReadString
                GuildMsg GuildNum, "Guild MOTD: " & Trim$(Guild(GuildNum).MOTD), Yellow
            Case 2 'Convidar
                mString = Buffer.ReadString
                PlayerIndex = FindPlayer(mString)
                If Not IsGuildFull(GuildNum) Then
                    If PlayerIndex > 0 Then
                        If Player(PlayerIndex).Guild = 0 Then
                            TempPlayer(PlayerIndex).GuildInvite = GuildNum
                            TempPlayer(PlayerIndex).GuildInviteIndex = Index
                            'PlayerMsg PlayerIndex, "O jogador " & GetPlayerName(Index) & " está te convidando para a guild " & Trim$(Guild(GuildNum).name) & ", para aceitar digite /aceitar ou /recusar para recusar", Yellow
                            SendGuildInvite PlayerIndex, "O jogador " & GetPlayerName(Index) & " está te convidando para a guild " & Trim$(Guild(GuildNum).Name) & "! Você deseja entrar na guild?"
                            PlayerMsg Index, "Convite enviado para " & mString & "!", brightgreen
                        Else
                            PlayerMsg Index, "Este jogador já está em uma guild!", brightred
                        End If
                    Else
                        PlayerMsg Index, "Jogador não encontrado!", brightred
                    End If
                Else
                    PlayerMsg Index, "Sua guild está cheia!", brightred
                End If
            Case 4 'Rebaixar
                mLong = Buffer.ReadLong
                If mLong > 0 And mLong <= 10 Then
                    If Trim$(Guild(GuildNum).Member(mLong).Name) = GetPlayerName(Index) Then
                        Exit Sub
                    End If
                    If Guild(GuildNum).Member(mLong).Rank > 0 Then
                        If Guild(GuildNum).Member(mLong).Rank < Guild(GuildNum).Member(GetPlayerGuildIndex(Index)).Rank Then
                            Guild(GuildNum).Member(mLong).Rank = Guild(GuildNum).Member(mLong).Rank - 1
                            GuildMsg GuildNum, GetPlayerName(Index) & " rebaixou " & Trim$(Guild(GuildNum).Member(mLong).Name) & " para o grau de " & RankName(Guild(GuildNum).Member(mLong).Rank), Yellow
                            SaveGuild GuildNum
                            SendUpdateGuildToAll GuildNum
                        Else
                            PlayerMsg Index, "Você não pode rebaixar um jogador com o grau maior que o seu!", brightred
                        End If
                    Else
                        PlayerMsg Index, "Este jogador já está no grau mínimo", brightred
                    End If
                End If
            Case 5 'Promover
                mLong = Buffer.ReadLong
                If mLong > 0 And mLong <= 10 Then
                    If Trim$(Guild(GuildNum).Member(mLong).Name) = GetPlayerName(Index) Then
                        Exit Sub
                    End If
                    If Guild(GuildNum).Member(mLong).Rank < GuildRank.Mestre Then
                        If Guild(GuildNum).Member(mLong).Rank + 1 < Guild(GuildNum).Member(GetPlayerGuildIndex(Index)).Rank Or Guild(GuildNum).Member(mLong).Rank + 1 = GuildRank.Mestre Then
                            Guild(GuildNum).Member(mLong).Rank = Guild(GuildNum).Member(mLong).Rank + 1
                            If Guild(GuildNum).Member(mLong).Rank = GuildRank.Mestre Then
                                Guild(GuildNum).Member(GetPlayerGuildIndex(Index)).Rank = Guild(GuildNum).Member(GetPlayerGuildIndex(Index)).Rank - 1
                                GuildMsg GuildNum, GetPlayerName(Index) & " foi rebaixado para o grau de " & RankName(Guild(GuildNum).Member(GetPlayerGuildIndex(Index)).Rank), Yellow
                            End If
                            GuildMsg GuildNum, GetPlayerName(Index) & " promoveu " & Trim$(Guild(GuildNum).Member(mLong).Name) & " para o grau de " & RankName(Guild(GuildNum).Member(mLong).Rank), Yellow
                            SaveGuild GuildNum
                            SendUpdateGuildToAll GuildNum
                        Else
                            PlayerMsg Index, "Você não pode promover um jogador para um grau maior que o seu!", brightred
                        End If
                    Else
                        PlayerMsg Index, "Este jogador já está no grau máximo", brightred
                    End If
                End If
            Case 6 'Expulsar
                mLong = Buffer.ReadLong
                If mLong > 0 And mLong <= 10 Then
                    If Trim$(Guild(GuildNum).Member(mLong).Name) = GetPlayerName(Index) Then
                        Exit Sub
                    End If
                    If Guild(GuildNum).Member(mLong).Rank < Guild(GuildNum).Member(GetPlayerGuildIndex(Index)).Rank Then
                        GuildMsg GuildNum, Trim(Guild(GuildNum).Member(mLong).Name) & " foi expulso da guild por " & GetPlayerName(Index), brightred
                        Guild(GuildNum).Member(mLong).Name = vbNullString
                        Guild(GuildNum).Member(mLong).Level = 0
                        'Está online?
                        PlayerIndex = FindPlayer(Trim(Guild(GuildNum).Member(mLong).Name))
                        If PlayerIndex > 0 Then
                            Player(PlayerIndex).Guild = 0
                            SendPlayerData PlayerIndex
                        End If
                        SaveGuild GuildNum
                        SendUpdateGuildToAll GuildNum
                    Else
                        PlayerMsg Index, "Você não pode expulsar um membro com o grau igual ou maior que o seu!", brightred
                    End If
                End If
            Case 7 'Sair da guild
                mLong = GetPlayerGuildIndex(Index)
                If mLong > 0 Then
                    If Guild(GuildNum).Member(mLong).Rank = GuildRank.Mestre Then
                        If Not IsGuildEmpty(GuildNum) Then
                            PlayerMsg Index, "Você não pode sair da guild sendo o mestre dela!", brightred
                            Exit Sub
                        End If
                    End If
                    GuildMsg GuildNum, GetPlayerName(Index) & " saiu da guild!", brightred
                    Player(Index).Guild = 0
                    Guild(GuildNum).Member(mLong).Name = vbNullString
                    Guild(GuildNum).Member(mLong).Level = 0
                    SendPlayerData Index
                    SaveGuild GuildNum
                    SendUpdateGuildToAll GuildNum
                End If
            Case 8 'Doação
                mByte = Buffer.ReadByte 'Produto
                mLong = Buffer.ReadLong 'Quantidade
                
                Dim ItemNum As Long
                If mByte = 0 Then ItemNum = EspV
                If mByte = 1 Then ItemNum = EspAz
                If mByte = 2 Then ItemNum = EspAm
                If mByte = 3 Then ItemNum = 26
                
                If HasItem(Index, ItemNum) >= mLong Then
                    TakeInvItem Index, ItemNum, mLong
                    If mByte = 0 Then Guild(GuildNum).Red = Guild(GuildNum).Red + mLong
                    If mByte = 1 Then Guild(GuildNum).Blue = Guild(GuildNum).Blue + mLong
                    If mByte = 2 Then Guild(GuildNum).Yellow = Guild(GuildNum).Yellow + mLong
                    If mByte = 3 Then Guild(GuildNum).Gold = Guild(GuildNum).Gold + mLong
                    SaveGuild GuildNum
                    SendUpdateGuildToAll GuildNum
                    GuildMsg GuildNum, GetPlayerName(Index) & " doou " & mLong & " " & Trim$(Item(ItemNum).Name) & " para a guild!", brightgreen
                Else
                    PlayerMsg Index, "Você não tem essa quantidade para doar!", brightred
                End If
            Case 9 'UpBlock
                mByte = Buffer.ReadByte
                Guild(GuildNum).UpBlock = mByte
                If mByte = 1 Then
                    GuildMsg GuildNum, "O mestre bloqueou a evolução da guild", Yellow
                Else
                    GuildMsg GuildNum, "O mestre desbloqueou a evolução da guild", Yellow
                End If
                SaveGuild GuildNum
        End Select
        
    Set Buffer = Nothing
End Sub
Public Sub HandleChallengeArena(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Action As Byte
    Dim Buffer As clsBuffer
    Dim MatchType As Byte
    Dim MatchPlayers As Byte
    Dim TotalPlayer As Byte
    Dim PlayerIndex As Long
    Dim PlayerName(0 To 4) As String
    Dim Aposta As Long
    Dim i As Long
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        Action = Buffer.ReadByte
        
        If Action = 1 Then 'Challenge
            If ArenaChallenge.Active > 0 Then
                PlayerMsg Index, "A arena está ocupada no momento, aguarde alguns instantes e tente novamente!", brightred
                Set Buffer = Nothing
                Exit Sub
            End If
            
            MatchType = Buffer.ReadByte
            ArenaChallenge.MatchType = MatchType
            MatchPlayers = Buffer.ReadByte
            Select Case MatchPlayers
                Case 0: TotalPlayer = 2
                Case 1: TotalPlayer = 4
                Case 2: TotalPlayer = 6
            End Select
            ArenaChallenge.TotalPlayers = TotalPlayer
            
            Dim PlayerCount As Long
            ArenaChallenge.Players(1) = GetPlayerName(Index)
            ArenaChallenge.PlayerAccept(1) = True
            PlayerCount = 1
            For i = 0 To 4
                PlayerName(i) = Buffer.ReadString
                ArenaChallenge.Players(i + 2) = Trim$(PlayerName(i))
                
                If PlayerCount < TotalPlayer Then
                    PlayerIndex = FindPlayer(Trim$(PlayerName(i)))
                    If PlayerIndex = 0 Then
                        PlayerMsg Index, "O jogador " & Trim$(PlayerName(i)) & " não está online!", brightred
                        Set Buffer = Nothing
                        Exit Sub
                    Else
                        If Player(PlayerIndex).GravityHours > 0 Then
                            PlayerMsg Index, "O jogador " & Trim$(PlayerName(i)) & " está na gravidade e não pode receber desafios!", brightred
                            Set Buffer = Nothing
                            Exit Sub
                        End If
                        PlayerCount = PlayerCount + 1
                    End If
                End If
            Next i
            Aposta = Buffer.ReadLong
            ArenaChallenge.Aposta = Aposta + 10000
            ArenaChallenge.Active = 1
            
            'Tudo certo, enviar pedidos
            Dim Msg As String
            If TotalPlayer > 2 Then
                Msg = "Os jogadores: "
            Else
                Msg = "O jogador: "
            End If
            For i = 2 To TotalPlayer
                ArenaChallenge.PlayerAccept(i) = False
                PlayerIndex = FindPlayer(Trim$(ArenaChallenge.Players(i)))
                If i < TotalPlayer Then
                    Msg = Msg & Trim$(ArenaChallenge.Players(i)) & ", "
                Else
                    Msg = Msg & Trim$(ArenaChallenge.Players(i)) & " "
                End If
                If PlayerIndex > 0 Then
                    SendArenaChallenge PlayerIndex, GetPlayerName(Index) & " está de desafiando para um combate na arena, isto custará 10,000z de taxa da arena e mais " & Aposta & "z que ele está propondo como aposta por jogador, você deseja aceitar?"
                End If
            Next i
            If TotalPlayer > 2 Then
                Msg = Msg & "foram desafiados a se enfrentar na arena por " & GetPlayerName(Index) & "!"
            Else
                Msg = Msg & "foi desafiado a enfrentar " & GetPlayerName(Index) & " na arena!"
            End If
            GlobalMsg Msg, Yellow
            ArenaChallenge.LastCall = GetTickCount
            Exit Sub
        End If
        
        If Action = 2 Then 'Accept Refuse
            If ArenaChallenge.Active <> 1 Then
                PlayerMsg Index, "Este convite de arena foi cancelado!", brightred
                Set Buffer = Nothing
                Exit Sub
            End If
            
            If Buffer.ReadByte = 1 Then 'Accept
                For i = 2 To ArenaChallenge.TotalPlayers
                    If GetPlayerName(Index) = ArenaChallenge.Players(i) Then
                        ArenaChallenge.PlayerAccept(i) = True
                        CheckArenaStart
                        Exit For
                    End If
                Next i
                GlobalMsg GetPlayerName(Index) & " aceitou o desafio da arena!", brightgreen
            Else
                ArenaChallenge.Active = 0
                GlobalMsg GetPlayerName(Index) & " recusou o pedido de arena, desafio cancelado.", brightred
            End If
        End If
        
    Set Buffer = Nothing
End Sub

Public Sub HandleAntiHackData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim DLLs As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        DLLs = Buffer.ReadString
        addAntiHackLog GetPlayerName(Index) & " " & DLLs
        
    Set Buffer = Nothing
End Sub

Public Sub HandlePlanetChange(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Action As Byte
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim MapNum As Long
    Dim Buffer As clsBuffer
    Dim sString As String
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        Action = Buffer.ReadByte
        
        Dim PlanetNum As Long
        MapNum = GetPlayerMap(Index)
        PlanetNum = PlayerMapIndex(MapNum)
        
        If PlanetNum = 0 Then Exit Sub
        If Trim$(LCase(PlayerPlanet(PlanetNum).PlanetData.Owner)) <> Trim$(LCase(GetPlayerName(Index))) Then Exit Sub
        
        If Action = 0 Then 'Remove Block
            Dim Price As Long
            X = Buffer.ReadLong
            Y = Buffer.ReadLong
            If Map(MapNum).Tile(X, Y).Type = TileType.tile_type_blocked Then Price = 10000
            If Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_NPCSPAWN Then Price = 30000
            If Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_RESOURCE Then Price = 50000
            If HasItem(Index, MoedaZ) >= Price Then
                If Price = 0 Then
                    For i = 1 To PlayerPlanet(PlanetNum).TotalSaibamans
                        If PlayerPlanet(PlanetNum).Saibaman(i).Working = 1 Then
                            If PlayerPlanet(PlanetNum).Saibaman(i).X = X And PlayerPlanet(PlanetNum).Saibaman(i).Y = Y Then
                                PlayerPlanet(PlanetNum).Saibaman(i).Working = 0
                                SendSaibamans Index, PlanetNum
                                SavePlayerPlanet PlanetNum
                            End If
                        End If
                    Next i
                    Exit Sub
                End If
                TakeInvItem Index, MoedaZ, Price
                Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE
                Map(MapNum).Tile(X, Y).Layer(2).Tileset = 0
                PlayerPlanet(PlanetNum).PlanetMap.Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE
                PlayerPlanet(PlanetNum).PlanetMap.Tile(X, Y).Layer(2).Tileset = 0
                SavePlayerPlanet PlanetNum
                MapCache_Create MapNum
                SendMap Index, MapNum
                If Price = 50000 Then
                    CacheResources MapNum
                    SendResourceCacheTo Index
                End If
                If Price = 30000 Then
                    ' Respawn NPCS
                    Map(MapNum).Npc(Map(MapNum).Tile(X, Y).data1) = 0
                    PlayerPlanet(PlanetNum).PlanetMap.Npc(PlayerPlanet(PlanetNum).PlanetMap.Tile(X, Y).data1) = 0
                    For i = 1 To MAX_MAP_NPCS
                        Call SpawnNpc(i, MapNum)
                    Next
                    SendMapNpcsToMap MapNum
                End If
                PlayerMsg Index, "Bloco removido com sucesso!", brightgreen
            Else
                PlayerMsg Index, "Você não tem dinheiro suficiente para remover este bloco!", brightred
            End If
        End If
        
        If Action = 1 Then 'Name planet
            sString = Buffer.ReadString
            If Len(sString) > 3 And Len(sString) < 12 Then
                Dim HaveItem As Boolean
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(Index, i) > 0 Then
                        If Item(GetPlayerInvItemNum(Index, i)).Type = ItemType.ITEM_TYPE_PLANETCHANGE And Item(GetPlayerInvItemNum(Index, i)).data1 = 0 Then
                            TakeInvItem Index, GetPlayerInvItemNum(Index, i), 1
                            HaveItem = True
                            Exit For
                        End If
                    End If
                Next i
                If HaveItem Then
                    sString = UCase(sString)
                    PlayerPlanet(PlanetNum).PlanetData.Name = sString
                    Map(MapNum).Name = sString
                    SavePlayerPlanet PlanetNum
                    MapCache_Create MapNum
                    SendMap Index, MapNum
                    PlayerMsg Index, "Nome alterado com sucesso!", brightgreen
                Else
                    PlayerMsg Index, "Você precisa ter um item nomeador de planetas para dar um nome!", brightred
                End If
            Else
                PlayerMsg Index, "O nome do seu planeta tem que estar entre 4 e 11 caracteres", brightred
            End If
        End If
        
        If Action = 2 Then 'Acelerar
            Dim SaibamanIndex As Long
            X = Buffer.ReadLong
            Y = Buffer.ReadLong
            For i = 1 To PlayerPlanet(PlanetNum).TotalSaibamans
                If PlayerPlanet(PlanetNum).Saibaman(i).Working = 1 Then
                    If PlayerPlanet(PlanetNum).Saibaman(i).X = X And PlayerPlanet(PlanetNum).Saibaman(i).Y = Y Then
                        SaibamanIndex = i
                        Exit For
                    End If
                End If
            Next i
            If SaibamanIndex > 0 Then
                Dim MoedasAcc As Long, ItemNum As Long
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(Index, i) = Relogio Then
                        MoedasAcc = GetPlayerInvItemValue(Index, i)
                        ItemNum = GetPlayerInvItemNum(Index, i)
                        HaveItem = True
                        Exit For
                    End If
                Next i
                If HaveItem Then
                    'PlayerPlanet(PlanetNum).Saibaman(SaibamanIndex).Accelerate = PlayerPlanet(PlanetNum).Saibaman(SaibamanIndex).Accelerate + 1
                    Dim QuantNeed As Long
                    Dim Minutes As Long
                    If PlayerPlanet(PlanetNum).Saibaman(SaibamanIndex).TaskType = 0 Then
                        Minutes = Npc(PlayerPlanet(PlanetNum).Saibaman(SaibamanIndex).TaskResult).TimeToEvolute
                    End If
                    If PlayerPlanet(PlanetNum).Saibaman(SaibamanIndex).TaskType = 1 Then
                        Minutes = Resource(PlayerPlanet(PlanetNum).Saibaman(SaibamanIndex).TaskResult).TimeToEvolute
                    End If
                    Minutes = Minutes - (DateDiff("n", PlayerPlanet(PlanetNum).Saibaman(SaibamanIndex).TaskInit, Now))
                    QuantNeed = (100 - (5 * Int(Minutes / 100)))
                    If QuantNeed < 50 Then QuantNeed = 50
                    QuantNeed = (Minutes / 100) * QuantNeed
                    If MoedasAcc >= QuantNeed Then
                        TakeInvItem Index, ItemNum, QuantNeed
                        PlayerPlanet(PlanetNum).Saibaman(SaibamanIndex).Accelerate = 1
                        UpdateSaibaman PlanetNum, SaibamanIndex
                        PlayerMsg Index, "Saibaman acelerado com sucesso!", brightgreen
                        SendSaibamans Index, PlanetNum
                    Else
                        PlayerMsg Index, "Você não tem aceleradores de partículas suficientes!", brightred
                    End If
                Else
                    PlayerMsg Index, "Você não tem aceleradores de partículas!", brightred
                End If
            Else
                PlayerMsg Index, "Não foi encontrado um saibaman neste local, tente novamente", brightred
            End If
        End If
        
        If Action = 3 Then ' Mover
            X = Buffer.ReadLong
            Y = Buffer.ReadLong
            
            Dim NewX As Long, NewY As Long
            NewX = Buffer.ReadLong
            NewY = Buffer.ReadLong
            
            If Map(MapNum).Tile(NewX, NewY).Type <> TileType.TILE_TYPE_WALKABLE Then Exit Sub
            If Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE Then Exit Sub
            
            Map(MapNum).Tile(NewX, NewY) = Map(MapNum).Tile(X, Y)
            
            Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_WALKABLE
            If Map(MapNum).Tile(NewX, NewY).Type = TileType.TILE_TYPE_NPCSPAWN Then
                ' Respawn NPCS
                For i = 1 To MAX_MAP_NPCS
                    Call SpawnNpc(i, MapNum)
                Next
                SendMapNpcsToMap MapNum
            Else
                If Resource(Map(MapNum).Tile(NewX, NewY).data1).ItemReward > 0 Then SwitchExtrator PlanetNum, X, Y, NewX, NewY
                CacheResources MapNum
            End If
            PlayerPlanet(PlanetNum).PlanetMap = Map(MapNum)
            SavePlayerPlanet PlanetNum
            MapCache_Create MapNum
            SendMap Index, MapNum
            SendResourceCacheToMap MapNum, , True
        End If
        
        If Action = 4 Then
            X = Buffer.ReadLong
            Y = Buffer.ReadLong
            
            If Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_RESOURCE Then
                Dim ResourceNum As Long
                ResourceNum = Map(MapNum).Tile(X, Y).data1
                If Resource(ResourceNum).Evolution > 0 Then
                    Dim EvNum As Long
                    EvNum = Resource(ResourceNum).Evolution
                    StartConstructResource Index, EvNum, MapNum, PlanetNum, X, Y, True
                End If
            End If
            
            If Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_NPCSPAWN Then
                Dim NpcNum As Long
                NpcNum = Map(MapNum).Npc(Map(MapNum).Tile(X, Y).data1)
                If Npc(NpcNum).Evolution > 0 Then
                    EvNum = Npc(NpcNum).Evolution
                    StartConstructNPC Index, MapNum, PlanetNum, X, Y, EvNum, True
                End If
            End If
        End If
        
        If Action = 5 Then
            X = Buffer.ReadLong
            Y = Buffer.ReadLong
            
            If Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_RESOURCE Then
                ResourceNum = Map(MapNum).Tile(X, Y).data1
                If Resource(ResourceNum).ToolRequired = 1 Then 'Fábrica
                    SendFabrica Index, PlanetNum
                    Exit Sub
                End If
                If Resource(ResourceNum).ToolRequired = 2 Then 'Exército
                    SendSoldados Index, PlanetNum
                    Exit Sub
                End If
            End If
        End If
        
        If Action = 6 Then
            X = Buffer.ReadLong
            Y = Buffer.ReadLong
            
            If Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_RESOURCE Then
                ResourceNum = Map(MapNum).Tile(X, Y).data1
                If Resource(ResourceNum).ToolRequired = 1 Then 'Fábrica
                    Dim Nivel As Byte, RedCost As Long, BlueCost As Long, YellowCost As Long, Time As Long, Quant As Long
                    Nivel = Buffer.ReadByte
                    If Nivel < 1 Or Nivel > 5 Then Exit Sub
                    Quant = Buffer.ReadByte
                    RedCost = (Fat(Nivel) * 25) * Quant
                    BlueCost = (Fat(Nivel) * 15) * Quant
                    YellowCost = (Fat(Nivel) * 5) * Quant
                    If Nivel > Resource(ResourceNum).ResourceLevel Then
                        PlayerMsg Index, "Você precisa de uma fábrica nível " & Nivel & " para fabricar estas sementes", brightred
                        Exit Sub
                    End If
                    If PlayerPlanet(PlanetNum).Sementes(Nivel).Fila + Quant > 50 + (25 * NucleoLevel(PlanetNum)) Then
                        PlayerMsg Index, "O limite máximo da fila de sementes é " & (50 + (25 * NucleoLevel(PlanetNum))) & " evolua seu centro para expandir a fila", brightred
                        Exit Sub
                    End If
                    If HasItem(Index, EspV) < RedCost Then
                        PlayerMsg Index, "Você não tem especiaria vermelha suficiente!", brightred
                        Exit Sub
                    End If
                    If HasItem(Index, EspAz) < BlueCost Then
                        PlayerMsg Index, "Você não tem especiaria azul suficiente!", brightred
                        Exit Sub
                    End If
                    If HasItem(Index, EspAm) < YellowCost Then
                        PlayerMsg Index, "Você não tem especiaria amarela suficiente!", brightred
                        Exit Sub
                    End If
                    TakeInvItem Index, EspV, RedCost
                    TakeInvItem Index, EspAz, BlueCost
                    TakeInvItem Index, EspAm, YellowCost
                    PlayerPlanet(PlanetNum).Sementes(Nivel).Fila = PlayerPlanet(PlanetNum).Sementes(Nivel).Fila + Quant
                    PlayerPlanet(PlanetNum).Sementes(Nivel).Start = Now
                    PlayerMsg Index, "Sementes adicionadas á fila com sucesso!", brightgreen
                    SavePlayerPlanet PlanetNum
                    SendFabrica Index, PlanetNum
                End If
                If Resource(ResourceNum).ToolRequired = 2 Then 'Saibamans
                    Dim GoldCost As Long
                    Nivel = Buffer.ReadByte
                    If Nivel < 1 Or Nivel > 5 Then Exit Sub
                    Quant = Buffer.ReadByte
                    GoldCost = (Quant * (Fat(Nivel + 2) * 100))
                    If Nivel > Resource(ResourceNum).ResourceLevel Then
                        PlayerMsg Index, "Você precisa de um exército nível " & Nivel & " para fabricar estes saibamans", brightred
                        Exit Sub
                    End If
                    If Allocated(PlanetNum) + Quant > Alloc(PlanetNum) Then
                        PlayerMsg Index, "Você não pode criar esta quantidade de saibamans pois excede seu numero de residencias! Crie mais casas ou as evolua!", brightred
                        Exit Sub
                    End If
                    If PlayerPlanet(PlanetNum).Soldados(Nivel).Fila + Quant > 50 + (25 * NucleoLevel(PlanetNum)) Then
                        PlayerMsg Index, "O limite máximo da fila de saibamans é " & (50 + (25 * NucleoLevel(PlanetNum))) & " evolua seu centro para expandir a fila", brightred
                        Exit Sub
                    End If
                    If HasItem(Index, MoedaZ) < GoldCost Then
                        PlayerMsg Index, "Você não tem moedas z suficiente!", brightred
                        Exit Sub
                    End If
                    TakeInvItem Index, MoedaZ, GoldCost
                    PlayerPlanet(PlanetNum).Soldados(Nivel).Fila = PlayerPlanet(PlanetNum).Soldados(Nivel).Fila + Quant
                    PlayerPlanet(PlanetNum).Soldados(Nivel).Start = Now
                    PlayerMsg Index, "Saibamans adicionadas á fila com sucesso!", brightgreen
                    SavePlayerPlanet PlanetNum
                    SendSoldados Index, PlanetNum
                End If
            End If
        End If
        
        If Action = 7 Then 'Acelerar
            X = Buffer.ReadLong
            Y = Buffer.ReadLong
            
            If Map(MapNum).Tile(X, Y).Type = TileType.TILE_TYPE_RESOURCE Then
                ResourceNum = Map(MapNum).Tile(X, Y).data1
                If Resource(ResourceNum).ItemReward > 0 Then
                    Dim ExtratorIndex As Long
                    ExtratorIndex = GetExtratorIndex(PlanetNum, X, Y)
                    If ExtratorIndex > 0 Then
                        If PlayerPlanet(PlanetNum).Extrator(ExtratorIndex).Acc = 0 Then
                            If HasItem(Index, Relogio) >= 100 Then
                                TakeInvItem Index, Relogio, 100
                                PlayerPlanet(PlanetNum).Extrator(ExtratorIndex).Acc = 1
                                PlayerPlanet(PlanetNum).Extrator(ExtratorIndex).AccStart = Now
                                PlayerMsg Index, "Extrator acelerado!", brightgreen
                                SavePlayerPlanet PlanetNum
                            Else
                                PlayerMsg Index, "Você não tem aceleradores de partículas suficientes para acelerar este extrator!", brightred
                            End If
                        Else
                            PlayerMsg Index, "Este extrator já está acelerado!", brightred
                        End If
                    End If
                End If
                If Resource(ResourceNum).ToolRequired = 1 Then
                    If HasItem(Index, Relogio) >= 50 Then
                        TakeInvItem Index, Relogio, 50
                        PlayerPlanet(PlanetNum).SementesAcc = 1
                        PlayerPlanet(PlanetNum).SementesStart = Now
                        PlayerMsg Index, "Fábricas aceleradas! As sementes serão produzidas na metade do tempo!", brightgreen
                        SavePlayerPlanet PlanetNum
                    Else
                        PlayerMsg Index, "Você não tem aceleradores de partículas suficientes para acelerar esta produção!", brightred
                    End If
                End If
                If Resource(ResourceNum).ToolRequired = 2 Then
                    If HasItem(Index, Relogio) >= 50 Then
                        TakeInvItem Index, Relogio, 50
                        PlayerPlanet(PlanetNum).SoldadosAcc = 1
                        PlayerPlanet(PlanetNum).SoldadosStart = Now
                        PlayerMsg Index, "Exércitos acelerados! Os saibamans serão produzidos na metade do tempo!", brightgreen
                        SavePlayerPlanet PlanetNum
                    Else
                        PlayerMsg Index, "Você não tem aceleradores de partículas suficientes para acelerar esta produção!", brightred
                    End If
                End If
            End If
        End If
        
    Set Buffer = Nothing
End Sub

Public Sub HandleConfirmation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim YesNo As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        YesNo = Buffer.ReadByte
        
        If YesNo = 1 Then
            Select Case TempPlayer(Index).Confirmation
            
                Case ConfirmType.NewPlanet
                    CapturePlanet Index, TempPlayer(Index).ConfirmationVar, True
                
                Case ConfirmType.DestroyItem
                    TakeInvItem Index, GetPlayerInvItemNum(Index, TempPlayer(Index).ConfirmationVar), GetPlayerInvItemValue(Index, TempPlayer(Index).ConfirmationVar)
                
            End Select
        Else
            TempPlayer(Index).Confirmation = 0
            TempPlayer(Index).ConfirmationVar = 0
        End If
        
    Set Buffer = Nothing
End Sub

Public Sub HandleSellEspeciaria(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Number As Byte
    Dim Quant As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        Number = Buffer.ReadByte
        Quant = Buffer.ReadLong
        
        If Number < 1 Or Number > 3 Then Exit Sub
        If Quant < 0 Or Quant > 2000000000 Then Exit Sub
        
        Dim ItemNumber As Long
        ItemNumber = 80 + (Number - 1)
        
        If HasItem(Index, ItemNumber) >= Quant Then
            TakeInvItem Index, ItemNumber, Quant
            GiveInvItem Index, MoedaZ, GetEspeciariaPrice(Number) * Quant
            EspAmount(Number) = EspAmount(Number) + Quant
            SendOpenTroca Index
            PlayerMsg Index, "Você recebeu " & (GetEspeciariaPrice(Number) * Quant) & "z pela venda de " & Quant & " especiarias no preço de " & GetEspeciariaPrice(Number) & "z cada.", Yellow
            SendPlaySound Index, "pagamento conclido.mp3"
        Else
            PlayerMsg Index, "Você não tem essa quantidade de especiaria para vender!", brightred
        End If
        
    Set Buffer = Nothing
End Sub

Public Sub HandleSupport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim MsgTo As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        Msg = Buffer.ReadString
        MsgTo = Buffer.ReadString
        
        Msg = GetPlayerName(Index) & ": " & Msg
        
        If Msg = GetPlayerName(Index) & ": open" And GetPlayerAccess(Index) > 0 Then
            GlobalMsg GetPlayerName(Index) & "(ADM) está disponível para suporte! Digite /suporte caso precise de ajuda.", White
            Exit Sub
        End If
        
        If MsgTo = "admin" Then
            Dim i As Long
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If GetPlayerAccess(i) > 0 Then
                        If Msg = GetPlayerName(Index) & ": open" Then
                            SendSupportMsg Index, GetPlayerName(i) & ": Chat iniciado"
                            SendSupportMsg i, GetPlayerName(Index) & ": Chat iniciado"
                        Else
                            SendSupportMsg i, Msg
                        End If
                        Exit Sub
                    End If
                End If
            Next i
            If Msg <> GetPlayerName(Index) & ": open" Then
                SendSupportMsg Index, "Servidor: Nenhum administrador se encontra online no momento, tente novamente mais tarde."
            Else
                PlayerMsg Index, "Nenhum administrador se encontra online no momento, tente novamente mais tarde.", brightred
            End If
        Else
            i = FindPlayer(MsgTo)
            If i > 0 Then
                SendSupportMsg i, Msg

            Else
                SendSupportMsg Index, "Jogador offline"
            End If
        End If
        
    Set Buffer = Nothing
End Sub

