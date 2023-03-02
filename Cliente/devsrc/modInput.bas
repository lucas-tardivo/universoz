Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub CheckKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If
    
    If GetKeyState(vbKeyTab) < 0 Then
        tabDown = True
    Else
        tabDown = False
    End If
    
    If Not chatOn Then
        If GetKeyState(vbKeySpace) < 0 Then
            CheckMapGetItem
        End If
        'Move Up
        If GetKeyState(vbKeyW) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirUp = False
        End If
    
        'Move Right
        If GetKeyState(vbKeyD) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
            Exit Sub
        Else
            DirRight = False
        End If
    
        'Move down
        If GetKeyState(vbKeyS) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirDown = False
        End If
    
        'Move left
        If GetKeyState(vbKeyA) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
            Exit Sub
        Else
            DirLeft = False
        End If
        'Move Up
        If GetKeyState(vbKeyUp) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirUp = False
        End If
    
        'Move Right
        If GetKeyState(vbKeyRight) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
            Exit Sub
        Else
            DirRight = False
        End If
    
        'Move down
        If GetKeyState(vbKeyDown) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirDown = False
        End If
    
        'Move left
        If GetKeyState(vbKeyLeft) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
            Exit Sub
        Else
            DirLeft = False
        End If

    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim chatText As String
Dim Name As String
Dim I As Long
Dim n As Long
Dim Command() As String
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    chatText = MyText
    
    If GUIWindow(GUI_CURRENCY).visible Then
        If (KeyAscii = vbKeyBack) Then
            If LenB(sDialogue) > 0 Then sDialogue = Mid$(sDialogue, 1, Len(sDialogue) - 1)
        End If
            
        ' And if neither, then add the character to the user's text buffer
        If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
            sDialogue = sDialogue & ChrW$(KeyAscii)
        End If
    End If
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
        chatOn = Not chatOn
        
        ' Broadcast message
        If Left$(chatText, 1) = "'" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call BroadcastMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Emote message
        If Left$(chatText, 1) = "-" Then
            MyText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call EmoteMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Player message
        If Left$(chatText, 1) = "!" Then
            If Mid$(chatText, 1, 2) = "! " Then GoTo Continue
            Name = vbNullString

            ' Get the desired player from the user text
            For I = 1 To Len(chatText)

                If Mid$(chatText, I, 1) <> Space(1) Then
                    Name = Name & Mid$(chatText, I, 1)
                Else
                    Exit For
                End If

            Next

            chatText = Mid$(chatText, I, Len(chatText) - 1)

            ' Make sure they are actually sending something
            If Len(chatText) - I > 0 Then
                chatText = Mid$(chatText, I + 1, Len(chatText) - I)
                ' Send the message to the player
                Call PlayerMsg(chatText, Name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case Command(0)
                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Broadcast Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /who, /fpslock", HelpColor)
                    Set Buffer = Nothing
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                    ' Checking fps
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' Request stats
                    ' // Monitor Admin Commands //
                    ' Admin Help
                Case "/admin"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue
                    frmAdminPanel.visible = Not frmAdminPanel.visible
                    ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    ' Map Editor
                Case "/editmap"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                    
                    SendRequestEditMap
                    ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If

                    ' Setting sprite
                Case "/setsprite"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    SendSetSprite CLng(Command(1))
                    ' Map report
                Case "/mapreport"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapReport
                    ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(chatText, Len(chatText) - 5)
                    ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendBanList
                    ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo Continue
                    End If

                    SendBan Command(1)
                    ' // Developer Admin Commands //
                    ' Editing item request
                Case "/edititem"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditItem
                ' Editing animation request
                Case "/editanimation"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditAnimation
                    ' Editing npc request
                Case "/editnpc"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditNpc
                Case "/editresource"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditResource
                    ' Editing shop request
                Case "/editshop"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditShop
                    ' Editing spell request
                Case "/editspell"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditSpell
                    ' // Creator Admin Commands //
                    ' Giving another player access
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                    ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    SendBanDestroy
                    ' Textures debug mode
                Case "/debug"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    DEBUG_MODE = Not DEBUG_MODE
                Case "/gui"
                    hideGUI = Not hideGUI
                Case "/quests"
                    GUIWindow(GUI_QUESTS).visible = Not GUIWindow(GUI_QUESTS).visible
                Case "/trade"
                    If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                        SendTradeRequest
                    Else
                        AddText "Antes de fazer uma troca, clique em cima do jogador com quem você deseja efetuá-la.", BrightRed
                    End If
                Case Else
                    AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
Continue:
            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Say message
        If Len(chatText) > 0 Then
            Call SayMsg(MyText)
        End If

        MyText = vbNullString
        UpdateShowChatText
        Exit Sub
    End If
    If Not chatOn Then Exit Sub
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
        UpdateShowChatText
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            MyText = MyText & ChrW$(KeyAscii)
            UpdateShowChatText
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub HandleMouseMove(ByVal X As Long, ByVal Y As Long, ByVal Button As Long)
Dim I As Long
    ' Set the global cursor position
    
    GlobalX = X
    GlobalY = Y
    GlobalX_Map = GlobalX + (TileView.Left * PIC_X) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.Top * PIC_Y) + Camera.Top
    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (X >= GUIWindow(I).X And X <= GUIWindow(I).X + GUIWindow(I).Width) And (Y >= GUIWindow(I).Y And Y <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_CHAT, GUI_BARS, GUI_MENU
                            ' Put nothing here and we can click through them!
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
    
    ' Handle the events
    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((Y + Camera.Top) \ PIC_Y)

    If InMapEditor Then
        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, X, Y)
        End If
    End If
End Sub
Public Sub HandleMouseDown(ByVal Button As Long)
Dim I As Long

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).X And GlobalX <= GUIWindow(I).X + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).Y And GlobalY <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_CHAT, GUI_BARS
                            ' Put nothing here and we can click through the
                        Case GUI_INVENTORY
                            Inventory_MouseDown Button
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_MouseDown Button
                            Exit Sub
                        Case GUI_MENU
                            Menu_MouseDown Button
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_MouseDown Button
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_MouseDown
                            Exit Sub
                        Case GUI_CURRENCY
                            Currency_MouseDown
                            Exit Sub
                        Case GUI_DIALOGUE
                            Dialogue_MouseDown
                            Exit Sub
                        Case GUI_SHOP
                            Shop_MouseDown
                            Exit Sub
                        Case GUI_PARTY
                            Party_MouseDown
                            Exit Sub
                        Case GUI_OPTIONS
                            Options_MouseDown
                            Exit Sub
                        Case GUI_TRADE
                            Trade_MouseDown
                            Exit Sub
                        Case GUI_EVENTCHAT
                            Chat_MouseDown
                            Exit Sub
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
        ' check chat buttons
        If Not inChat Then
            ChatScroll_MouseDown
        End If
    End If
    
    ' Handle events
    If InMapEditor Then
        Call MapEditorMouseDown(Button, GlobalX, GlobalY, False)
    Else
        ' left click
        If Button = vbLeftButton Then
            ' targetting
            FindTarget
            'FindTarget
        End If
    End If
End Sub

Public Sub HandleMouseUp(ByVal Button As Long)
Dim I As Long

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).X And GlobalX <= GUIWindow(I).X + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).Y And GlobalY <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_CHAT, GUI_BARS
                            ' Put nothing here and we can click through the
                        Case GUI_INVENTORY
                            Inventory_MouseUp
                        Case GUI_SPELLS
                            Spells_MouseUp
                        Case GUI_MENU
                            Menu_MouseUp
                        Case GUI_HOTBAR
                            Hotbar_MouseUp
                        Case GUI_CHARACTER
                            Character_MouseUp
                        Case GUI_CURRENCY
                            Currency_MouseUp
                        Case GUI_DIALOGUE
                            Dialogue_MouseUp
                        Case GUI_SHOP
                            Shop_MouseUp
                        Case GUI_PARTY
                            Party_MouseUp
                        Case GUI_OPTIONS
                            Options_MouseUp
                        Case GUI_TRADE
                            Trade_MouseUp
                        Case GUI_EVENTCHAT
                            Chat_MouseUp
                    End Select
                End If
            End If
        Next
    End If

    ' Stop dragging if we haven't catched it already
    DragInvSlotNum = 0
    DragBankSlotNum = 0
    DragSpell = 0
    ' reset buttons
    resetClickedButtons
    ' stop scrolling chat
    ChatButtonUp = False
    ChatButtonDown = False
End Sub

Public Sub HandleDoubleClick()
Dim I As Long

    On Error Resume Next
    If frmEditor_Events.visible = True Then
        If frmEditor_Events.fraMapWarp.visible = True Then
            frmEditor_Events.scrlWarpMap.Value = GetPlayerMap(MyIndex)
            frmEditor_Events.scrlWarpX.Value = CurX
            frmEditor_Events.scrlWarpY.Value = CurY
            frmEditor_Events.SetFocus
            Exit Sub
        End If
    End If

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).X And GlobalX <= GUIWindow(I).X + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).Y And GlobalY <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_INVENTORY
                            Inventory_DoubleClick
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_DoubleClick
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_DoubleClick
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_DoubleClick
                            Exit Sub
                        Case GUI_SHOP
                            Shop_DoubleClick
                            Exit Sub
                        Case GUI_BANK
                            Bank_DoubleClick
                            Exit Sub
                        Case GUI_TRADE
                            Trade_DoubleClick
                            Exit Sub
                        Case GUI_QUESTS
                            Quests_DoubleClick
                            Exit Sub
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
End Sub

Public Sub OpenGuiWindow(ByVal Index As Long)
Dim Buffer As clsBuffer
    If Index = 1 Then
        GUIWindow(GUI_INVENTORY).visible = Not GUIWindow(GUI_INVENTORY).visible
    Else
        GUIWindow(GUI_INVENTORY).visible = False
    End If
    
    If Index = 2 Then
        GUIWindow(GUI_SPELLS).visible = Not GUIWindow(GUI_SPELLS).visible
        ' Update the spells on the pic
        Set Buffer = New clsBuffer
        Buffer.WriteLong CSpells
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    Else
        GUIWindow(GUI_SPELLS).visible = False
    End If
    
    If Index = 3 Then
        GUIWindow(GUI_CHARACTER).visible = Not GUIWindow(GUI_CHARACTER).visible
    Else
        GUIWindow(GUI_CHARACTER).visible = False
    End If
    
    If Index = 4 Then
        GUIWindow(GUI_OPTIONS).visible = Not GUIWindow(GUI_OPTIONS).visible
    Else
        GUIWindow(GUI_OPTIONS).visible = False
    End If
    
    If Index = 5 Then
        GUIWindow(GUI_QUESTS).visible = Not GUIWindow(GUI_QUESTS).visible
    Else
        GUIWindow(GUI_QUESTS).visible = False
    End If
    
    If Index = 6 Then
        GUIWindow(GUI_PARTY).visible = Not GUIWindow(GUI_PARTY).visible
    Else
        GUIWindow(GUI_PARTY).visible = False
    End If
End Sub

Public Sub Currency_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        CurrencyAcceptState = 2 ' clicked
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        CurrencyCloseState = 2 ' clicked
    End If
End Sub
Public Sub Currency_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long, Buffer As clsBuffer
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        If CurrencyAcceptState = 2 Then
            ' do stuffs
            If IsNumeric(sDialogue) Then
                Select Case CurrencyMenu
                    Case 1 ' drop item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        SendDropItem tmpCurrencyItem, Val(sDialogue)
                    Case 2 ' deposit item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        DepositItem tmpCurrencyItem, Val(sDialogue)
                    Case 3 ' withdraw item
                        If Val(sDialogue) > GetBankItemValue(tmpCurrencyItem) Then sDialogue = GetBankItemValue(tmpCurrencyItem)
                        WithdrawItem tmpCurrencyItem, Val(sDialogue)
                    Case 4 ' offer trade item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        TradeItem tmpCurrencyItem, Val(sDialogue)
                End Select
            Else
                AddText "Please enter a valid amount.", BrightRed
                Exit Sub
            End If
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    ' check if we're on the button
    If (GlobalX >= X And GlobalX <= X + Buttons(12).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(12).Height) Then
        If CurrencyCloseState = 2 Then
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    
    CurrencyAcceptState = 0
    CurrencyCloseState = 0
    GUIWindow(GUI_CURRENCY).visible = False
    inChat = False
    chatOn = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    sDialogue = vbNullString
    ' reset buttons
    resetClickedButtons
End Sub
Public Sub Dialogue_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 90
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(1) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(2) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 120
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(3) = 2 ' clicked
        End If
    End If
End Sub

Public Sub Dialogue_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        X = GUIWindow(GUI_CHAT).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 90
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(1) = 2 Then
                Dialogue_Button_MouseDown (2)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(1) = 0
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        X = GUIWindow(GUI_CHAT).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(2) = 2 Then
                Dialogue_Button_MouseDown (1)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(2) = 0
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        X = GUIWindow(GUI_CHAT).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 120
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(3) = 2 Then
                Dialogue_Button_MouseDown (3)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(3) = 0
    End If
End Sub

' scroll bar
Public Sub ChatScroll_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    ' find out which button we're clicking
    For I = 34 To 35
        X = GUIWindow(GUI_CHAT).X + Buttons(I).X
        Y = GUIWindow(GUI_CHAT).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).State = 2 ' clicked
            ' scroll the actual chat
            Select Case I
                Case 34 ' up
                    'ChatScroll = ChatScroll + 1
                    ChatButtonUp = True
                Case 35 ' down
                    'ChatScroll = ChatScroll - 1
                    'If ChatScroll < 8 Then ChatScroll = 8
                    ChatButtonDown = True
            End Select
        End If
    Next
End Sub

' Shop
Public Sub Shop_MouseUp()
Dim I As Long, X As Long, Y As Long, Buffer As clsBuffer

    ' find out which button we're clicking
    For I = 23 To 23
        X = GUIWindow(GUI_SHOP).X + Buttons(I).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).State = 2 Then
                ' do stuffs
                Select Case I
                    Case 23
                        ' exit
                        Set Buffer = New clsBuffer
                        Buffer.WriteLong CCloseShop
                        SendData Buffer.ToArray()
                        Set Buffer = Nothing
                        GUIWindow(GUI_SHOP).visible = False
                        InShop = 0
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Shop_MouseDown()
Dim I As Long, X As Long, Y As Long

    ' find out which button we're clicking
    For I = 23 To 23
        X = GUIWindow(GUI_SHOP).X + Buttons(I).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).State = 2 ' clicked
        End If
    Next
End Sub

Public Sub Shop_DoubleClick()
Dim shopSlot As Long

    shopSlot = IsShopItem(GlobalX, GlobalY)

    If shopSlot > 0 Then
        ' buy item code
        BuyItem shopSlot
    End If
End Sub
Public Sub Bank_DoubleClick()
Dim bankNum As Long
    bankNum = IsBankItem(GlobalX, GlobalY)
    If bankNum <> 0 Then
        If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetBankItemNum(bankNum)).Stackable > 0 Then
            CurrencyMenu = 3 ' withdraw
            CurrencyText = "How many do you want withdraw?"
            tmpCurrencyItem = bankNum
            sDialogue = vbNullString
            GUIWindow(GUI_CURRENCY).visible = True
            inChat = True
            chatOn = True
            Exit Sub
        End If
        WithdrawItem bankNum, 0
        Exit Sub
    End If
End Sub
Public Sub Trade_DoubleClick()
Dim tradeNum As Long
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum <> 0 Then
        UntradeItem tradeNum
        Exit Sub
    End If
End Sub
Public Sub Trade_MouseDown()
Dim I As Long, X As Long, Y As Long

    ' find out which button we're clicking
    For I = 40 To 41
        X = Buttons(I).X
        Y = Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).State = 2 ' clicked
        End If
    Next
End Sub
Public Sub Trade_MouseUp()
Dim I As Long, X As Long, Y As Long, Buffer As clsBuffer

    ' find out which button we're clicking
    For I = 40 To 41
        X = Buttons(I).X
        Y = Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).State = 2 Then
                ' do stuffs
                Select Case I
                    Case 40
                        AcceptTrade
                    Case 41
                        DeclineTrade
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

' Party
Public Sub Party_MouseUp()
Dim I As Long, X As Long, Y As Long, Buffer As clsBuffer

    ' find out which button we're clicking
    For I = 24 To 25
        X = GUIWindow(GUI_PARTY).X + Buttons(I).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).State = 2 Then
                ' do stuffs
                Select Case I
                    Case 24 ' invite
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendPartyRequest
                        Else
                            AddText "Invalid invitation target.", BrightRed
                        End If
                    Case 25 ' leave
                        If Party.Leader > 0 Then
                            SendPartyLeave
                        Else
                            AddText "You are not in a party.", BrightRed
                        End If
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Party_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 24 To 25
        X = GUIWindow(GUI_PARTY).X + Buttons(I).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).State = 2 ' clicked
        End If
    Next
End Sub

'Options
Public Sub Options_MouseUp()
Dim I As Long, X As Long, Y As Long, Buffer As clsBuffer, layerNum As Long

    ' find out which button we're clicking
    For I = 26 To 31
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).State = 3 Then
                ' do stuffs
                Select Case I
                    Case 26 ' music on
                        Options.Music = 1
                        PlayMusic Trim$(Map.Music)
                        SaveOptions
                        Buttons(26).State = 2
                        Buttons(27).State = 0
                    Case 27 ' music off
                        Options.Music = 0
                        StopMusic
                        SaveOptions
                        Buttons(26).State = 0
                        Buttons(27).State = 2
                    Case 28 ' sound on
                        Options.Sound = 1
                        SaveOptions
                        Buttons(28).State = 2
                        Buttons(29).State = 0
                    Case 29 ' sound off
                        Options.Sound = 0
                        StopAllSounds
                        SaveOptions
                        Buttons(28).State = 0
                        Buttons(29).State = 2
                    Case 30 ' debug on
                        Options.Debug = 1
                        SaveOptions
                        Buttons(30).State = 2
                        Buttons(31).State = 0
                    Case 31 ' debug off
                        Options.Debug = 0
                        SaveOptions
                        Buttons(30).State = 0
                        Buttons(31).State = 2
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    For I = 42 To 45
    ' set co-ordinate
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).State = 2 Then
                Select Case I
                    Case 42
                        If Options.FPS = 15 Then Options.FPS = 20
                        SaveOptions
                    Case 43
                        If Options.FPS = 20 Then Options.FPS = 15
                        SaveOptions
                    Case 44
                        If Options.volume - 10 >= 0 Then
                            Options.volume = Options.volume - 10
                            StopMusic
                            PlayMusic Trim$(Map.Music)
                        Else
                            Options.volume = 0
                        End If
                        SaveOptions
                    Case 45
                        If Options.volume + 10 <= 150 Then
                            Options.volume = Options.volume + 10
                            StopMusic
                            PlayMusic Trim$(Map.Music)
                        Else
                            Options.volume = 150
                        End If
                        SaveOptions
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Options_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 26 To 31
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).State = 0 Then
                Buttons(I).State = 3 ' clicked
            End If
        End If
    Next
    For I = 42 To 45
    ' set co-ordinate
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).State = 2 ' clicked
        End If
    Next
End Sub

' Menu
Public Sub Menu_MouseUp()
Dim I As Long, X As Long, Y As Long, Buffer As clsBuffer

    ' find out which button we're clicking
    For I = 1 To 6
        X = GUIWindow(GUI_MENU).X + Buttons(I).X
        Y = GUIWindow(GUI_MENU).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).State = 2 Then
                ' do stuffs
                Select Case I
                    Case 1
                        ' open window
                        OpenGuiWindow 1
                    Case 2
                        ' open window
                        OpenGuiWindow 2
                    Case 3
                        ' open window
                        OpenGuiWindow 3
                    Case 4
                        ' open window
                        OpenGuiWindow 4
                    Case 5
                        OpenGuiWindow 5
                    Case 6
                        ' open window
                        OpenGuiWindow 6
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Menu_MouseDown(ByVal Button As Long)
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 1 To 6
        If Buttons(I).visible Then
            X = GUIWindow(GUI_MENU).X + Buttons(I).X
            Y = GUIWindow(GUI_MENU).Y + Buttons(I).Y
            ' check if we're on the button
            If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
                Buttons(I).State = 2 ' clicked
            End If
        End If
    Next
End Sub

' Inventory
Public Sub Inventory_MouseUp()
Dim invSlot As Long
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        invSlot = IsInvItem(GlobalX, GlobalY, True)
        If invSlot = 0 Then Exit Sub
        ' change slots
        SendChangeInvSlots DragInvSlotNum, invSlot
    End If

    DragInvSlotNum = 0
End Sub

Public Sub Inventory_MouseDown(ByVal Button As Long)
Dim invNum As Long

    invNum = IsInvItem(GlobalX, GlobalY)

    If Button = 1 Then
        If invNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = invNum
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If invNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                    If GetPlayerInvItemValue(MyIndex, invNum) > 0 Then
                        CurrencyMenu = 1 ' drop
                        CurrencyText = "How many do you want to drop?"
                        tmpCurrencyItem = invNum
                        sDialogue = vbNullString
                        GUIWindow(GUI_CURRENCY).visible = True
                        inChat = True
                        chatOn = True
                    End If
                Else
                    Call SendDropItem(invNum, 0)
                End If
            End If
        End If
    End If
End Sub

Public Sub Inventory_DoubleClick()
    Dim invNum As Long, Value As Long, multiplier As Double, I As Long

    DragInvSlotNum = 0
    invNum = IsInvItem(GlobalX, GlobalY)

    If invNum > 0 Then
        ' are we in a shop?
        If InShop > 0 Then
            SellItem invNum
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                CurrencyMenu = 2 ' deposit
                CurrencyText = "How many do you want to deposit?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).visible = True
                inChat = True
                chatOn = True
                Exit Sub
            End If
                
            Call DepositItem(invNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For I = 1 To MAX_INV
                If TradeYourOffer(I).Num = invNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).Num)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).Num)).Stackable > 0 Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(I).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(I).Num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                CurrencyMenu = 4 ' offer in trade
                CurrencyText = "How many do you want to trade?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).visible = True
                inChat = True
                chatOn = True
                Exit Sub
            End If
            
            Call TradeItem(invNum, 0)
            Exit Sub
        End If

        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_SCOUTER Then
            PlaySound "ligando scouter.mp3", -1, -1
            ScouterOn = Not ScouterOn
            myTarget = 0
            myTargetType = 0
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(invNum)
        Exit Sub
    End If
End Sub

' Spells
Public Sub Spells_DoubleClick()
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)

    If spellnum <> 0 Then
        Call CastSpell(spellnum)
        Exit Sub
    End If
End Sub

Public Sub Spells_MouseDown(ByVal Button As Long)
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)
    If Button = 1 Then ' left click
        If spellnum <> 0 Then
            DragSpell = spellnum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If spellnum <> 0 Then
            If PlayerSpells(spellnum) > 0 Then
                Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(Spell(PlayerSpells(spellnum)).Name) & "?", DIALOGUE_TYPE_FORGET, True, spellnum
            End If
        End If
    End If
End Sub

Public Sub Spells_MouseUp()
Dim spellSlot As Long

    If DragSpell > 0 Then
        spellSlot = IsPlayerSpell(GlobalX, GlobalY, True)
        If spellSlot = 0 Then Exit Sub
        SendChangeSpellSlots DragSpell, spellSlot
    End If

    DragSpell = 0
End Sub

' character
Public Sub Character_DoubleClick()
Dim eqNum As Long

    eqNum = IsEqItem(GlobalX, GlobalY)

    If eqNum <> 0 Then
        SendUnequip eqNum
    End If
End Sub
' hotbar
Public Sub Hotbar_DoubleClick()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarUse slotNum
    End If
End Sub

Public Sub Hotbar_MouseDown(ByVal Button As Long)
Dim slotNum As Long
    
    If Button <> 2 Then Exit Sub ' right click
    
    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarChange 0, 0, slotNum
    End If
End Sub

Public Sub Hotbar_MouseUp()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum = 0 Then Exit Sub
    
    ' inventory
    If DragInvSlotNum > 0 Then
        SendHotbarChange 1, DragInvSlotNum, slotNum
        DragInvSlotNum = 0
        Exit Sub
    End If
    
    ' spells
    If DragSpell > 0 Then
        SendHotbarChange 2, DragSpell, slotNum
        DragSpell = 0
        Exit Sub
    End If
End Sub
Public Sub Dialogue_Button_MouseDown(Index As Integer)
    ' call the handler
    dialogueHandler Index
    GUIWindow(GUI_DIALOGUE).visible = False
    inChat = False
    dialogueIndex = 0
End Sub
Public Sub Character_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 16 To 20
        X = GUIWindow(GUI_CHARACTER).X + Buttons(I).X
        Y = GUIWindow(GUI_CHARACTER).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).State = 2 ' clicked
        End If
    Next
End Sub

Public Sub Character_MouseUp()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = 16 To 20
        X = GUIWindow(GUI_CHARACTER).X + Buttons(I).X
        Y = GUIWindow(GUI_CHARACTER).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' send the level up
            If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
            SendTrainStat (I - 15)
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    Next
End Sub
' Npc Chat
Public Sub Chat_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long

Select Case CurrentEvent.Type
    Case Evt_Menu
    For I = 1 To UBound(CurrentEvent.Text) - 1
        If Len(Trim$(CurrentEvent.Text(I + 1))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(CurrentEvent.Text(I + 1)) & "]")
            X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
            Y = GUIWindow(GUI_EVENTCHAT).Y + 115 - ((I - 1) * 15)
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                chatOptState(I) = 2 ' clicked
            End If
        End If
    Next
    Case Evt_Message
    Width = EngineGetTextWidth(Font_Default, "[Continue]")
    X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
    Y = GUIWindow(GUI_EVENTCHAT).Y + 100
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        chatContinueState = 2 ' clicked
    End If
End Select

End Sub
Public Sub Chat_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long

Select Case CurrentEvent.Type
    Case Evt_Menu
        For I = 1 To UBound(CurrentEvent.Text) - 1
            If Len(Trim$(CurrentEvent.Text(I + 1))) > 0 Then
                Width = EngineGetTextWidth(Font_Default, "[" & Trim$(CurrentEvent.Text(I + 1)) & "]")
                X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
                Y = GUIWindow(GUI_EVENTCHAT).Y + 115 - ((I - 1) * 15)
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' are we clicked?
                    If chatOptState(I) = 2 Then
                        Events_SendChooseEventOption CurrentEvent.Data(I)
                        ' play sound
                        PlaySound Sound_ButtonClick, -1, -1
                    End If
                End If
            End If
        Next
        
        For I = 1 To UBound(CurrentEvent.Text) - 1
            chatOptState(I) = 0 ' normal
        Next
    Case Evt_Message
        Width = EngineGetTextWidth(Font_Default, "[Continue]")
        X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
        Y = GUIWindow(GUI_EVENTCHAT).Y + 100
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' are we clicked?
            If chatContinueState = 2 Then
                Events_SendChooseEventOption CurrentEventIndex + 1
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        
        chatContinueState = 0
End Select
End Sub
Public Sub HandleKeyUp(ByVal KeyCode As Long)
Dim I As Long
Dim Buffer As clsBuffer
    ' hotbar
    If Not chatOn Then
        For I = 1 To 9
            If KeyCode = 48 + I Then
                SendHotbarUse I
            End If
        Next
        If KeyCode = 48 Then ' 0
            SendHotbarUse 10
        ElseIf KeyCode = 189 Then ' -
            SendHotbarUse 11
        ElseIf KeyCode = 187 Then ' =
            SendHotbarUse 12
        End If
    End If
    
    If KeyCode = vbKeyControl Then
        If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_RESOURCE Then
            If Resource(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Data1).ResourceType = 3 Then
                If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(MyIndex, Weapon)).Data3 = 2 Then
                        FishingTime = 0
                        If FishTime > GetTickCount Then
                            Set Buffer = New clsBuffer
                            Buffer.WriteLong CAttack
                            SendData Buffer.ToArray()
                            Set Buffer = Nothing
                            FishTime = 0
                            Else
                            AlertX = GetTickCount + 500
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Sub
'Quests
Public Sub Quests_DoubleClick()
Dim QuestNum As Long
    
    QuestNum = IsPlayerQuest(GlobalX, GlobalY)
    
    If QuestNum <> 0 Then
        Call SendQuestInfo(QuestNum)
        Exit Sub
    End If
End Sub
