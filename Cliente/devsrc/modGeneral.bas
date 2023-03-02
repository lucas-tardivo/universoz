Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
'Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' Text API
Declare Function GeneralWinDirApi Lib "kernel32" _
        Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
        ByVal nSize As Long) As Long
        
Function GetTickCount() As Currency
    GetTickCount = timeGetTime
    If GetTickCount < 0 Then GetTickCount = GetTickCount + MAX_LONG
End Function

Public Sub Main()
Dim I As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InitialiseGUI
    ' load options
    Call SetStatus("Loading Options...")
    LoadOptions
    
    ' load main menu
    Call SetStatus("Loading Menu...")
    ' general menu stuff
    frmMenu.Caption = Options.Game_Name & " - Developer suite"

    ' Load the username + pass
    frmMenu.txtLUser.Text = Trim$(Options.Username)
    frmMenu.txtLIP.Text = Trim$(Options.IP)
    frmMenu.txtLPort.Text = Trim$(Options.Port)
    If Options.savePass = 1 Then
        frmMenu.txtLPass.Text = Trim$(Options.Password)
        frmMenu.chkPass.Value = Options.savePass
    End If
    ' we can now see it
    frmMenu.visible = True
    
    ' load gui
    Call SetStatus("Loading interface...")
    InitialiseGUI
    
    setOptionsState
    
    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\data files\", "graphics"
    ChkDir App.Path & "\data files\graphics\", "animations"
    ChkDir App.Path & "\data files\graphics\", "characters"
    ChkDir App.Path & "\data files\graphics\", "items"
    ChkDir App.Path & "\data files\graphics\", "paperdolls"
    ChkDir App.Path & "\data files\graphics\", "resources"
    ChkDir App.Path & "\data files\graphics\", "spellicons"
    ChkDir App.Path & "\data files\graphics\", "tilesets"
    ChkDir App.Path & "\data files\graphics\", "faces"
    ChkDir App.Path & "\data files\graphics\", "gui"
    ChkDir App.Path & "\data files\graphics\gui\", "menu"
    ChkDir App.Path & "\data files\graphics\gui\menu\", "buttons"
    ChkDir App.Path & "\data files\", "logs"
    ChkDir App.Path & "\data files\", "maps"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"
    
    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34) ' "
    
    ' Update the form with the game's name before it's loaded
    frmMain.Caption = Options.Game_Name & " - Developer suite"
    
    EngineInitFontSettings
    
    InitDX8
    
    ' randomize rnd's seed
    Randomize
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    Call InitMessages
    Call SetStatus("Initializing DirectX...")
    
    ' load music/sound engine
    InitFmod
    
    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMusic Trim$(Options.MenuMusic)
    
    ' Reset values
    Ping = -1

    ' cache the buttons then reset & render them
    Call SetStatus("Loading data...")
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.Weapon

    frmMain.Width = 12090
    frmMain.Height = 9420
    Call SetStatus("All editors grouped for game developers")
    frmMenu.lblLAccept.Enabled = True
    MenuLoop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub InitialiseGUI(Optional ByVal loadingScreen As Boolean = False)

'Loading Interface.ini data
Dim filename As String
filename = App.Path & "\data files\interface.ini"
Dim I As Long
    ' re-set chat scroll
    ChatScroll = 8
    ReDim GUIWindow(1 To GUI_Count) As GUIWindowRec
    
    ' 1 - Chat
    With GUIWindow(GUI_CHAT)
        .X = Val(GetVar(filename, "GUI_CHAT", "X"))
        .Y = Val(GetVar(filename, "GUI_CHAT", "Y")) - 15
        .Width = Val(GetVar(filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHAT", "Height"))
        .visible = True
    End With
    
    ' 2 - Hotbar
    With GUIWindow(GUI_HOTBAR)
        .X = Val(GetVar(filename, "GUI_HOTBAR", "X"))
        .Y = Val(GetVar(filename, "GUI_HOTBAR", "Y"))
        .Height = Val(GetVar(filename, "GUI_HOTBAR", "Height"))
        .Width = ((9 + 36) * (MAX_HOTBAR - 1))
    End With
    
    ' 3 - Menu
    With GUIWindow(GUI_MENU)
        .X = Val(GetVar(filename, "GUI_MENU", "X"))
        .Y = Val(GetVar(filename, "GUI_MENU", "Y")) - 15
        .Width = Val(GetVar(filename, "GUI_MENU", "Width"))
        .Height = Val(GetVar(filename, "GUI_MENU", "Height"))
        .visible = True
    End With
    
    ' 4 - Bars
    With GUIWindow(GUI_BARS)
        .X = Val(GetVar(filename, "GUI_BARS", "X"))
        .Y = Val(GetVar(filename, "GUI_BARS", "Y"))
        .Width = Val(GetVar(filename, "GUI_BARS", "Width"))
        .Height = Val(GetVar(filename, "GUI_BARS", "Height"))
        .visible = True
    End With
    
    ' 5 - Inventory
    With GUIWindow(GUI_INVENTORY)
        .X = Val(GetVar(filename, "GUI_INVENTORY", "X"))
        .Y = Val(GetVar(filename, "GUI_INVENTORY", "Y")) - 15
        .Width = Val(GetVar(filename, "GUI_INVENTORY", "Width"))
        .Height = Val(GetVar(filename, "GUI_INVENTORY", "Height"))
        .visible = False
    End With
    
    ' 6 - Spells
    With GUIWindow(GUI_SPELLS)
        .X = Val(GetVar(filename, "GUI_SPELLS", "X"))
        .Y = Val(GetVar(filename, "GUI_SPELLS", "Y")) - 15
        .Width = Val(GetVar(filename, "GUI_SPELLS", "Width"))
        .Height = Val(GetVar(filename, "GUI_SPELLS", "Height"))
        .visible = False
    End With
    
    ' 7 - Character
    With GUIWindow(GUI_CHARACTER)
        .X = Val(GetVar(filename, "GUI_CHARACTER", "X"))
        .Y = Val(GetVar(filename, "GUI_CHARACTER", "Y")) - 15
        .Width = Val(GetVar(filename, "GUI_CHARACTER", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHARACTER", "Height"))
        .visible = False
    End With
    
    ' 8 - Options
    With GUIWindow(GUI_OPTIONS)
        .X = Val(GetVar(filename, "GUI_OPTIONS", "X"))
        .Y = Val(GetVar(filename, "GUI_OPTIONS", "Y")) - 15
        .Width = Val(GetVar(filename, "GUI_OPTIONS", "Width"))
        .Height = Val(GetVar(filename, "GUI_OPTIONS", "Height"))
        .visible = False
    End With
    
    ' 9 - Party
    With GUIWindow(GUI_PARTY)
        .X = Val(GetVar(filename, "GUI_PARTY", "X"))
        .Y = Val(GetVar(filename, "GUI_PARTY", "Y")) - 15
        .Width = Val(GetVar(filename, "GUI_PARTY", "Width"))
        .Height = Val(GetVar(filename, "GUI_PARTY", "Height"))
        .visible = False
    End With
    
    ' 10 - Description
    With GUIWindow(GUI_DESCRIPTION)
        .X = Val(GetVar(filename, "GUI_DESCRIPTION", "X"))
        .Y = Val(GetVar(filename, "GUI_DESCRIPTION", "Y"))
        .Width = Val(GetVar(filename, "GUI_DESCRIPTION", "Width"))
        .Height = Val(GetVar(filename, "GUI_DESCRIPTION", "Height"))
        .visible = False
    End With
    
    ' 11 - Main Menu
    With GUIWindow(GUI_MAINMENU)
        .X = Val(GetVar(filename, "GUI_MAINMENU", "X"))
        .Y = Val(GetVar(filename, "GUI_MAINMENU", "Y"))
        .Width = Val(GetVar(filename, "GUI_MAINMENU", "Width"))
        .Height = Val(GetVar(filename, "GUI_MAINMENU", "Height"))
        .visible = False
    End With
    
    ' 12 - Shop
    With GUIWindow(GUI_SHOP)
        .X = Val(GetVar(filename, "GUI_SHOP", "X"))
        .Y = Val(GetVar(filename, "GUI_SHOP", "Y"))
        .Width = Val(GetVar(filename, "GUI_SHOP", "Width"))
        .Height = Val(GetVar(filename, "GUI_SHOP", "Height"))
        .visible = False
    End With
    
    ' 13 - Bank
    With GUIWindow(GUI_BANK)
        .X = 5
        .Y = 62
        .Width = 480
        .Height = 384
        .visible = False
    End With
    
    ' 14 - Trade
    With GUIWindow(GUI_TRADE)
        .X = 5
        .Y = 62
        .Width = 480
        .Height = 384
        .visible = False
    End With
    
    ' 15 - Currency
    With GUIWindow(GUI_CURRENCY)
        .X = Val(GetVar(filename, "GUI_CHAT", "X"))
        .Y = Val(GetVar(filename, "GUI_CHAT", "Y")) - 15
        .Width = Val(GetVar(filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHAT", "Height"))
        .visible = False
    End With
    ' 16 - Dialogue
    With GUIWindow(GUI_DIALOGUE)
        .X = Val(GetVar(filename, "GUI_CHAT", "X"))
        .Y = Val(GetVar(filename, "GUI_CHAT", "Y")) - 15
        .Width = Val(GetVar(filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHAT", "Height"))
        .visible = False
    End With
    ' 17 - Event Chat
    With GUIWindow(GUI_EVENTCHAT)
        .X = Val(GetVar(filename, "GUI_CHAT", "X"))
        .Y = Val(GetVar(filename, "GUI_CHAT", "Y")) - 15
        .Width = Val(GetVar(filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(filename, "GUI_CHAT", "Height"))
        .visible = False
    End With
    ' 18 - News
    With GUIWindow(GUI_NEWS)
        .X = 0
        .Y = 0
        .Width = 320
        .Height = 320
        .visible = False
    End With
    ' 19 - Quests
    With GUIWindow(GUI_QUESTS)
        .X = 578
        .Y = 255
        .Width = 195
        .Height = 250
        .visible = False
    End With
    ' BUTTONS
    With Buttons(1)
        .State = 0 ' normal
        .X = 111
        .Y = 11
        .Width = 36
        .Height = 36
        .visible = True
        .PicNum = 1
    End With
    
    ' main - skills
    With Buttons(2)
        .State = 0 ' normal
        .X = 153
        .Y = 11
        .Width = 36
        .Height = 36
        .visible = True
        .PicNum = 2
    End With
    
    ' main - char
    With Buttons(3)
        .State = 0 ' normal
        .X = 195
        .Y = 11
        .Width = 36
        .Height = 36
        .visible = True
        .PicNum = 3
    End With
    
    ' main - opt
    With Buttons(4)
        .State = 0 ' normal
        .X = 237
        .Y = 11
        .Width = 36
        .Height = 36
        .visible = True
        .PicNum = 4
    End With
    
    ' main - trade
    With Buttons(5)
        .State = 0 ' normal
        .X = 279
        .Y = 11
        .Width = 36
        .Height = 36
        .visible = True
        .PicNum = 5
    End With
    
    ' main - party
    With Buttons(6)
        .State = 0 ' normal
        .X = 321
        .Y = 11
        .Width = 36
        .Height = 36
        .visible = True
        .PicNum = 6
    End With
    
    
    
    ' menu - login
    With Buttons(7)
        .State = 0 ' normal
        .X = 172
        .Y = 481
        .Width = 110
        .Height = 32
        .visible = True
        .PicNum = 7
    End With
    
    ' menu - register
    With Buttons(8)
        .State = 0 ' normal
        .X = 302
        .Y = 481
        .Width = 110
        .Height = 32
        .visible = True
        .PicNum = 8
    End With
    
    ' menu - credits
    With Buttons(9)
        .State = 0 ' normal
        .X = 432
        .Y = 481
        .Width = 110
        .Height = 32
        .visible = True
        .PicNum = 9
    End With
    
    ' menu - exit
    With Buttons(10)
        .State = 0 ' normal
        .X = 562
        .Y = 481
        .Width = 110
        .Height = 32
        .visible = True
        .PicNum = 10
    End With
    
    ' menu - Login Accept
    With Buttons(11)
        .State = 0 ' normal
        .X = 350
        .Y = 368
        .Width = 110
        .Height = 32
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Register Accept
    With Buttons(12)
        .State = 0 ' normal
        .X = 350
        .Y = 373
        .Width = 110
        .Height = 32
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Accept
    With Buttons(13)
        .State = 0 ' normal
        .X = 350
        .Y = 445
        .Width = 110
        .Height = 32
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Next
    With Buttons(14)
        .State = 0 ' normal
        .X = 348
        .Y = 445
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 12
    End With
    
    ' menu - NewChar Accept
    With Buttons(15)
        .State = 0 ' normal
        .X = 350
        .Y = 373
        .Width = 110
        .Height = 32
        .visible = True
        .PicNum = 11
    End With
    
    ' main - AddStats
    For I = 16 To 20
        With Buttons(I)
            .State = 0 'normal
            .Width = 12
            .Height = 11
            .visible = True
            .PicNum = 13
        End With
    Next
    ' set the individual spaces
    For I = 16 To 20 ' first 3
        With Buttons(I)
            .X = 165
            .Y = 64 + ((I - 16) * 30)
        End With
    Next
    
    ' main - shop buy
    With Buttons(21)
        .State = 0 ' normal
        .X = 12
        .Y = 276
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 14
    End With
    
    ' main - shop sell
    With Buttons(22)
        .State = 0 ' normal
        .X = 90
        .Y = 276
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 15
    End With
    
    ' main - shop exit
    With Buttons(23)
        .State = 0 ' normal
        .X = 90
        .Y = 276
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 16
    End With
    
    ' main - party invite
    With Buttons(24)
        .State = 0 ' normal
        .X = 14
        .Y = 209
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 17
    End With
    
    ' main - party invite
    With Buttons(25)
        .State = 0 ' normal
        .X = 101
        .Y = 209
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 18
    End With
    
    ' main - music on
    With Buttons(26)
        .State = 0 ' normal
        .X = 77
        .Y = 14
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - music off
    With Buttons(27)
        .State = 0 ' normal
        .X = 132
        .Y = 14
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - sound on
    With Buttons(28)
        .State = 0 ' normal
        .X = 77
        .Y = 39
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - sound off
    With Buttons(29)
        .State = 0 ' normal
        .X = 132
        .Y = 39
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - debug on
    With Buttons(30)
        .State = 0 ' normal
        .X = 77
        .Y = 64
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - debug off
    With Buttons(31)
        .State = 0 ' normal
        .X = 132
        .Y = 64
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - scroll up
    With Buttons(34)
        .State = 0 ' normal
        .X = 391
        .Y = 2
        .Width = 19
        .Height = 19
        .visible = True
        .PicNum = 21
    End With
    
    ' main - scroll down
    With Buttons(35)
        .State = 0 ' normal
        .X = 391
        .Y = 105
        .Width = 19
        .Height = 19
        .visible = True
        .PicNum = 22
    End With
    ' main - Select Gender Left
        With Buttons(36)
            .State = 0 'normal
            .X = 327
            .Y = 318
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 23
        End With
        
    ' main - Select Gender Right
        With Buttons(37)
            .State = 0 'normal
            .X = 363
            .Y = 318
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 24
        End With
    
    ' main - Select Hair Left
        With Buttons(38)
            .State = 0 'normal
            .X = 327
            .Y = 345
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 23
        End With
        
    ' main - Select Gender Right
        With Buttons(39)
            .State = 0 'normal
            .X = 363
            .Y = 345
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 24
        End With
    ' main - Accept Trade
        With Buttons(40)
            .State = 0 'normal
            .X = GUIWindow(GUI_TRADE).X + 165
            .Y = GUIWindow(GUI_TRADE).Y + 335
            .Width = 69
            .Height = 29
            .visible = True
            .PicNum = 25
        End With
    ' main - Decline Trade
        With Buttons(41)
            .State = 0 'normal
            .X = GUIWindow(GUI_TRADE).X + 245
            .Y = GUIWindow(GUI_TRADE).Y + 335
            .Width = 69
            .Height = 29
            .visible = True
            .PicNum = 26
        End With
    ' main - FPS Cap left
        With Buttons(42)
            .State = 0 'normal
            .X = 92
            .Y = 100
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 23
        End With
    ' main - FPS Cap Right
        With Buttons(43)
            .State = 0 'normal
            .X = 147
            .Y = 100
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 24
        End With
    ' main - Volume left
        With Buttons(44)
            .State = 0 'normal
            .X = 92
            .Y = 120
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 23
        End With
    ' main - Volume Right
        With Buttons(45)
            .State = 0 'normal
            .X = 147
            .Y = 120
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 24
        End With
End Sub
Public Sub MenuState()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Options.IP = frmMenu.txtLIP.Text
    Options.Port = frmMenu.txtLPort.Text
    frmMain.Socket.Close
    TcpInit
    If ConnectToServer(1) Then
        Call SetStatus("Connected, sending login information...")
        Call SendLogin(frmMenu.txtLUser.Text, frmMenu.txtLPass.Text)
        Exit Sub
    End If

    If Not IsConnected Then
        Call SetStatus("All editors grouped for game developers")
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & GAME_WEBSITE, vbOKOnly, Options.Game_Name)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub logoutGame()
Dim Buffer As clsBuffer, I As Long

    isLogging = True
    InGame = False
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuit
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    Call DestroyTCP
    
    ' destroy the animations loaded
    For I = 1 To MAX_BYTE
        ClearAnimInstance (I)
    Next
    
    ' destroy temp values
    DragInvSlotNum = 0
    LastItemDesc = 0
    TempPlayer(MyIndex).SpellBuffer = 0
    TempPlayer(MyIndex).SpellBufferTimer = 0
    TempPlayer(MyIndex).SpellBufferNum = 0
    MyIndex = 0
    InventoryItemSelected = 0
    tmpCurrencyItem = 0
    
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    Unload frmAdminPanel
    Call SetStatus("All editors grouped for game developers")
    
    HideGame
End Sub

Sub GameInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EnteringGame = True
    frmMenu.visible = False
    EnteringGame = False
    
    ' hide gui
    InBank = False
    InShop = False
    InTrade = False
    
    ' Set font
    frmMain.Font = "Arial Bold"
    frmMain.FontSize = 10
    frmMain.Show
    
    ' get ping
    GetPing
    
    ' set values for amdin panel
    frmAdminPanel.scrlAItem.max = MAX_ITEMS
    frmAdminPanel.scrlAItem.Value = 1
    'stop the song playing
    StopMusic
    ShowGame
    chatShowLine = "|"
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyGame()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' break out of GameLoop
    InGame = False
    Call DestroyTCP
    HideGame
    
    'destroy objects in reverse order
    DestroyDX8
    
    DestroyFmod

    'Call UnloadAllForms
    End
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "destroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadAllForms()
Dim frm As Form

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For Each frm In VB.Forms
        Unload frm
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStatus(ByVal Caption As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'frmMenu.lblStatus.Caption = Caption
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If

    Txt.SelStart = Len(Txt.Text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Rand = Int((High - Low + 1) * Rnd) + Low
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent high ascii chars
    For I = 1 To Len(sInput)

        If Asc(Mid$(sInput, I, 1)) < vbKeySpace Or Asc(Mid$(sInput, I, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If

    Next

    isStringLegal = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Public Sub resetClickedButtons()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For I = 1 To MAX_BUTTONS
        Select Case I
            ' option buttons
            Case 26, 27, 28, 29, 30, 31, 32, 33, 55, 56
            Case 51, 52, 53, 54
                ' Nothing in here
            ' Everything else - reset
            Case Else
                ' reset state and render
                Buttons(I).State = 0 'normal
        End Select
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub PopulateLists()
Dim strLoad As String, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Cache music list
    strLoad = Dir(App.Path & MUSIC_PATH & "*.*")
    I = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To I) As String
        musicCache(I) = strLoad
        strLoad = Dir
        I = I + 1
    Loop
    
    ' Cache sound list
    strLoad = Dir(App.Path & SOUND_PATH & "*.*")
    I = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To I) As String
        soundCache(I) = strLoad
        strLoad = Dir
        I = I + 1
    Loop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShowGame()
Dim I As Long

    For I = 1 To 4
        GUIWindow(I).visible = True
    Next
End Sub

Public Sub HideGame()
Dim I As Long
    For I = 1 To GUI_Count - 1
        GUIWindow(I).visible = False
    Next
End Sub

Function MenuButtonName(ByVal MenuButton As Byte) As String

    Select Case MenuButton
        Case 1: MenuButtonName = "Bolsa"
        Case 2: MenuButtonName = "Skills"
        Case 3: MenuButtonName = "Personagem"
        Case 4: MenuButtonName = "Opções"
        Case 5: MenuButtonName = "Quests"
        Case 6: MenuButtonName = "Grupo"
    End Select
    
End Function

