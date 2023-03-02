Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' Text API
Declare Function GeneralWinDirApi Lib "kernel32" _
        Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
        ByVal nSize As Long) As Long
        
Public DrawMousePosition As Boolean
Public StateMenu As Byte
Public MsgScreen As String
Public InGameTick As Long
Public NewCharTick As Long
Public Flying As Long
Public Const NewCharClasse = 1
Public AutoLogin As Boolean
Public HaveDragonball(1 To 7) As Boolean
Public MoedasZ As Long

Public Sub Main()
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    ' load options
    Call SetStatus("Loading Options...")
    LoadOptions
    LoadCredits
    LoadLanguage
    
    'If App.LogMode = 1 Then
    '    If Command() = "" Then 'Se ele ja nao foi previamente executado
    '        Update
    '        If FileExist(App.Path & "\masterversion.ini", True) Then 'Se o arquivo existe
    '            If GetVar(App.Path & "\masterversion.ini", "MASTER", "Executável") <> "" Then 'Tem executável
    '                If App.EXEName <> GetVar(App.Path & "\masterversion.ini", "MASTER", "Executável") Then 'Não somos este executável
    '                    If FileExist(App.Path & "\" & GetVar(App.Path & "\masterversion.ini", "MASTER", "Executável") & ".exe", True) Then 'Se o executavel existe
    '                        'Shell App.Path & "\" & GetVar(App.Path & "\masterversion.ini", "MASTER", "Executável") & ".exe -NotRunAgain", vbNormalFocus 'Executar ele
    '                        'End
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End If
    'End If
    
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
    ChkDir App.Path & "\data files\graphics\gui\", "newmenu"
    ChkDir App.Path & "\data files\", "logs"
    ChkDir App.Path & "\data files\", "maps"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"
    
    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34) ' "
    
    ' Update the form with the game's name before it's loaded
    frmMain.Caption = Options.Game_Name
    
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
    'If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMusic Trim$(Options.MenuMusic)
    
    ' Reset values
    Ping = -1
    
    ' load main menu
    Call SetStatus("Loading Menu...")
    ' general menu stuff
    frmMenu.Caption = Options.Game_Name

    ' Load the username + pass
    frmMenu.txtLUser.Text = Trim$(Options.Username)
    If Options.savePass = 1 Then
        frmMenu.txtLPass.Text = Trim$(Options.Password)
        frmMenu.chkPass.value = Options.savePass
    End If
    
    ' cache the buttons then reset & render them
    Call SetStatus("Loading buttons...")
    cacheButtons
    resetButtons_Menu
    
    ' we can now see it
    'frmMenu.visible = True
    
    ' hide all pics
    'frmMenu.picCredits.visible = False
    'frmMenu.picCharacter.visible = False
    'frmMenu.picRegister.visible = False
    
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
    PaperdollOrder(1) = Equipment.Shield
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Armor
    PaperdollOrder(4) = Equipment.Weapon

    frmMain.Width = 12090
    frmMain.Height = 9420
    
    isLogging = True
    If Command() <> "silence" Then
        frmMain.visible = True
        If Command() <> "light" Then StartAntiHack
        MenuLoop
    Else
        frmMain.Width = 1
        frmMain.Height = 1
        frmAntiHack.Show
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub InitialiseGUI(Optional ByVal loadingScreen As Boolean = False)

'Loading Interface.ini data
Dim Filename As String
Filename = App.Path & "\data files\interface.ini"
Dim i As Long
    ' re-set chat scroll
    ChatScroll = 8
     ' menu
    'frmMenu.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\background.jpg")
    ReDim GUIWindow(1 To Gui_Count) As GUIWindowRec
    
    ' 1 - Chat
    With GUIWindow(GUI_CHAT)
        .X = Val(GetVar(Filename, "GUI_CHAT", "X"))
        .Y = Val(GetVar(Filename, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(Filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(Filename, "GUI_CHAT", "Height"))
        .visible = True
    End With
    
    ' 2 - Hotbar
    With GUIWindow(GUI_HOTBAR)
        .X = Val(GetVar(Filename, "GUI_HOTBAR", "X"))
        .Y = Val(GetVar(Filename, "GUI_HOTBAR", "Y"))
        .Height = Val(GetVar(Filename, "GUI_HOTBAR", "Height"))
        .Width = ((9 + 36) * (MAX_HOTBAR - 1))
    End With
    
    ' 3 - Menu
    With GUIWindow(GUI_MENU)
        .X = Val(GetVar(Filename, "GUI_MENU", "X"))
        .Y = Val(GetVar(Filename, "GUI_MENU", "Y"))
        .Width = Val(GetVar(Filename, "GUI_MENU", "Width"))
        .Height = Val(GetVar(Filename, "GUI_MENU", "Height"))
        .visible = True
    End With
    
    ' 4 - Bars
    With GUIWindow(GUI_BARS)
        .X = Val(GetVar(Filename, "GUI_BARS", "X"))
        .Y = Val(GetVar(Filename, "GUI_BARS", "Y"))
        .Width = Val(GetVar(Filename, "GUI_BARS", "Width"))
        .Height = Val(GetVar(Filename, "GUI_BARS", "Height"))
        .visible = True
    End With
    
    ' 5 - Inventory
    With GUIWindow(GUI_INVENTORY)
        .X = Val(GetVar(Filename, "GUI_INVENTORY", "X"))
        .Y = Val(GetVar(Filename, "GUI_INVENTORY", "Y"))
        .Width = Val(GetVar(Filename, "GUI_INVENTORY", "Width"))
        .Height = Val(GetVar(Filename, "GUI_INVENTORY", "Height"))
        .visible = False
    End With
    
    ' 6 - Spells
    With GUIWindow(GUI_SPELLS)
        .X = Val(GetVar(Filename, "GUI_SPELLS", "X"))
        .Y = Val(GetVar(Filename, "GUI_SPELLS", "Y"))
        .Width = Val(GetVar(Filename, "GUI_SPELLS", "Width"))
        .Height = Val(GetVar(Filename, "GUI_SPELLS", "Height"))
        .visible = False
    End With
    
    ' 7 - Character
    With GUIWindow(GUI_CHARACTER)
        .X = Val(GetVar(Filename, "GUI_CHARACTER", "X"))
        .Y = Val(GetVar(Filename, "GUI_CHARACTER", "Y"))
        .Width = Val(GetVar(Filename, "GUI_CHARACTER", "Width"))
        .Height = Val(GetVar(Filename, "GUI_CHARACTER", "Height"))
        .visible = False
    End With
    
    ' 8 - Options
    With GUIWindow(GUI_OPTIONS)
        .X = Val(GetVar(Filename, "GUI_OPTIONS", "X"))
        .Y = Val(GetVar(Filename, "GUI_OPTIONS", "Y"))
        .Width = Val(GetVar(Filename, "GUI_OPTIONS", "Width"))
        .Height = Val(GetVar(Filename, "GUI_OPTIONS", "Height"))
        .visible = False
    End With
    
    ' 9 - Party
    With GUIWindow(GUI_PARTY)
        .X = Val(GetVar(Filename, "GUI_PARTY", "X"))
        .Y = Val(GetVar(Filename, "GUI_PARTY", "Y"))
        .Width = Val(GetVar(Filename, "GUI_PARTY", "Width"))
        .Height = Val(GetVar(Filename, "GUI_PARTY", "Height"))
        .visible = False
    End With
    
    ' 10 - Description
    With GUIWindow(GUI_DESCRIPTION)
        .X = Val(GetVar(Filename, "GUI_DESCRIPTION", "X"))
        .Y = Val(GetVar(Filename, "GUI_DESCRIPTION", "Y"))
        .Width = Val(GetVar(Filename, "GUI_DESCRIPTION", "Width"))
        .Height = Val(GetVar(Filename, "GUI_DESCRIPTION", "Height"))
        .visible = False
    End With
    
    ' 11 - Main Menu
    With GUIWindow(GUI_MAINMENU)
        .X = Val(GetVar(Filename, "GUI_MAINMENU", "X"))
        .Y = Val(GetVar(Filename, "GUI_MAINMENU", "Y"))
        .Width = Val(GetVar(Filename, "GUI_MAINMENU", "Width"))
        .Height = Val(GetVar(Filename, "GUI_MAINMENU", "Height"))
        .visible = False
    End With
    
    ' 12 - Shop
    With GUIWindow(GUI_SHOP)
        .X = Val(GetVar(Filename, "GUI_SHOP", "X"))
        .Y = Val(GetVar(Filename, "GUI_SHOP", "Y"))
        .Width = Val(GetVar(Filename, "GUI_SHOP", "Width"))
        .Height = Val(GetVar(Filename, "GUI_SHOP", "Height"))
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
        .X = Val(GetVar(Filename, "GUI_CHAT", "X"))
        .Y = Val(GetVar(Filename, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(Filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(Filename, "GUI_CHAT", "Height"))
        .visible = False
    End With
    ' 16 - Dialogue
    With GUIWindow(GUI_DIALOGUE)
        .X = Val(GetVar(Filename, "GUI_CHAT", "X"))
        .Y = Val(GetVar(Filename, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(Filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(Filename, "GUI_CHAT", "Height"))
        .visible = False
    End With
    ' 17 - Event Chat
    With GUIWindow(GUI_EVENTCHAT)
        .X = Val(GetVar(Filename, "GUI_CHAT", "X"))
        .Y = Val(GetVar(Filename, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(Filename, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(Filename, "GUI_CHAT", "Height"))
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
    ' 19 - Death
    With GUIWindow(GUI_DEATH)
        .X = (frmMain.ScaleWidth / 2) - 160
        .Y = (frmMain.ScaleHeight / 2) - 160
        .Width = 320
        .Height = 320
        .visible = False
    End With
    ' 20 - Quests
    With GUIWindow(GUI_QUESTS)
        .X = 578
        .Y = 255
        .Width = 195
        .Height = 250
        .visible = False
    End With
    
    With GUIWindow(GUI_CONQUISTAS)
        .X = 150
        .Y = 150
        .Width = 500
        .Height = 300
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
        .Y = -47
        .Width = 36
        .Height = 36
        .visible = False
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
    For i = 16 To 20
        With Buttons(i)
            .State = 0 'normal
            .Width = 12
            .Height = 11
            .visible = True
            .PicNum = 13
        End With
    Next
    ' set the individual spaces
    For i = 16 To 20 ' first 3
        With Buttons(i)
            .X = 180
            .Y = 138 + 12 * (i - 16)
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
        .X = 80
        .Y = 276
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 16
    End With
    
    ' main - party invite
    With Buttons(24)
        .State = 0 ' normal
        .X = 0
        .Y = 203
        .Width = 71
        .Height = 47
        .visible = True
        .PicNum = 17
    End With
    
    ' main - party invite
    With Buttons(25)
        .State = 0 ' normal
        .X = 125
        .Y = 203
        .Width = 71
        .Height = 47
        .visible = True
        .PicNum = 18
    End With
    
    ' main - music on
    With Buttons(26)
        .State = 0 ' normal
        .X = 77
        .Y = 37
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - music off
    With Buttons(27)
        .State = 0 ' normal
        .X = 132
        .Y = 37
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - sound on
    With Buttons(28)
        .State = 0 ' normal
        .X = 77
        .Y = 62
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - sound off
    With Buttons(29)
        .State = 0 ' normal
        .X = 132
        .Y = 62
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - debug on
    With Buttons(30)
        .State = 0 ' normal
        .X = 77
        .Y = 85
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - debug off
    With Buttons(31)
        .State = 0 ' normal
        .X = 132
        .Y = 85
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
            .X = GUIWindow(GUI_TRADE).X + 165 - 32
            .Y = GUIWindow(GUI_TRADE).Y + 335
            .Width = 89
            .Height = 29
            .visible = True
            .PicNum = 25
        End With
    ' main - Decline Trade
        With Buttons(41)
            .State = 0 'normal
            .X = GUIWindow(GUI_TRADE).X + 245 + 32
            .Y = GUIWindow(GUI_TRADE).Y + 335
            .Width = 89
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
            .Y = 97
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 23
        End With
    ' main - Volume Right
        With Buttons(45)
            .State = 0 'normal
            .X = 147
            .Y = 97
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 24
        End With
    ' main - ambiente on
        With Buttons(46)
            .State = 0 ' normal
            .X = 77
            .Y = 141
            .Width = 49
            .Height = 19
            .visible = True
            .PicNum = 19
        End With
        
    ' main - ambiente off
        With Buttons(47)
            .State = 0 ' normal
            .X = 132
            .Y = 141
            .Width = 49
            .Height = 19
            .visible = True
            .PicNum = 20
        End With
        
    ' main - cor da tela on
        With Buttons(48)
            .State = 0 ' normal
            .X = 77
            .Y = 164
            .Width = 49
            .Height = 19
            .visible = True
            .PicNum = 19
        End With
        
    ' main - cor da tela off
        With Buttons(49)
            .State = 0 ' normal
            .X = 132
            .Y = 164
            .Width = 49
            .Height = 19
            .visible = True
            .PicNum = 20
        End With
        
    ' main - clima on
        With Buttons(50)
            .State = 0 ' normal
            .X = 77
            .Y = 187
            .Width = 49
            .Height = 19
            .visible = True
            .PicNum = 19
        End With
        
    ' main - clima off
        With Buttons(51)
            .State = 0 ' normal
            .X = 132
            .Y = 187
            .Width = 49
            .Height = 19
            .visible = True
            .PicNum = 20
        End With
        
    ' main - neblina on
        With Buttons(52)
            .State = 0 ' normal
            .X = 77
            .Y = 210
            .Width = 49
            .Height = 19
            .visible = True
            .PicNum = 19
        End With
        
    ' main - neblina off
        With Buttons(53)
            .State = 0 ' normal
            .X = 132
            .Y = 210
            .Width = 49
            .Height = 19
            .visible = True
            .PicNum = 20
        End With
        
    ' Death Window - normal
        With Buttons(54)
            .State = 0 ' normal
            .X = 96
            .Y = 169
            .Width = 107
            .Height = 24
            .visible = False
            .PicNum = 27
        End With
        
        ' Remover
        With Buttons(55)
            .State = 0 ' normal
            .X = 0
            .Y = 0
            .Width = 89
            .Height = 29
            .visible = False
            .PicNum = 29
        End With
        
        ' Acelerar
        With Buttons(56)
            .State = 0 ' normal
            .X = 0
            .Y = 0
            .Width = 89
            .Height = 29
            .visible = False
            .PicNum = 35
        End With
        
        ' Mover
        With Buttons(57)
            .State = 0 ' normal
            .X = 0
            .Y = 0
            .Width = 89
            .Height = 29
            .visible = False
            .PicNum = 30
        End With
        
        ' Mover
        With Buttons(58)
            .State = 0 ' normal
            .X = 0
            .Y = 0
            .Width = 89
            .Height = 29
            .visible = False
            .PicNum = 31
        End With
        
        ' Abrir
        With Buttons(59)
            .State = 0 ' normal
            .X = 0
            .Y = 0
            .Width = 89
            .Height = 29
            .visible = False
            .PicNum = 32
        End With
        
        ' Acelerar
        With Buttons(60)
            .State = 0 ' normal
            .X = 0
            .Y = 0
            .Width = 89
            .Height = 29
            .visible = False
            .PicNum = 36
        End With
        
        ' Mais
        With Buttons(61)
            .State = 0 ' normal
            .X = 321
            .Y = 11
            .Width = 36
            .Height = 36
            .visible = True
            .PicNum = 37
        End With
        
        ' Conquistas
        With Buttons(62)
            .State = 0 ' normal
            .X = 321
            .Y = -94
            .Width = 36
            .Height = 36
            .visible = False
            .PicNum = 38
        End With
        
        ' Prox - Conquistas
        With Buttons(63)
            .State = 0 ' normal
            .X = 619
            .Y = 422
            .Width = 19
            .Height = 19
            .visible = False
            .PicNum = 24
        End With
        
        ' Ant - Conquistas
        With Buttons(64)
            .State = 0 ' normal
            .X = 158
            .Y = 422
            .Width = 19
            .Height = 19
            .visible = False
            .PicNum = 23
        End With
        
        InitialiseNewGui
End Sub
Public Sub MenuState(ByVal State As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Select Case State
        Case MENU_STATE_ADDCHAR
            frmMenu.picCredits.visible = False
            frmMenu.picLogin.visible = False
            frmMenu.picCharacter.visible = False
            frmMenu.picRegister.visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending character addition data...")

                If frmMenu.optMale.value Then
                    Call SendAddChar(frmMenu.txtCName, SEX_MALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite, newCharHair)
                Else
                    Call SendAddChar(frmMenu.txtCName, SEX_FEMALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite, newCharHair)
                End If
            End If
            
        Case MENU_STATE_NEWACCOUNT
            frmMenu.picCredits.visible = False
            frmMenu.picLogin.visible = False
            frmMenu.picCharacter.visible = False
            frmMenu.picRegister.visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmMenu.txtRUser.Text, frmMenu.txtRPass.Text)
            End If

        Case MENU_STATE_LOGIN
            frmMenu.picCredits.visible = False
            frmMenu.picLogin.visible = False
            frmMenu.picCharacter.visible = False
            frmMenu.picRegister.visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmMenu.txtLUser.Text, frmMenu.txtLPass.Text)
                Exit Sub
            End If
    End Select

        If Not IsConnected Then
            frmMenu.picCredits.visible = False
            frmMenu.picCharacter.visible = False
            frmMenu.picRegister.visible = False
            frmMenu.picLogin.visible = True
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
Dim buffer As clsBuffer, i As Long
    
    If isLogging = False Then
        isLogging = True
        InGame = False
        Set buffer = New clsBuffer
        buffer.WriteLong CQuit
        SendData buffer.ToArray()
        Set buffer = Nothing
        Call DestroyTCP
        
        ' destroy the animations loaded
        For i = 1 To MAX_BYTE
            ClearAnimInstance (i)
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
        End
    Else
        End
    End If
End Sub

Sub GameInit()
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

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
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
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
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' Prevent high ascii chars
    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
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

' ####################
' ## Buttons - Menu ##
' ####################
Public Sub cacheButtons()
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' menu - login
    With MenuButton(1)
        .Filename = "login"
        .State = 0 ' normal
    End With
    
    ' menu - register
    With MenuButton(2)
        .Filename = "register"
        .State = 0 ' normal
    End With
    
    ' menu - credits
    With MenuButton(3)
        .Filename = "credits"
        .State = 0 ' normal
    End With
    
    ' menu - exit
    With MenuButton(4)
        .Filename = "exit"
        .State = 0 ' normal
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cacheButtons", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub resetClickedButtons()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_BUTTONS
        Select Case i
            ' option buttons
            Case 26, 27, 28, 29, 30, 31, 32, 33, 46, 47, 48, 49, 50, 55, 56
            Case 51, 52, 53, 54
                ' Nothing in here
            ' Everything else - reset
            Case Else
                ' reset state and render
                Buttons(i).State = 0 'normal
        End Select
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


' menu specific buttons
Public Sub resetButtons_Menu(Optional ByVal exceptionNum As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_MENUBUTTONS
        ' only change if different and not exception
        If Not MenuButton(i).State = 0 And Not i = exceptionNum Then
            ' reset state and render
            MenuButton(i).State = 0 'normal
            renderButton_Menu i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Menu = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub renderButton_Menu(ByVal ButtonNum As Long)
Dim bSuffix As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' get the suffix
    Select Case MenuButton(ButtonNum).State
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMenu.imgButton(ButtonNum).Picture = LoadPicture(App.Path & MENUBUTTON_PATH & MenuButton(ButtonNum).Filename & bSuffix & ".jpg")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "renderButton_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub changeButtonState_Menu(ByVal ButtonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MenuButton(ButtonNum).State = bState Then Exit Sub
        ' change and render
        MenuButton(ButtonNum).State = bState
        renderButton_Menu ButtonNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "changeButtonState_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub PopulateLists()
Dim strLoad As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' Cache music list
    strLoad = Dir(App.Path & MUSIC_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop
    
    ' Cache sound list
    strLoad = Dir(App.Path & SOUND_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShowGame()
Dim i As Long

    For i = 1 To 4
        GUIWindow(i).visible = True
    Next
End Sub

Public Sub HideGame()
Dim i As Long
    For i = 1 To Gui_Count - 1
        GUIWindow(i).visible = False
    Next
End Sub

Function IsClick(GuiNum As Long) As Boolean
    IsClick = False
    If GlobalX > NewGUIWindow(GuiNum).X And GlobalX <= NewGUIWindow(GuiNum).X + NewGUIWindow(GuiNum).Width Then
        If GlobalY > NewGUIWindow(GuiNum).Y And GlobalY <= NewGUIWindow(GuiNum).Y + NewGUIWindow(GuiNum).Height Then
            IsClick = True
        End If
    End If
End Function
Sub CheckNewGuiMove()
    
    NewGUIWindow(LOGINBUTTON).visible = IsClick(LOGINBUTTON)
    NewGUIWindow(COLORBUTTON).visible = IsClick(COLORBUTTON)
    NewGUIWindow(HAIRBUTTON).visible = IsClick(HAIRBUTTON)
    NewGUIWindow(CREATEBUTTON).visible = IsClick(CREATEBUTTON)
    
End Sub

Sub InitialiseNewGui()
    ReDim NewGUIWindow(1 To NewGUIWindows.NewGui_Count - 1) As NewGUIWindowRec
    
    ' Login Text
    With NewGUIWindow(TEXTLOGIN)
        .X = 166
        .Y = 34
        .Width = 150
        .Height = 25
        .visible = False
        .value = ""
    End With
    
    ' Password Text
    With NewGUIWindow(TEXTPASSWORD)
        .X = 423
        .Y = 34
        .Width = 150
        .Height = 25
        .visible = False
        .value = ""
    End With
    
    ' Login button
    With NewGUIWindow(LOGINBUTTON)
        .X = 614
        .Y = 30
        .Width = 82
        .Height = 30
        .visible = False
        .value = ""
    End With
    
    ' Nome do char
    With NewGUIWindow(TEXTCHARNAME)
        .X = 324
        .Y = 214
        .Width = 168
        .Height = 17
        .visible = False
        .value = ""
    End With
    
    ' Mudar pele
    With NewGUIWindow(COLORBUTTON)
        .X = 343
        .Y = 248
        .Width = 132
        .Height = 30
        .visible = False
        .value = ""
    End With
    
    ' Mudar Cabelo
    With NewGUIWindow(HAIRBUTTON)
        .X = 364
        .Y = 288
        .Width = 87
        .Height = 30
        .visible = False
        .value = ""
    End With
    
    ' Criar
    With NewGUIWindow(CREATEBUTTON)
        .X = 367
        .Y = 356
        .Width = 82
        .Height = 30
        .visible = False
        .value = ""
    End With
    
End Sub

Sub CheckNewGui()
    
    If NewCharTick = 0 Then
        'TextLogin
        NewGUIWindow(TEXTLOGIN).visible = IsClick(TEXTLOGIN)
        
        'TextPassword
        NewGUIWindow(TEXTPASSWORD).visible = IsClick(TEXTPASSWORD)
        
        If Not InGame Then
        If GlobalX > (frmMain.ScaleWidth / 2) - 150 And GlobalX < (frmMain.ScaleWidth / 2) + 150 Then
            If GlobalY > (frmMain.ScaleHeight / 2) + 140 And GlobalY < (frmMain.ScaleHeight / 2) + 190 Then
                Dim r As Long
                r = ShellExecute(0, "open", GetVar(App.Path & "\masterversion.ini", "MASTER", "PatchNotes"), 0, 0, 1)
            End If
        End If
        End If
        
        If IsClick(LOGINBUTTON) Then
            If Len(NewGUIWindow(TEXTLOGIN).value) <= 3 Then
                MsgScreen = "O login deve ter no minimo 3 caracteres!"
                Exit Sub
            End If
            
            If Len(NewGUIWindow(TEXTPASSWORD).value) <= 3 Then
                MsgScreen = "A senha deve ter no minimo 3 caracteres!"
                Exit Sub
            End If
            If MsgScreen <> "Conectado!, aguardando informações de login..." And InGameTick = 0 Then
                If ConnectToServer(1) Then
                    MsgScreen = "Conectado!, aguardando informações de login..."
                    Call SendLogin(NewGUIWindow(TEXTLOGIN).value, NewGUIWindow(TEXTPASSWORD).value)
                Else
                    MsgScreen = "Não foi possível conectar ao servidor, ele pode estar em manutenção ou você está sem conexão com a internet"
                End If
            End If
        End If
    Else
        'Textcharname
        NewGUIWindow(TEXTCHARNAME).visible = IsClick(TEXTCHARNAME)
        
        If IsClick(COLORBUTTON) Then
            Dim spritecount As Long
            
            spritecount = UBound(Class(NewCharClasse).MaleSprite)
        
            If newCharSprite >= spritecount Then
                newCharSprite = 0
            Else
                newCharSprite = newCharSprite + 1
            End If
        End If
        
        If IsClick(HAIRBUTTON) Then
            If newCharHair + 1 <= NumHair(0) Then
                newCharHair = newCharHair + 1
            Else
                newCharHair = 1
            End If
        End If
        
        If IsClick(CREATEBUTTON) Then
            Call SendAddChar(NewGUIWindow(TEXTCHARNAME).value, SEX_MALE, NewCharClasse, newCharSprite, newCharHair)
        End If
    End If
End Sub

Function MenuButtonName(ByVal MenuButton As Byte) As String

    Select Case MenuButton
        Case 1: MenuButtonName = printf("Bolsa")
        Case 2: MenuButtonName = printf("Skills")
        Case 3: MenuButtonName = printf("Personagem")
        Case 4: MenuButtonName = printf("Opções")
        Case 5: MenuButtonName = printf("Quests")
        Case 6: MenuButtonName = printf("Grupo")
        Case 61: MenuButtonName = printf("Mais...")
        Case 62: MenuButtonName = printf("Conquistas")
    End Select
    
End Function
