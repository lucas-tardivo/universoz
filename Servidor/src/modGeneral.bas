Attribute VB_Name = "modGeneral"
Option Explicit
' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Public Sub Main()
    Call InitServer
End Sub

Public Sub InitServer()
    Dim i As Long
    Dim F As Long
    Dim time1 As Long
    Dim time2 As Long
    
    Call InitMessages
    time1 = GetTickCount
    EventGlobalTick = GetTickCount
    frmServer.Show
    ' Initialize the random-number generator
    Randomize ', seed
    
    If UZ Then
        If Not ConnectDatabase Then
            MsgBox "Não foi possível conectar ao banco de dados (MYSQL)"
            End
        End If
    Else
        frmServer.lstNiveis.AddItem ("Sistema espacial desativado")
    End If

    ' Check if the directory is there, if its not make it
    ChkDir App.path & "\Data\", "accounts"
    ChkDir App.path & "\Data\", "animations"
    ChkDir App.path & "\Data\", "banks"
    ChkDir App.path & "\Data\", "items"
    ChkDir App.path & "\Data\", "logs"
    ChkDir App.path & "\Data\", "maps"
    ChkDir App.path & "\Data\", "npcs"
    ChkDir App.path & "\Data\", "resources"
    ChkDir App.path & "\Data\", "shops"
    ChkDir App.path & "\Data\", "spells"
    ChkDir App.path & "\Data\", "quests"
    ChkDir App.path & "\Data\", "events"
    ChkDir App.path & "\Data\", "effects"
    ChkDir App.path & "\Data\", "guilds"
    ChkDir App.path & "\data\logs\", "global"
    ChkDir App.path & "\data\logs\", "map"
    ChkDir App.path & "\data\logs\", "emote"
    ChkDir App.path & "\data\logs\", "player"
    ChkDir App.path & "\data\logs\", "system"

    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' load options, set if they dont exist
    If Not FileExist(App.path & "\data\options.ini", True) Then
        Options.Game_Name = "Universo Z"
        Options.Port = 7001
        Options.MOTD = "Welcome."
        Options.Website = ""
        Options.EventChance = 60
        SaveOptions
    Else
        LoadOptions
    End If
    
    MAX_PLAYERS = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "Players"))
    
    Call DoRedims
    UsersOnline_Start
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    
    frmServer.txtMOTD.Text = Trim$(Options.MOTD)
    frmServer.chkLogs.Value = Options.Logs
    
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call CheckEXP
    Call ClearGameData
    Call LoadGameData
    If UZ Then
        For i = 1 To MAX_ITEMS
            If Item(i).Type = ItemType.ITEM_TYPE_PLANETCHANGE And Item(i).data1 = 9 Then
                Relogio = i
                Exit For
            End If
        Next i
    
        Call SetStatus("Creating planets...")
        StartPlanets
    End If
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache
    Call SetStatus("Loading System Tray...")
    Call LoadSystemTray
    Call SetStatus("Loading Experience...")
    Call LoadLanguage
    Call LoadWishes

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If

    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    time2 = GetTickCount
    Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & "ms.")
    frmServer.SSTab1.Enabled = True
    frmServer.SSTab2.Enabled = True
    
    ' reset shutdown value
    isShuttingDown = False
    frmServer.sckWeb.Listen
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
    Dim i As Long
    ServerOnline = False
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")
    
    For i = 1 To GetMaxPlayerPlanets
        SavePlayerPlanet i
    Next i

    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next
    
    Call SetStatus("Saving options...")
    SaveOptions
    Call SetStatus("Saving Esp. Banks...")
    SaveEspAmount

    If frmServer.sckWeb.State = sckConnected Then
        Call SetStatus("Sending WebManager close signal")
        CloseWebManager
    End If

    End
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status, ChatSystem)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing Resources...")
    Call ClearResources
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing events...")
    Call ClearEvents
    Call SetStatus("Clearing effects...")
    Call ClearEffects
    Call SetStatus("Clearing quests...")
    Call Clearquests
End Sub

Private Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading Resources...")
    Call LoadResources
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    Call SetStatus("Loading switches...")
    Call LoadSwitches
    Call SetStatus("Loading variables...")
    Call LoadVariables
    Call SetStatus("Loading events...")
    Call LoadEvents
    Call SetStatus("Loading effects...")
    Call LoadEffects
    Call SetStatus("Loading quests...")
    Call LoadQuests
    Call SetStatus("Loading guilds...")
    Call LoadGuilds
    Call SetStatus("Loading event rewards...")
    Call LoadEventsRewards
    Call SetStatus("Loading houses...")
    Call LoadHouses
    Call SetStatus("Checking transportes...")
    Call CheckTransportes
    Call SetStatus("Checking provações...")
    Call CheckProvacoes
    Call SetStatus("Loading NPC Base...")
    Call LoadNPCBase
    Call SetStatus("Loading Daily missions...")
    Call LoadDailyMission
    Call SetStatus("Loading Conquistas...")
    Call LoadConquistas
    Call SetStatus("Loading Esp. Banks...")
    Call LoadEspAmount
End Sub

Public Sub TextAdd(Msg As String, ByVal TextType As Byte)
    On Error Resume Next
    NumLines(TextType) = NumLines(TextType) + 1
    
    If TextType = ChatSystem Then
        SaveChatLine TextType, Msg
    End If
    
    If NumLines(TextType) >= MAX_LINES Then
        If TextType <> ChatSystem Then
            SaveChatLog TextType
        End If
        frmServer.txtText(TextType).Text = vbNullString
        NumLines(TextType) = 0
    End If
    
    frmServer.txtText(TextType).Text = frmServer.txtText(TextType).Text & vbNewLine & Msg
    frmServer.txtText(TextType).SelStart = Len(frmServer.txtText(TextType).Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function

Public Function KeepTwoDigit(Num As Long)
    If (Num < 10) Then
        KeepTwoDigit = "0" & Num
    Else
        KeepTwoDigit = Num
    End If
End Function

Public Sub SwitchValue(ByRef Var1 As Long, ByRef Var2 As Long)
    Dim temp As Long
    temp = Var1
    Var1 = Var2
    Var2 = temp
End Sub

Sub UsersOnline_Start()
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (i)

        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If

        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(4) = vbNullString
    Next

End Sub

Public Function Semana() As Integer
  Semana = Format(Now, "w") - 1
End Function

Sub SetupBonuses()
    If Semana() >= 1 And Semana() <= 2 Then
        Options.DropFactor = 2
    End If
    If Semana() >= 3 And Semana() <= 4 Then
        Options.ResourceFactor = 1.25
    End If
    If Semana() >= 5 And Semana() <= 6 Then
        Options.ExpFactor = 1.25
    End If
    If Semana() = 0 Then
        Options.GoldFactor = 1.5
    End If
End Sub
