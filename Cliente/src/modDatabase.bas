Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public DllHand As String

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim Filename As String
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Filename = App.Path & "\data files\logs\errors.txt"
    Open Filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleDLL(ByVal DllName As String)
Dim Filename As String
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Filename = App.Path & "\data files\logs\errors.txt"
    Open Filename For Append As #1
        Print #1, "DLL Found: " & DllName
    Close #1
    DllHand = DllHand & "@" & DllName
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal Filename As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & Filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(Filename)) > 0 Then
            FileExist = True
        End If
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, value As String)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Call WritePrivateProfileString$(Header, Var, value, File)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
Dim Filename As String

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Filename = App.Path & "\Data Files\config.ini"
    
    Call PutVar(Filename, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(Filename, "Options", "Username", Trim$(Options.Username))
    Call PutVar(Filename, "Options", "Password", Trim$(Options.Password))
    Call PutVar(Filename, "Options", "SavePass", str(Options.savePass))
    Call PutVar(Filename, "Options", "IP", Options.IP)
    Call PutVar(Filename, "Options", "Port", str(Options.Port))
    Call PutVar(Filename, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(Filename, "Options", "Music", str(Options.Music))
    Call PutVar(Filename, "Options", "Sound", str(Options.Sound))
    Call PutVar(Filename, "Options", "Debug", str(Options.Debug))
    Call PutVar(Filename, "Options", "Volume", str(Options.volume))
    Call PutVar(Filename, "Options", "FPSCap", str(Options.FPS))
    Call PutVar(Filename, "Options", "Ambiente", str(Options.Ambiente))
    Call PutVar(Filename, "Options", "Clima", str(Options.Clima))
    Call PutVar(Filename, "Options", "Tela", str(Options.Tela))
    Call PutVar(Filename, "Options", "Neblina", str(Options.Neblina))
    Call PutVar(Filename, "Options", "Pick", str(Options.PickMenu))
    
    If Options.Window = 1 Then
        Call PutVar(Filename, "Options", "Window", "1")
    Else
        Call PutVar(Filename, "Options", "Window", "1")
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim Filename As String

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Filename = App.Path & "\Data Files\config.ini"
    
    If Not FileExist(Filename, True) Then
        Options.Game_Name = "DBZ Online"
        Options.Password = vbNullString
        Options.savePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.Port = 7001
        Options.MenuMusic = vbNullString
        Options.Music = 1
        Options.Sound = 1
        Options.Debug = 0
        Options.volume = 150
        Options.FPS = 20
        Options.Ambiente = 1
        Options.Clima = 1
        Options.Tela = 1
        Options.Neblina = 1
        Options.Window = 1
        Options.Language = "pt"
        SaveOptions
    Else
        Options.Game_Name = GetVar(Filename, "Options", "Game_Name")
        Options.Username = GetVar(Filename, "Options", "Username")
        Options.Password = GetVar(Filename, "Options", "Password")
        Options.savePass = Val(GetVar(Filename, "Options", "SavePass"))
        Options.IP = GetVar(Filename, "MASTERSERVER", "ServerIP1")
        Options.Port = Val(GetVar(Filename, "Options", "Port"))
        Options.MenuMusic = GetVar(Filename, "Options", "MenuMusic")
        Options.Music = GetVar(Filename, "Options", "Music")
        Options.Sound = GetVar(Filename, "Options", "Sound")
        Options.Debug = GetVar(Filename, "Options", "Debug")
        Options.volume = GetVar(Filename, "Options", "Volume")
        Options.FPS = GetVar(Filename, "Options", "FPSCap")
        Options.Ambiente = Val(GetVar(Filename, "Options", "Ambiente"))
        Options.Clima = Val(GetVar(Filename, "Options", "Clima"))
        Options.Tela = Val(GetVar(Filename, "Options", "Tela"))
        Options.Neblina = Val(GetVar(Filename, "Options", "Neblina"))
        Options.Language = GetVar(Filename, "Options", "Lang")
        Options.PickMenu = Val(GetVar(Filename, "Options", "Pick"))
        If Options.Language = "" Then Options.Language = "pt"
        
        Dim i As Long
        i = 1
        Do While GetVar(Filename, "MASTERSERVER", "Server" & i) <> ""
            i = i + 1
        Loop
        i = i - 1
        
        ReDim Options.Servers(1 To i) As ServerRec
        For i = 1 To UBound(Options.Servers)
            Options.Servers(i).name = GetVar(Filename, "MASTERSERVER", "Server" & i)
            Options.Servers(i).IP = GetVar(Filename, "MASTERSERVER", "ServerIP" & i)
        Next i
        
        If App.LogMode = 0 Then
            ReDim Preserve Options.Servers(1 To i) As ServerRec
            Options.Servers(i).name = "Servidor de teste local"
            Options.Servers(i).IP = "localhost"
        End If
        
        SelectedServer = 1
        
        If Val(GetVar(Filename, "Options", "Window")) = 1 Then
            Options.Window = 1
        Else
            Options.Window = 0
        End If
        
        
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub LoadCredits()
Dim Filename As String

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    frmMenu.lblCredits.Caption = "Robin: Eclipse Origins" & vbNewLine & "Boasfesta: Programador" & vbNewLine & "Pirata_254: In-Game, Gráficos e roteiro" & vbNewLine & "Neeto: Mapas" & vbNewLine & "Enterbrain: Gráficos" & vbNewLine & "www.shockwave-sound.com: BGM" & vbNewLine & "Agradecimentos: Fryeja, MMODEV"
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadCredits", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal mapnum As Long)
Dim Filename As String
Dim f As Long
Dim X As Long
Dim Y As Long, i As Long, z As Long, w As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Filename = App.Path & MAP_PATH & "map" & mapnum & MAP_EXT

    f = FreeFile
    Open Filename For Binary As #f
    Put #f, , Map.name
    Put #f, , Map.Music
    Put #f, , Map.BGS
    Put #f, , Map.Revision
    Put #f, , Map.Moral
    Put #f, , Map.Up
    Put #f, , Map.Down
    Put #f, , Map.Left
    Put #f, , Map.Right
    Put #f, , Map.BootMap
    Put #f, , Map.BootX
    Put #f, , Map.BootY
    
    Put #f, , Map.Weather
    Put #f, , Map.WeatherIntensity
    
    Put #f, , Map.Fog
    Put #f, , Map.FogSpeed
    Put #f, , Map.FogOpacity
    
    Put #f, , Map.Red
    Put #f, , Map.Green
    Put #f, , Map.Blue
    Put #f, , Map.Alpha
    
    Put #f, , Map.MaxX
    Put #f, , Map.MaxY

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Put #f, , Map.Tile(X, Y)
        Next

        DoEvents
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #f, , Map.Npc(X)
        Put #f, , Map.NpcSpawnType(X)
    Next
    
    Put #f, , Map.Panorama
    Put #f, , Map.Fly
    Put #f, , Map.Ambiente
    Put #f, , Map.FogDir

    Close #f
    
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal mapnum As Long)
Dim Filename As String
Dim f As Long
Dim X As Long
Dim Y As Long, i As Long, z As Long, w As Long, p As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Filename = App.Path & MAP_PATH & "map" & mapnum & MAP_EXT
    ClearMap
    f = FreeFile
    Open Filename For Binary As #f
    Get #f, , Map.name
    Get #f, , Map.Music
    Get #f, , Map.BGS
    Get #f, , Map.Revision
    Get #f, , Map.Moral
    Get #f, , Map.Up
    Get #f, , Map.Down
    Get #f, , Map.Left
    Get #f, , Map.Right
    Get #f, , Map.BootMap
    Get #f, , Map.BootX
    Get #f, , Map.BootY
    
    Get #f, , Map.Weather
    Get #f, , Map.WeatherIntensity
        
    Get #f, , Map.Fog
    Get #f, , Map.FogSpeed
    Get #f, , Map.FogOpacity
        
    Get #f, , Map.Red
    Get #f, , Map.Green
    Get #f, , Map.Blue
    Get #f, , Map.Alpha
    
    Get #f, , Map.MaxX
    Get #f, , Map.MaxY
    ' have to set the tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Get #f, , Map.Tile(X, Y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Get #f, , Map.Npc(X)
        Get #f, , Map.NpcSpawnType(X)
    Next
    
    Get #f, , Map.Panorama
    
    Get #f, , Map.Fly
    Get #f, , Map.Ambiente
    Get #f, , Map.FogDir
    Close #f
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTilesets()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumTileSets = 1
    
    ReDim Tex_Tileset(1)

    While FileExist(GFX_PATH & "tilesets\" & i & GFX_EXT) Or FileExist(GFX_PATH & "tilesets\" & i & ".png")
        ReDim Preserve Tex_Tileset(NumTileSets)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        If FileExist(GFX_PATH & "tilesets\" & i & GFX_EXT) Then
            Tex_Tileset(NumTileSets).filepath = App.Path & GFX_PATH & "tilesets\" & i & GFX_EXT
        Else
            Tex_Tileset(NumTileSets).filepath = App.Path & GFX_PATH & "tilesets\" & i & ".png"
        End If
        Tex_Tileset(NumTileSets).Texture = NumTextures
        NumTileSets = NumTileSets + 1
        i = i + 1
    Wend
    
    NumTileSets = NumTileSets - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckCharacters()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumCharacters = 1
    
    ReDim Tex_Character(1)
    

    While FileExist(GFX_PATH & "characters\" & i & GFX_EXT)
        ReDim Preserve Tex_Character(NumCharacters)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Character(NumCharacters).filepath = App.Path & GFX_PATH & "characters\" & i & GFX_EXT
        Tex_Character(NumCharacters).Texture = NumTextures
        Tex_Character(NumCharacters).Transparency = True
        NumCharacters = NumCharacters + 1
        i = i + 1
    Wend
    
    NumCharacters = NumCharacters - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTutorials()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumTutoriais = 1
    
    ReDim Tex_Tutorial(1)
    

    While FileExist(GFX_PATH & "gui\tutorial\" & i & ".png")
        ReDim Preserve Tex_Tutorial(NumTutoriais)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Tutorial(NumTutoriais).filepath = App.Path & GFX_PATH & "gui\tutorial\" & i & ".png" 'GFX_EXT
        Tex_Tutorial(NumTutoriais).Texture = NumTextures
        Tex_Tutorial(NumTutoriais).Transparency = True
        NumTutoriais = NumTutoriais + 1
        i = i + 1
    Wend
    
    NumTutoriais = NumTutoriais - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckTutorials", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPaperdolls()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumPaperdolls = 1
    
    ReDim Tex_Paperdoll(1)

    While FileExist(GFX_PATH & "paperdolls\" & i & GFX_EXT)
        ReDim Preserve Tex_Paperdoll(NumPaperdolls)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Paperdoll(NumPaperdolls).filepath = App.Path & GFX_PATH & "paperdolls\" & i & GFX_EXT
        Tex_Paperdoll(NumPaperdolls).Texture = NumTextures
        NumPaperdolls = NumPaperdolls + 1
        i = i + 1
    Wend
    
    NumPaperdolls = NumPaperdolls - 1
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumAnimations = 1
    
    ReDim Tex_Animation(1)
    ReDim AnimationTimer(1)

    While FileExist(GFX_PATH & "animations\" & i & ".png") Or FileExist(GFX_PATH & "animations\" & i & GFX_EXT)
        ReDim Preserve Tex_Animation(NumAnimations)
        ReDim Preserve AnimationTimer(NumAnimations)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Animation(NumAnimations).Texture = NumTextures
        If FileExist(GFX_PATH & "animations\" & i & ".png") Then
            Tex_Animation(NumAnimations).filepath = App.Path & GFX_PATH & "animations\" & i & ".png"
        Else
            Tex_Animation(NumAnimations).filepath = App.Path & GFX_PATH & "animations\" & i & GFX_EXT
        End If
        NumAnimations = NumAnimations + 1
        i = i + 1
    Wend
    
    NumAnimations = NumAnimations - 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    numitems = 1
    
    ReDim Tex_Item(1)

    While FileExist(GFX_PATH & "items\" & i & GFX_EXT)
        ReDim Preserve Tex_Item(numitems)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Item(numitems).filepath = App.Path & GFX_PATH & "items\" & i & GFX_EXT
        Tex_Item(numitems).Texture = NumTextures
        Tex_Item(numitems).Transparency = True
        numitems = numitems + 1
        i = i + 1
    Wend
    
    numitems = numitems - 1

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumResources = 1
    
    ReDim Tex_Resource(1)

    While FileExist(GFX_PATH & "resources\" & i & GFX_EXT) Or FileExist(GFX_PATH & "resources\" & i & ".png")
        ReDim Preserve Tex_Resource(NumResources)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        If FileExist(GFX_PATH & "resources\" & i & GFX_EXT) Then
            Tex_Resource(NumResources).filepath = App.Path & GFX_PATH & "resources\" & i & GFX_EXT
        Else
            Tex_Resource(NumResources).filepath = App.Path & GFX_PATH & "resources\" & i & ".png"
        End If
        Tex_Resource(NumResources).Texture = NumTextures
        NumResources = NumResources + 1
        i = i + 1
    Wend
    
    NumResources = NumResources - 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckSpellIcons()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumSpellIcons = 1
    
    ReDim Tex_SpellIcon(1)

    While FileExist(GFX_PATH & "spellicons\" & i & GFX_EXT)
        ReDim Preserve Tex_SpellIcon(NumSpellIcons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_SpellIcon(NumSpellIcons).filepath = App.Path & GFX_PATH & "spellicons\" & i & GFX_EXT
        Tex_SpellIcon(NumSpellIcons).Texture = NumTextures
        Tex_SpellIcon(NumSpellIcons).Transparency = True
        NumSpellIcons = NumSpellIcons + 1
        i = i + 1
    Wend
    
    NumSpellIcons = NumSpellIcons - 1

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFaces()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumFaces = 1
    
    ReDim Tex_Face(1)

    While FileExist(GFX_PATH & "Faces\" & i & GFX_EXT)
        ReDim Preserve Tex_Face(NumFaces)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Face(NumFaces).filepath = App.Path & GFX_PATH & "faces\" & i & GFX_EXT
        Tex_Face(NumFaces).Texture = NumTextures
        NumFaces = NumFaces + 1
        i = i + 1
    Wend
    
    NumFaces = NumFaces - 1

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFogs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumFogs = 1
    
    ReDim Tex_Fog(1)
    While FileExist(GFX_PATH & "fogs\" & i & GFX_EXT)
        ReDim Preserve Tex_Fog(NumFogs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Fog(NumFogs).filepath = App.Path & GFX_PATH & "fogs\" & i & GFX_EXT
        Tex_Fog(NumFogs).Texture = NumTextures
        NumFogs = NumFogs + 1
        i = i + 1
    Wend
    
    NumFogs = NumFogs - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFogs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub CheckGUIs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumGUIs = 1
    
    ReDim Tex_GUI(1)
    While FileExist(GFX_PATH & "gui\" & i & ".png")
        ReDim Preserve Tex_GUI(NumGUIs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_GUI(NumGUIs).filepath = App.Path & GFX_PATH & "gui\" & i & ".png"
        Tex_GUI(NumGUIs).Texture = NumTextures
        NumGUIs = NumGUIs + 1
        i = i + 1
    Wend
    
    i = 1
    ReDim Tex_NewGUI(1)
    NumNewGUIs = 1
    While FileExist(GFX_PATH & "gui\newmenu\" & i & ".png")
        ReDim Preserve Tex_NewGUI(NumNewGUIs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_NewGUI(NumNewGUIs).filepath = App.Path & GFX_PATH & "gui\newmenu\" & i & ".png"
        Tex_NewGUI(NumNewGUIs).Texture = NumTextures
        NumNewGUIs = NumNewGUIs + 1
        i = i + 1
    Wend
    
    NumNewGUIs = NumNewGUIs - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckGUIs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPanoramas()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumPanoramas = 1
    
    ReDim Tex_Panorama(1)
    While FileExist(GFX_PATH & "panoramas\" & i & GFX_EXT)
        ReDim Preserve Tex_Panorama(NumPanoramas)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Panorama(NumPanoramas).filepath = App.Path & GFX_PATH & "panoramas\" & i & GFX_EXT
        Tex_Panorama(NumPanoramas).Texture = NumTextures
        NumPanoramas = NumPanoramas + 1
        i = i + 1
    Wend
    
    NumPanoramas = NumPanoramas - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPanoramas", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckProjectiles()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumProjectiles = 1
    
    ReDim Tex_Projectile(1)
    While FileExist(GFX_PATH & "projectiles\" & i & GFX_EXT)
        ReDim Preserve Tex_Projectile(NumProjectiles)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Projectile(NumProjectiles).filepath = App.Path & GFX_PATH & "projectiles\" & i & GFX_EXT
        Tex_Projectile(NumProjectiles).Texture = NumTextures
        NumProjectiles = NumProjectiles + 1
        i = i + 1
    Wend
    
    NumProjectiles = NumProjectiles - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckProjectiles", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Public Sub CheckHair()
Dim i As Long, n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ReDim Tex_Hair(0 To TotalHairTypes)
    
    For n = 0 To TotalHairTypes
        i = 1
        NumHair(n) = 1
        
        ReDim Tex_Hair(n).TexHair(1)
    
        While FileExist(GFX_PATH & "hair\" & n & "\" & i & GFX_EXT)
            ReDim Preserve Tex_Hair(n).TexHair(NumHair(n))
            NumTextures = NumTextures + 1
            ReDim Preserve gTexture(NumTextures)
            If FileExist(GFX_PATH & "hair\" & n & "\" & i & GFX_EXT) Then
                Tex_Hair(n).TexHair(NumHair(n)).filepath = App.Path & GFX_PATH & "hair\" & n & "\" & i & GFX_EXT
            
            End If
            Tex_Hair(n).TexHair(NumHair(n)).Texture = NumTextures
            NumHair(n) = NumHair(n) + 1
            i = i + 1
        Wend
        
        NumHair(n) = NumHair(n) - 1
    Next n
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckHair", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTransportes()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumTransportes = 1
    
    ReDim Tex_Transportes(1)
    While FileExist(GFX_PATH & "transportes\" & i & GFX_EXT)
        ReDim Preserve Tex_Transportes(NumTransportes)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Transportes(NumTransportes).filepath = App.Path & GFX_PATH & "transportes\" & i & GFX_EXT
        Tex_Transportes(NumTransportes).Texture = NumTextures
        NumTransportes = NumTransportes + 1
        i = i + 1
    Wend
    
    NumTransportes = NumTransportes - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckProjectiles", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPlanetas()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumPlanetas = 1
    
    ReDim Tex_Planetas(1)
    While FileExist(GFX_PATH & "planets\" & i & ".png")
        ReDim Preserve Tex_Planetas(NumPlanetas)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Planetas(NumPlanetas).filepath = App.Path & GFX_PATH & "planets\" & i & ".png"
        Tex_Planetas(NumPlanetas).Texture = NumTextures
        NumPlanetas = NumPlanetas + 1
        i = i + 1
    Wend
    
    NumPlanetas = NumPlanetas - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckProjectiles", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckParticles()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumParticles = 1
    
    ReDim Tex_Particle(1)
    While FileExist(GFX_PATH & "particles\" & i & GFX_EXT)
        ReDim Preserve Tex_Particle(NumParticles)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Particle(NumParticles).filepath = App.Path & GFX_PATH & "particles\" & i & GFX_EXT
        Tex_Particle(NumParticles).Texture = NumTextures
        NumParticles = NumParticles + 1
        i = i + 1
    Wend
    
    NumParticles = NumParticles - 1
    
    If NumParticles > 0 Then
        For i = 1 To NumParticles
            LoadTexture Tex_Particle(i)
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Checkparticles", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub CheckButtons()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumButtons = 1
    
    ReDim Tex_Buttons(1)
    While FileExist(GFX_PATH & "gui\buttons\" & i & GFX_EXT)
        ReDim Preserve Tex_Buttons(NumButtons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Buttons(NumButtons).filepath = App.Path & GFX_PATH & "gui\buttons\" & i & GFX_EXT
        Tex_Buttons(NumButtons).Texture = NumTextures
        NumButtons = NumButtons + 1
        i = i + 1
    Wend
    
    NumButtons = NumButtons - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckButtons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub CheckButtons_c()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumButtons_c = 1
    
    ReDim Tex_Buttons_c(1)
    While FileExist(GFX_PATH & "gui\buttons\" & i & "_c" & GFX_EXT)
        ReDim Preserve Tex_Buttons_c(NumButtons_c)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Buttons_c(NumButtons_c).filepath = App.Path & GFX_PATH & "gui\buttons\" & i & "_c" & GFX_EXT
        Tex_Buttons_c(NumButtons_c).Texture = NumTextures
        NumButtons_c = NumButtons_c + 1
        i = i + 1
    Wend
    
    NumButtons_c = NumButtons_c - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckButtons_c", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub CheckButtons_h()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    i = 1
    NumButtons_h = 1
    
    ReDim Tex_Buttons_h(1)
    While FileExist(GFX_PATH & "gui\buttons\" & i & "_h" & GFX_EXT)
        ReDim Preserve Tex_Buttons_h(NumButtons_h)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Buttons_h(NumButtons_h).filepath = App.Path & GFX_PATH & "gui\buttons\" & i & "_h" & GFX_EXT
        Tex_Buttons_h(NumButtons_h).Texture = NumTextures
        NumButtons_h = NumButtons_h + 1
        i = i + 1
    Wend
    
    NumButtons_h = NumButtons_h - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckButtons_h", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If AnimInstance(Index).LockToNPC > 0 Then
        TempMapNpc(AnimInstance(Index).LockToNPC).SpawnDelay = 0
    End If

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).name = vbNullString
    Animation(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNPC(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).name = vbNullString
    Npc(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpell(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShop(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResource(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMap()
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.name = vbNullString
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    initAutotiles
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).name)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal name As String)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).name = name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Sprite = Sprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).EXP
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(Index) + 1) ^ 3 - (6 * (GetPlayerLevel(Index) + 1) ^ 2) + 17 * (GetPlayerLevel(Index) + 1) - 12)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerNextLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).EXP = EXP
    If GetPlayerLevel(Index) = MAX_LEVELS And Player(Index).EXP > GetPlayerNextLevel(Index) Then
        Player(Index).EXP = GetPlayerNextLevel(Index)
        Exit Sub
    End If
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerMaxVital = Player(Index).MaxVital(Vital)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    If value <= 0 Then value = 1
    If value > MAX_LONG Then value = MAX_LONG
    Player(Index).Stat(Stat) = value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerStatPoints(ByVal Index As Long, Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStatPoints = Player(Index).StatPoints(Stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStatPoints(ByVal Index As Long, Stat As Stats, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    If value > MAX_LONG Then value = MAX_LONG
    Player(Index).StatPoints(Stat) = value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    
    'Dim TotalPoints, i As Long
    'For i = 1 To 5
    '    TotalPoints = TotalPoints + Player(Index).StatPoints(i)
    'Next i
    
    'GetPlayerPOINTS = Player(Index).RawPDL - TotalPoints
    GetPlayerPOINTS = Player(Index).POINTS
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).POINTS = POINTS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal mapnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = mapnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).X = X
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Y
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Y = Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Dir = Dir
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invSlot).num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).num = ItemNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invSlot).value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = invNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearEvents()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Call ClearEvent(i)
    Next i
End Sub

Public Sub ClearEvent(ByVal Index As Long)
    If Index <= 0 Or Index > MAX_EVENTS Then Exit Sub
    
    Call ZeroMemory(ByVal VarPtr(Events(Index)), LenB(Events(Index)))
    Events(Index).name = vbNullString
End Sub

