Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data files\logs\errors.txt"
    Open filename For Append As #1
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

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call WritePrivateProfileString$(Header, Var, Value, File)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
Dim filename As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\Data Files\config.ini"
    
    Call PutVar(filename, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(filename, "Options", "Username", Trim$(Options.Username))
    Call PutVar(filename, "Options", "Password", Trim$(Options.Password))
    Call PutVar(filename, "Options", "SavePass", str(Options.savePass))
    Call PutVar(filename, "Options", "IP", Options.IP)
    Call PutVar(filename, "Options", "Port", str(Options.Port))
    Call PutVar(filename, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(filename, "Options", "Music", str(Options.Music))
    Call PutVar(filename, "Options", "Sound", str(Options.Sound))
    Call PutVar(filename, "Options", "Debug", str(Options.Debug))
    Call PutVar(filename, "Options", "Volume", str(Options.volume))
    Call PutVar(filename, "Options", "FPSCap", str(Options.FPS))
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim filename As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\Data Files\config.ini"
    
    If Not FileExist(filename, True) Then
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
        SaveOptions
    Else
        Options.Game_Name = GetVar(filename, "Options", "Game_Name")
        Options.Username = GetVar(filename, "Options", "Username")
        Options.Password = GetVar(filename, "Options", "Password")
        Options.savePass = Val(GetVar(filename, "Options", "SavePass"))
        Options.IP = GetVar(filename, "Options", "IP")
        Options.Port = Val(GetVar(filename, "Options", "Port"))
        Options.MenuMusic = GetVar(filename, "Options", "MenuMusic")
        Options.Music = GetVar(filename, "Options", "Music")
        Options.Sound = Val(GetVar(filename, "Options", "Sound"))
        Options.Debug = GetVar(filename, "Options", "Debug")
        Options.volume = GetVar(filename, "Options", "Volume")
        Options.FPS = GetVar(filename, "Options", "FPSCap")
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim filename As String
Dim F As Long
Dim X As Long
Dim Y As Long, i As Long, Z As Long, w As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT

    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Map.Name
    Put #F, , Map.Music
    Put #F, , Map.BGS
    Put #F, , Map.Revision
    Put #F, , Map.Moral
    Put #F, , Map.Up
    Put #F, , Map.Down
    Put #F, , Map.Left
    Put #F, , Map.Right
    Put #F, , Map.BootMap
    Put #F, , Map.BootX
    Put #F, , Map.BootY
    
    Put #F, , Map.Weather
    Put #F, , Map.WeatherIntensity
    
    Put #F, , Map.Fog
    Put #F, , Map.FogSpeed
    Put #F, , Map.FogOpacity
    
    Put #F, , Map.Red
    Put #F, , Map.Green
    Put #F, , Map.Blue
    Put #F, , Map.Alpha
    
    Put #F, , Map.MaxX
    Put #F, , Map.MaxY

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Put #F, , Map.Tile(X, Y)
        Next

        DoEvents
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #F, , Map.Npc(X)
        Put #F, , Map.NpcSpawnType(X)
    Next
    
    Put #F, , Map.Panorama
    Put #F, , Map.Fly
    Put #F, , Map.Ambiente
    Put #F, , Map.FogDir

    Close #F
    
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim filename As String
Dim F As Long
Dim X As Long
Dim Y As Long, i As Long, Z As Long, w As Long, p As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
    ClearMap
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Map.Name
    Get #F, , Map.Music
    Get #F, , Map.BGS
    Get #F, , Map.Revision
    Get #F, , Map.Moral
    Get #F, , Map.Up
    Get #F, , Map.Down
    Get #F, , Map.Left
    Get #F, , Map.Right
    Get #F, , Map.BootMap
    Get #F, , Map.BootX
    Get #F, , Map.BootY
    
    Get #F, , Map.Weather
    Get #F, , Map.WeatherIntensity
        
    Get #F, , Map.Fog
    Get #F, , Map.FogSpeed
    Get #F, , Map.FogOpacity
        
    Get #F, , Map.Red
    Get #F, , Map.Green
    Get #F, , Map.Blue
    Get #F, , Map.Alpha
    
    Get #F, , Map.MaxX
    Get #F, , Map.MaxY
    ' have to set the tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Get #F, , Map.Tile(X, Y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Get #F, , Map.Npc(X)
        Get #F, , Map.NpcSpawnType(X)
    Next
    
    Get #F, , Map.Panorama
    Get #F, , Map.Fly
    Get #F, , Map.Ambiente
    Get #F, , Map.FogDir
    Close #F
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumCharacters = 1
    
    ReDim Tex_Character(1)
    

    While FileExist(GFX_PATH & "characters\" & i & GFX_EXT)
        ReDim Preserve Tex_Character(NumCharacters)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Character(NumCharacters).filepath = App.Path & GFX_PATH & "characters\" & i & GFX_EXT
        Tex_Character(NumCharacters).Texture = NumTextures
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

Public Sub CheckPaperdolls()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    numitems = 1
    
    ReDim Tex_Item(1)

    While FileExist(GFX_PATH & "items\" & i & GFX_EXT)
        ReDim Preserve Tex_Item(numitems)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Item(numitems).filepath = App.Path & GFX_PATH & "items\" & i & GFX_EXT
        Tex_Item(numitems).Texture = NumTextures
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumSpellIcons = 1
    
    ReDim Tex_SpellIcon(1)

    While FileExist(GFX_PATH & "spellicons\" & i & GFX_EXT)
        ReDim Preserve Tex_SpellIcon(NumSpellIcons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_SpellIcon(NumSpellIcons).filepath = App.Path & GFX_PATH & "spellicons\" & i & GFX_EXT
        Tex_SpellIcon(NumSpellIcons).Texture = NumTextures
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    
    NumGUIs = NumGUIs - 1
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
Public Sub CheckParticles()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If AnimInstance(Index).LockToNPC > 0 Then TempMapNpc(AnimInstance(Index).LockToNPC).SpawnDelay = 0

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).Name = vbNullString
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Name = Name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = Value

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_LONG Then Value = MAX_LONG
    Player(Index).Stat(Stat) = Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerStatPoints(ByVal Index As Long, Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStatPoints = Player(Index).StatPoints(Stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStatPoints(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    If Value > MAX_LONG Then Value = MAX_LONG
    Player(Index).StatPoints(Stat) = Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = MapNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invSlot).Num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long, ByVal itemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).Num = itemNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invSlot).Value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).Value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
    Events(Index).Name = vbNullString
End Sub

Public Sub CheckHair()
Dim i As Long, n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ReDim Tex_Hair(0 To TotalHairTypes)
    
    For n = 0 To TotalHairTypes
        i = 1
        NumHair(n) = 1
        
        ReDim Tex_Hair(n).TexHair(1)
    
        While FileExist(GFX_PATH & "hair\" & n & "\" & i & GFX_EXT)
            ReDim Preserve Tex_Hair(n).TexHair(NumHair(n))
            NumTextures = NumTextures + 1
            ReDim Preserve gTexture(NumTextures)
            Tex_Hair(n).TexHair(NumHair(n)).filepath = App.Path & GFX_PATH & "hair\" & n & "\" & i & GFX_EXT
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
    If Options.Debug = 1 Then On Error GoTo errorhandler

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
