Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    filename = App.path & "\data files\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub

Public Sub addReportLog(ByVal Mensagem As String)
Dim filename As String
    filename = App.path & "\data\logs\reports.txt"
    Open filename For Append As #1
        Print #1, "[" & Now & "] " & Mensagem
        Print #1, ""
    Close #1
End Sub

Public Sub addAntiHackLog(ByVal Mensagem As String)
Dim filename As String
    filename = App.path & "\data\logs\antihack.txt"
    Open filename For Append As #1
        Print #1, "[" & Now & "] " & Mensagem
        Print #1, ""
    Close #1
End Sub

Public Sub addItemLog(ByVal Mensagem As String)
Dim filename As String
    filename = App.path & "\data\logs\shopitem.txt"
    Open filename For Append As #1
        Print #1, "[" & Now & "] " & Mensagem
        Print #1, ""
    Close #1
End Sub

Public Sub addFeedback(ByVal Index As Long, ByVal Tipo As Long, ByVal Mensagem As String)
Dim filename As String
Dim TipoName As String
Dim F As Long
Dim Tentativa As Long
    
    Select Case Tipo
        Case 0: TipoName = "Bugs"
        Case 1: TipoName = "Sugestão"
        Case 2: TipoName = "Feedback"
    End Select
    
    Tentativa = 0
    
    SetStatus "Novo feedback de " & GetPlayerName(Index) & "!"
    
    filename = App.path & "\data\feedbacks\[" & Tentativa & "] " & TipoName & " de " & GetPlayerName(Index) & ".txt"
    
    
    Do While FileExist(filename, True)
        Tentativa = Tentativa + 1
        filename = App.path & "\data\feedbacks\[" & Tentativa & "] " & TipoName & " de " & GetPlayerName(Index) & ".txt"
    Loop
    
    F = FreeFile
    Open filename For Output As #F
    Close #F
    
    Open filename For Append As #1
        Print #1, Mensagem
        Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    PutVar App.path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.path & "\data\options.ini", "OPTIONS", "Port", STR(Options.Port)
    PutVar App.path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    PutVar App.path & "\data\options.ini", "OPTIONS", "Logs", STR(Options.Logs)
    PutVar App.path & "\data\options.ini", "OPTIONS", "EventChance", STR(Options.EventChance)
    PutVar App.path & "\data\options.ini", "OPTIONS", "ExpFactor", STR(Options.ExpFactor)
    PutVar App.path & "\data\options.ini", "OPTIONS", "DropFactor", STR(Options.DropFactor)
    PutVar App.path & "\data\options.ini", "OPTIONS", "ResourceFactor", STR(Options.ResourceFactor)
End Sub

Public Sub LoadOptions()
    Options.Game_Name = GetVar(App.path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.path & "\data\options.ini", "OPTIONS", "Website")
    Options.Logs = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "Logs"))
    Options.EventChance = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "EventChance"))
    Options.Language = GetVar(App.path & "\data\options.ini", "OPTIONS", "Language")
    Options.ExpFactor = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "ExpFactor"))
    Options.DropFactor = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "DropFactor"))
    Options.ResourceFactor = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "ResourceFactor"))
    Options.GoldFactor = 1
    If Options.ExpFactor <= 0 Then Options.ExpFactor = 1
    
    SetupBonuses
    frmServer.txtExpFactor = Options.ExpFactor
    frmServer.txtDrop = Options.DropFactor
    frmServer.txtResource = Options.ResourceFactor
    frmServer.txtGold = Options.GoldFactor
    
    START_MAP = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "StartMap"))
    START_X = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "StartY"))
    START_Y = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "StartX"))
    
    RESPAWN_MAP = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "RespawnMap"))
    RESPAWN_X = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "RespawnY"))
    RESPAWN_Y = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "RespawnX"))
    
    MoedaZ = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "Zeni"))
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    filename = App.path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open filename For Append As #F
    Print #F, IP & "," & GetPlayerName(BannedByIndex)
    Close #F
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AlertMSG(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    filename = App.path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open filename For Append As #F
    Print #F, IP & "," & "Server"
    Close #F
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & "the Server" & "!", White)
    Call AlertMSG(BanPlayerIndex, "You have been banned by " & "The Server" & "!")
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    filename = "data\accounts\" & Trim(Name) & ".bin"

    If FileExist(filename) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    If AccountExist(Name) Then
        filename = App.path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    
    ClearPlayer Index
    
    Player(Index).Login = Name
    Player(Index).Password = Password

    Call SavePlayer(Index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.path & "\data\accounts\charlist.txt", App.path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal Index As Long) As Boolean

    If LenB(Trim$(Player(Index).Name)) > 0 Then
        CharExist = True
    End If

End Function

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long, ByVal Hair As Byte)
    Dim F As Long
    Dim n As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(Player(Index).Name)) = 0 Then
        
        spritecheck = False
        
        Player(Index).Name = Name
        Player(Index).Sex = Sex
        Player(Index).Class = ClassNum
        Player(Index).Hair = Hair
        
        If Player(Index).Sex = SEX_MALE Then
            Player(Index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(Index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If
        Player(Index).RealSprite = Player(Index).Sprite

        Player(Index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(Index).stat(n) = 5 'Class(ClassNum).stat(n)
            Player(Index).statPoints(n) = 0
        Next n
        
        Player(Index).Points = 0
        Player(Index).PDL = 0

        Player(Index).Dir = DIR_DOWN
        Player(Index).Map = START_MAP
        Player(Index).X = START_X
        Player(Index).Y = START_Y
        Player(Index).Vital(Vitals.HP) = GetPlayerMaxVital(Index, Vitals.HP)
        Player(Index).Vital(Vitals.MP) = GetPlayerMaxVital(Index, Vitals.MP)
        Player(Index).Version = CHARACTER_VERSION
        
        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(Index).Inv(n).Num = Class(ClassNum).StartItem(n)
                        Player(Index).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If
        
        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartSpell(n)).Name)) > 0 Then
                        Player(Index).Spell(n) = Class(ClassNum).StartSpell(n)
                    End If
                End If
            Next
        End If
        
        ' Append name to file
        F = FreeFile
        Open App.path & "\data\accounts\charlist.txt" For Append As #F
        Print #F, Name
        Close #F
        Call SavePlayer(Index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim s As String
    F = FreeFile
    Open App.path & "\data\accounts\charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim filename As String
    Dim F As Long

    filename = App.path & "\data\accounts\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    
    On Error Resume Next
    Open filename For Binary As #F
    Put #F, , Player(Index)
    Close #F
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long
    Call ClearPlayer(Index)
    filename = App.path & "\data\accounts\" & Trim(Name) & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Player(Index)
    Close #F
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    Player(Index).Name = vbNullString
    Player(Index).Class = 1
    Player(Index).GravityInit = vbNullString
    Player(Index).Daily.LastDate = vbNullString
    Player(Index).LastLogin = vbNullString

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(4) = vbNullString
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim filename As String
    Dim File As String
    filename = App.path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(filename, True) Then
        File = FreeFile
        Open filename For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim X As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        filename = App.path & "\data\classes.ini"
        Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' continue
        Class(i).stat(Stats.Strength) = Val(GetVar(filename, "CLASS" & i, "Strength"))
        Class(i).stat(Stats.Endurance) = Val(GetVar(filename, "CLASS" & i, "Endurance"))
        Class(i).stat(Stats.Intelligence) = Val(GetVar(filename, "CLASS" & i, "Intelligence"))
        Class(i).stat(Stats.agility) = Val(GetVar(filename, "CLASS" & i, "Agility"))
        Class(i).stat(Stats.Willpower) = Val(GetVar(filename, "CLASS" & i, "Willpower"))
        
        ' how many starting items?
        startItemCount = Val(GetVar(filename, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)
        
        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For X = 1 To startItemCount
                Class(i).StartItem(X) = Val(GetVar(filename, "CLASS" & i, "StartItem" & X))
                Class(i).StartValue(X) = Val(GetVar(filename, "CLASS" & i, "StartValue" & X))
            Next
        End If
        
        ' how many starting spells?
        startSpellCount = Val(GetVar(filename, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)
        
        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_INV Then
            For X = 1 To startSpellCount
                Class(i).StartSpell(X) = Val(GetVar(filename, "CLASS" & i, "StartSpell" & X))
            Next
        End If
    Next

End Sub

Sub SaveClasses()
    Dim filename As String
    Dim i As Long
    Dim X As Long
    
    filename = App.path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "Maleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Strength", STR(Class(i).stat(Stats.Strength)))
        Call PutVar(filename, "CLASS" & i, "Endurance", STR(Class(i).stat(Stats.Endurance)))
        Call PutVar(filename, "CLASS" & i, "Intelligence", STR(Class(i).stat(Stats.Intelligence)))
        Call PutVar(filename, "CLASS" & i, "Agility", STR(Class(i).stat(Stats.agility)))
        Call PutVar(filename, "CLASS" & i, "Willpower", STR(Class(i).stat(Stats.Willpower)))
        ' loop for items & values
        For X = 1 To UBound(Class(i).StartItem)
            Call PutVar(filename, "CLASS" & i, "StartItem" & X, STR(Class(i).StartItem(X)))
            Call PutVar(filename, "CLASS" & i, "StartValue" & X, STR(Class(i).StartValue(X)))
        Next
        ' loop for spells
        For X = 1 To UBound(Class(i).StartSpell)
            Call PutVar(filename, "CLASS" & i, "StartSpell" & X, STR(Class(i).StartSpell(X)))
        Next
    Next

End Sub

Function CheckClasses() As Boolean
    Dim filename As String
    filename = App.path & "\data\classes.ini"

    If Not FileExist(filename, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim filename As String
    Dim F  As Long
    filename = App.path & "\data\items\item" & ItemNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Item(ItemNum)
    Close #F
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Item(i)
        Close #F
    Next

End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\shops\shop" & shopNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Shop(shopNum)
    Close #F
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.path & "\data\shops\shop" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Shop(i)
        Close #F
    Next

End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

'Guilds
Sub SaveGuilds()
    Dim i As Long

    For i = 1 To MAX_GUILDS
        Call SaveGuild(i)
    Next

End Sub

Sub SaveGuild(ByVal GuildNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\Guilds\Guild" & GuildNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Guild(GuildNum)
    Close #F
End Sub

Sub LoadGuilds()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckGuilds

    For i = 1 To MAX_GUILDS
        filename = App.path & "\data\Guilds\Guild" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Guild(i)
        Close #F
    Next

End Sub

Sub CheckGuilds()
    Dim i As Long

    For i = 1 To MAX_GUILDS

        If Not FileExist("\Data\Guilds\Guild" & i & ".dat") Then
            Call SaveGuild(i)
        End If

    Next

End Sub

Sub ClearGuild(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Guild(Index)), LenB(Guild(Index)))
    Guild(Index).Name = vbNullString
End Sub

Sub ClearGuilds()
    Dim i As Long

    For i = 1 To MAX_GUILDS
        Call ClearGuild(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal SpellNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\spells\spells" & SpellNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Spell(SpellNum)
    Close #F
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.path & "\data\spells\spells" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Spell(i)
        Close #F
    Next

End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal NpcNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\npcs\npc" & NpcNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Npc(NpcNum)
    Close #F
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        filename = App.path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Npc(i)
        Close #F
    Next

End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).Name = vbNullString
    Npc(Index).AttackSay = vbNullString
    Npc(Index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Resource(i)
        Close #F
    Next

End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        filename = App.path & "\data\animations\animation" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Animation(i)
        Close #F
    Next

End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub DeprecatedSaveMap(ByVal MapNum As Long)
    Dim filename As String
    If MapNum >= PlanetStart And MapNum <= PlanetStart + MAX_PLANET_BASE + MAX_PLAYER_PLANETS Then Exit Sub 'Não salvar estes mapas
    filename = App.path & "\data\maps\map" & MapNum & ".dat"
    SaveMapData Map(MapNum), filename
End Sub
Sub SaveMapData(ByRef MapData As MapRec, ByVal filename As String)
    Dim F As Long
    Dim X As Long
    Dim Y As Long, i As Long, z As Long, w As Long
    
    'Filename = App.path & "\data\maps\map" & MapNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , MapData.Name
    Put #F, , MapData.Music
    Put #F, , MapData.BGS
    Put #F, , MapData.Revision
    Put #F, , MapData.Moral
    Put #F, , MapData.Up
    Put #F, , MapData.Down
    Put #F, , MapData.Left
    Put #F, , MapData.Right
    Put #F, , MapData.BootMap
    Put #F, , MapData.BootX
    Put #F, , MapData.BootY
    
    Put #F, , MapData.Weather
    Put #F, , MapData.WeatherIntensity
    
    Put #F, , MapData.Fog
    Put #F, , MapData.FogSpeed
    Put #F, , MapData.FogOpacity
    
    Put #F, , MapData.Red
    Put #F, , MapData.Green
    Put #F, , MapData.Blue
    Put #F, , MapData.Alpha
    
    Put #F, , MapData.MaxX
    Put #F, , MapData.MaxY
    
    If MapData.MaxX = 0 Or MapData.MaxY = 0 Then
        MapData.MaxX = 24
        MapData.MaxY = 18
        ReDim MapData.Tile(0 To MapData.MaxX, 0 To MapData.MaxY)
        SaveMapData MapData, filename
        Exit Sub
    End If

    For X = 0 To MapData.MaxX
        For Y = 0 To MapData.MaxY
            Put #F, , MapData.Tile(X, Y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #F, , MapData.Npc(X)
        Put #F, , MapData.NpcSpawnType(X)
    Next
    Put #F, , MapData.Panorama
    
    Put #F, , MapData.Fly
    Put #F, , MapData.Ambiente
    Put #F, , MapData.FogDir
    Close #F
    
    
    
    DoEvents
End Sub

Sub SaveMap(ByVal MapNum As Long)
    Dim filename As String
    Dim F As Long
    Dim X As Long
    Dim Y As Long
    filename = App.path & "\data\maps\map" & MapNum & ".dat"
    F = FreeFile
    
    If Map(MapNum).MaxX = 0 Or Map(MapNum).MaxY = 0 Then
        ReDim Map(MapNum).Tile(0 To 24, 0 To 18) As TileRec
        Map(MapNum).MaxX = 24
        Map(MapNum).MaxY = 18
    End If
    
    Open filename For Binary As #F
    Put #F, , Map(MapNum).Name
    Put #F, , Map(MapNum).Music
    Put #F, , Map(MapNum).Revision
    Put #F, , Map(MapNum).Moral
    Put #F, , Map(MapNum).Up
    Put #F, , Map(MapNum).Down
    Put #F, , Map(MapNum).Left
    Put #F, , Map(MapNum).Right
    Put #F, , Map(MapNum).BootMap
    Put #F, , Map(MapNum).BootX
    Put #F, , Map(MapNum).BootY
    Put #F, , Map(MapNum).MaxX
    Put #F, , Map(MapNum).MaxY

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            Put #F, , Map(MapNum).Tile(X, Y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #F, , Map(MapNum).Npc(X)
    Next
    
    Put #F, , Map(MapNum).BossNpc
    Put #F, , Map(MapNum).Fog
    Put #F, , Map(MapNum).FogOpacity
    Put #F, , Map(MapNum).FogSpeed
    Put #F, , Map(MapNum).Panorama
    Put #F, , Map(MapNum).SunRays
    Put #F, , Map(MapNum).Fly
    Close #F
    
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Sub DeprecatedLoadMaps()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim X As Long
    Dim Y As Long, z As Long, p As Long, w As Long
    Dim newtileset As Long, newtiley As Long
    
    If UZ Then
        MAX_PLANET_CUSTOM = GetTotalCustomPlanets + 1
        MAX_PLANETS = MAX_PLANET_BASE + MAX_PLANET_CUSTOM + 1
    End If
    Call CheckMaps

    For i = 1 To MAX_MAPS
        filename = App.path & "\data\maps\map" & i & ".dat"
        F = FreeFile
        
        If UZ Then
            If i >= PlanetStart And i < PlanetStart + MAX_PLANET_BASE Then
                ReDim MapNpc(i).Npc(1 To MAX_MAP_NPCS) As MapNpcRec
                ReDim MapNpc(i).TempNpc(1 To MAX_MAP_NPCS) As TempMapNpcRec
                GoTo NextMap
            End If
        End If
        
        DeprecatedLoadMap filename, Map(i)
        SetupMapNpcs i
        CacheResources i
        DoEvents
        CacheMapBlocks i
NextMap:
    Next
    
End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim X As Long
    Dim Y As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        filename = App.path & "\data\maps\map" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Map(i).Name
        Get #F, , Map(i).Music
        Get #F, , Map(i).Revision
        Get #F, , Map(i).Moral
        Get #F, , Map(i).Up
        Get #F, , Map(i).Down
        Get #F, , Map(i).Left
        Get #F, , Map(i).Right
        Get #F, , Map(i).BootMap
        Get #F, , Map(i).BootX
        Get #F, , Map(i).BootY
        Get #F, , Map(i).MaxX
        Get #F, , Map(i).MaxY
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        For X = 0 To Map(i).MaxX
            For Y = 0 To Map(i).MaxY
                Get #F, , Map(i).Tile(X, Y)
                'Dim OldTile As TileRec2
                'Get #F, , OldTile 'MapData.Tile(X, Y)
                'Dim n As Long
                'For n = 1 To 5
                'Map(i).Tile(X, Y).Autotile(n) = OldTile.Autotile(n)
                'Map(i).Tile(X, Y).Layer(n) = OldTile.Layer(n)
                'Next n
                'Map(i).Tile(X, Y).data1 = OldTile.data1
                'Map(i).Tile(X, Y).data2 = OldTile.data2
                'Map(i).Tile(X, Y).data3 = OldTile.data3
                'Map(i).Tile(X, Y).Data4 = vbNullString
                'Map(i).Tile(X, Y).DirBlock = OldTile.DirBlock
                
                'Map(i).Tile(X, Y).Type = OldTile.Type
            Next
        Next

        For X = 1 To MAX_MAP_NPCS
            Get #F, , Map(i).Npc(X)
            ReDim MapNpc(i).Npc(1 To MAX_MAP_NPCS) As MapNpcRec
            ReDim MapNpc(i).TempNpc(1 To MAX_MAP_NPCS) As TempMapNpcRec
            MapNpc(i).Npc(X).Num = Map(i).Npc(X)
        Next
        
        Get #F, , Map(i).BossNpc
        Get #F, , Map(i).Fog
        Get #F, , Map(i).FogOpacity
        Get #F, , Map(i).FogSpeed
        Get #F, , Map(i).Panorama
        Get #F, , Map(i).SunRays
        Get #F, , Map(i).Fly
        Close #F
        
        SetupMapNpcs i
        
        CacheResources i
        CacheMapBlocks i
        DoEvents
    Next
End Sub

Sub DeprecatedLoadMap(ByVal filename As String, ByRef MapData As MapRec)
    Dim F As Long
    Dim X As Long, Y As Long
    F = FreeFile
    
    Open filename For Binary As #F
        Get #F, , MapData.Name
        Get #F, , MapData.Music
        Get #F, , MapData.BGS
        Get #F, , MapData.Revision
        Get #F, , MapData.Moral
        Get #F, , MapData.Up
        Get #F, , MapData.Down
        Get #F, , MapData.Left
        Get #F, , MapData.Right
        Get #F, , MapData.BootMap
        Get #F, , MapData.BootX
        Get #F, , MapData.BootY
        
        Get #F, , MapData.Weather
        Get #F, , MapData.WeatherIntensity
        
        Get #F, , MapData.Fog
        Get #F, , MapData.FogSpeed
        Get #F, , MapData.FogOpacity
        
        Get #F, , MapData.Red
        Get #F, , MapData.Green
        Get #F, , MapData.Blue
        Get #F, , MapData.Alpha
        
        Get #F, , MapData.MaxX
        Get #F, , MapData.MaxY
        ' have to set the tile()
        ReDim MapData.Tile(0 To MapData.MaxX, 0 To MapData.MaxY)

        For X = 0 To MapData.MaxX
            For Y = 0 To MapData.MaxY
                Dim OldTile As TileRec2
                Get #F, , OldTile 'MapData.Tile(X, Y)
                Dim n As Long
                For n = 1 To 5
                MapData.Tile(X, Y).Autotile(n) = OldTile.Autotile(n)
                MapData.Tile(X, Y).Layer(n) = OldTile.Layer(n)
                Next n
                MapData.Tile(X, Y).data1 = OldTile.data1
                MapData.Tile(X, Y).data2 = OldTile.data2
                MapData.Tile(X, Y).data3 = OldTile.data3
                MapData.Tile(X, Y).Data4 = vbNullString
                MapData.Tile(X, Y).DirBlock = OldTile.DirBlock
                
                MapData.Tile(X, Y).Type = OldTile.Type
            Next
        Next

        For X = 1 To MAX_MAP_NPCS
            Get #F, , MapData.Npc(X)
            Get #F, , MapData.NpcSpawnType(X)
            'ReDim MapNpc(i).Npc(1 To MAX_MAP_NPCS) As MapNpcRec
            'ReDim MapNpc(i).TempNpc(1 To MAX_MAP_NPCS) As TempMapNpcRec
            'MapNpc(i).Npc(X).Num = MapData.Npc(X)
        Next
        
        Get #F, , MapData.Panorama
        
        Get #F, , MapData.Fly
        Get #F, , MapData.Ambiente
        Close #F
End Sub

Sub SetupMapNpcs(ByVal MapNum As Long)
    Dim X As Long
    ReDim MapNpc(MapNum).Npc(1 To MAX_MAP_NPCS) As MapNpcRec
    ReDim MapNpc(MapNum).TempNpc(1 To MAX_MAP_NPCS) As TempMapNpcRec
    For X = 1 To MAX_MAP_NPCS
        MapNpc(MapNum).Npc(X).Num = Map(MapNum).Npc(X)
    Next
End Sub

Sub CheckMaps()
    Dim i As Long
    
    MAX_MAPS = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "Maps"))
    If UZ Then MAX_MAPS = PlanetStart + MAX_PLANET_BASE + MAX_PLAYER_PLANETS
    If MAX_MAPS < 0 Then MAX_MAPS = 1
    
    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim MapCache(1 To MAX_MAPS) As Cache
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNpc(1 To MAX_MAPS) As MapDataRec
    ReDim MapBlocks(1 To MAX_MAPS) As MapBlockRec

    For i = 1 To MAX_MAPS
        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If
    Next

End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, Index)), LenB(MapItem(MapNum, Index)))
    MapItem(MapNum, Index).PlayerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    ReDim MapNpc(MapNum).Npc(1 To MAX_MAP_NPCS)
    ReDim MapNpc(MapNum).TempNpc(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).Npc(Index)), LenB(MapNpc(MapNum).Npc(Index)))
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).TempNpc(Index)), LenB(MapNpc(MapNum).TempNpc(Index)))
End Sub

Sub ClearMapNpcs()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, Y)
        Next
    Next

End Sub

Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).Name = vbNullString
    Map(MapNum).MaxX = MAX_MAPX
    Map(MapNum).MaxY = MAX_MAPY
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.stat(Intelligence) * 10) + 2
            End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal stat As Stats) As Long
    GetClassStat = Class(ClassNum).stat(stat)
End Function

Sub SaveBank(ByVal Index As Long)
    Dim filename As String
    Dim F As Long
    
    filename = App.path & "\data\banks\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Bank(Index)
    Close #F
End Sub

Public Sub LoadBank(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long

    Call ClearBank(Index)

    filename = App.path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(filename, True) Then
        Call SaveBank(Index)
        Exit Sub
    End If

    F = FreeFile
    Open filename For Binary As #F
        Get #F, , Bank(Index)
    Close #F

End Sub

Sub ClearBank(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(Index)), LenB(Bank(Index)))
End Sub

Sub ClearParty(ByVal PartyNum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(PartyNum)), LenB(Party(PartyNum)))
End Sub

Sub SaveSwitches()
Dim i As Long, filename As String
filename = App.path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Call PutVar(filename, "Switches", "Switch" & CStr(i) & "Name", Switches(i))
Next

End Sub

Sub SaveVariables()
Dim i As Long, filename As String
filename = App.path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Call PutVar(filename, "Variables", "Variable" & CStr(i) & "Name", Variables(i))
Next

End Sub

Sub LoadSwitches()
Dim i As Long, filename As String
    filename = App.path & "\data\switches.ini"
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = GetVar(filename, "Switches", "Switch" & CStr(i) & "Name")
    Next
End Sub

Sub LoadVariables()
Dim i As Long, filename As String
    filename = App.path & "\data\variables.ini"
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = GetVar(filename, "Variables", "Variable" & CStr(i) & "Name")
    Next
End Sub

Public Sub SaveChatLog(ByVal LogType As Long)
    Dim Foldername As String, filename As String
    Dim F As Long
    
    If Options.Logs > 0 Then
        Select Case LogType
            Case ChatGlobal
                Foldername = "global\"
            Case ChatMap
                Foldername = "map\"
            Case ChatEmote
                Foldername = "emote\"
            Case ChatPlayer
                Foldername = "player\"
            Case ChatSystem
                Foldername = "system\"
        End Select

        filename = App.path & "\data\logs\" & Trim(Foldername) & DateValue(Now) & ".log"

        If Not FileExist(filename, True) Then
            F = FreeFile
            Open filename For Output As #F
            Close #F
        End If

        F = FreeFile
        Open filename For Append As #F
        Print #F, frmServer.txtText(LogType)
        Close #F
    End If
End Sub

Public Sub SaveChatLine(ByVal LogType As Long, ByVal Line As String)
    Dim Foldername As String, filename As String
    Dim F As Long
    
    If Options.Logs > 0 Then
        Select Case LogType
            Case ChatGlobal
                Foldername = "global\"
            Case ChatMap
                Foldername = "map\"
            Case ChatEmote
                Foldername = "emote\"
            Case ChatPlayer
                Foldername = "player\"
            Case ChatSystem
                Foldername = "system\"
            Case Else
                Foldername = ""
        End Select

        filename = App.path & "\data\logs\" & Trim(Foldername) & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & ".log"

        If Not FileExist(filename, True) Then
            F = FreeFile
            Open filename For Output As #F
            Close #F
        End If

        F = FreeFile
        Open filename For Append As #F
        Print #F, Line
        Close #F
    End If
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

Public Sub LoadEvents()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Call LoadEvent(i)
    Next i
End Sub

Public Sub LoadEvent(ByVal Index As Long)
    On Error GoTo Errorhandle
    
    Dim F As Long, SCount As Long, s As Long, DCount As Long, D As Long
    Dim filename As String
    filename = App.path & "\data\events\event" & Index & ".dat"
    If FileExist(filename, True) Then
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Events(Index).Name
            Get #F, , Events(Index).chkSwitch
            Get #F, , Events(Index).chkVariable
            Get #F, , Events(Index).chkHasItem
            Get #F, , Events(Index).SwitchIndex
            Get #F, , Events(Index).SwitchCompare
            Get #F, , Events(Index).VariableIndex
            Get #F, , Events(Index).VariableCompare
            Get #F, , Events(Index).VariableCondition
            Get #F, , Events(Index).HasItemIndex
            Get #F, , SCount
            If SCount <= 0 Then
                Events(Index).HasSubEvents = False
                Erase Events(Index).SubEvents
            Else
                Events(Index).HasSubEvents = True
                ReDim Events(Index).SubEvents(1 To SCount)
                For s = 1 To SCount
                    With Events(Index).SubEvents(s)
                        Get #F, , .Type
                        Get #F, , DCount
                        If DCount <= 0 Then
                            .HasText = False
                            Erase .Text
                        Else
                            .HasText = True
                            ReDim .Text(1 To DCount)
                            For D = 1 To DCount
                                Get #F, , .Text(D)
                            Next D
                        End If
                        Get #F, , DCount
                        If DCount <= 0 Then
                            .HasData = False
                            Erase .Data
                        Else
                            .HasData = True
                            ReDim .Data(1 To DCount)
                            For D = 1 To DCount
                                Get #F, , .Data(D)
                            Next D
                        End If
                    End With
                Next s
            End If
            Get #F, , Events(Index).Trigger
            Get #F, , Events(Index).WalkThrought
        Close #F
    Else
        Call ClearEvent(Index)
        Call SaveEvent(Index)
    End If
    Exit Sub
Errorhandle:
    HandleError "LoadEvent(Long)", "modEvents", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Call ClearEvent(Index)
End Sub

Public Sub SaveEvents()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Call SaveEvent(i)
    Next i
End Sub
Public Sub SaveEvent(ByVal Index As Long)
    Dim F As Long, SCount As Long, s As Long, DCount As Long, D As Long
    Dim filename As String
    filename = App.path & "\data\events\event" & Index & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Events(Index).Name
        Put #F, , Events(Index).chkSwitch
        Put #F, , Events(Index).chkVariable
        Put #F, , Events(Index).chkHasItem
        Put #F, , Events(Index).SwitchIndex
        Put #F, , Events(Index).SwitchCompare
        Put #F, , Events(Index).VariableIndex
        Put #F, , Events(Index).VariableCompare
        Put #F, , Events(Index).VariableCondition
        Put #F, , Events(Index).HasItemIndex
        If Not (Events(Index).HasSubEvents) Then
            SCount = 0
            Put #F, , SCount
        Else
            SCount = UBound(Events(Index).SubEvents)
            Put #F, , SCount
            For s = 1 To SCount
                With Events(Index).SubEvents(s)
                    Put #F, , .Type
                    If Not (.HasText) Then
                        DCount = 0
                        Put #F, , DCount
                    Else
                        DCount = UBound(.Text)
                        Put #F, , DCount
                        For D = 1 To DCount
                            Put #F, , .Text(D)
                        Next D
                    End If
                    If Not (.HasData) Then
                        DCount = 0
                        Put #F, , DCount
                    Else
                        DCount = UBound(.Data)
                        Put #F, , DCount
                        For D = 1 To DCount
                            Put #F, , .Data(D)
                        Next D
                    End If
                End With
            Next s
        End If
        Put #F, , Events(Index).Trigger
        Put #F, , Events(Index).WalkThrought
    Close #F
End Sub

Sub SaveEffects()
    Dim i As Long

    For i = 1 To MAX_EFFECTS
        Call SaveEffect(i)
    Next

End Sub

Sub SaveEffect(ByVal EffectNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\effects\effect" & EffectNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Effect(EffectNum)
    Close #F
End Sub

Sub LoadEffects()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckEffects

    For i = 1 To MAX_EFFECTS
        filename = App.path & "\data\effects\effect" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Effect(i)
        Close #F
    Next

End Sub

Sub CheckEffects()
    Dim i As Long

    For i = 1 To MAX_EFFECTS

        If Not FileExist("\Data\Effects\Effect" & i & ".dat") Then
            Call SaveEffect(i)
        End If

    Next

End Sub

Sub ClearEffect(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Effect(Index)), LenB(Effect(Index)))
    Effect(Index).Name = vbNullString
    Effect(Index).Sound = "None."
End Sub

Sub ClearEffects()
    Dim i As Long

    For i = 1 To MAX_EFFECTS
        Call ClearEffect(i)
    Next
End Sub

Sub LoadHouses()
    Dim i As Long
    i = 1
    Do While GetVar(App.path & "\houses.ini", "CASAS", "Proprietario" & i) <> ""
    i = i + 1
    Loop
    i = i - 1
    
    TotalHouses = i
    If TotalHouses < 1 Then
        ReDim House(1 To 1) As HouseRec
    Else
        ReDim House(1 To TotalHouses) As HouseRec
        
        Dim n As Long
        For n = 1 To TotalHouses
            House(n).Proprietario = GetVar(App.path & "\houses.ini", "CASAS", "Proprietario" & n)
            House(n).DataDeInicio = GetVar(App.path & "\houses.ini", "CASAS", "DataDeInicio" & n)
            House(n).Dias = Val(GetVar(App.path & "\houses.ini", "CASAS", "Dias" & n))
            
            If frmServer.lstHouse.ListItems.Count < n Then frmServer.lstHouse.ListItems.Add (n)
            
            For i = 1 To 4
            If frmServer.lstHouse.ListItems(n).SubItems(i) <> "" Then frmServer.lstHouse.ListItems(n).SubItems(i) = vbNullString
            Next i
            
            frmServer.lstHouse.ListItems(n).SubItems(1) = House(n).Proprietario
            frmServer.lstHouse.ListItems(n).SubItems(2) = House(n).DataDeInicio
            frmServer.lstHouse.ListItems(n).SubItems(3) = House(n).Dias
            frmServer.lstHouse.ListItems(n).SubItems(4) = House(n).Dias - (DateDiff("d", House(n).DataDeInicio, Date))
        Next n
    End If
End Sub

Sub SaveHouses()
    Dim n As Long
    For n = 1 To TotalHouses
        Call PutVar(App.path & "\houses.ini", "CASAS", "Proprietario" & n, House(n).Proprietario)
        Call PutVar(App.path & "\houses.ini", "CASAS", "DataDeInicio" & n, House(n).DataDeInicio)
        Call PutVar(App.path & "\houses.ini", "CASAS", "Dias" & n, STR(House(n).Dias))
    Next n
End Sub

Sub CheckTransportes()
    Dim filename As String
    Dim i As Long
    
    Dim n As Long
    n = 1
    Do While FileExist("data\transports\" & n & ".ini")
        n = n + 1
    Loop
    n = n - 1
    
    For i = 1 To n
        filename = App.path & "\data\transports\" & i & ".ini"
        
        If FileExist(filename, True) Then
            With Transporte(i)
                .Nome = GetVar(filename, "TRANSPORT", "Name")
                .Map = Val(GetVar(filename, "TRANSPORT", "Map"))
                .AlterMap = Val(GetVar(filename, "TRANSPORT", "AlterMap"))
                .LoadMap = Val(GetVar(filename, "TRANSPORT", "LoadMap"))
                .LoadX = Val(GetVar(filename, "TRANSPORT", "LoadX"))
                .LoadY = Val(GetVar(filename, "TRANSPORT", "LoadY"))
                .TravelMap = Val(GetVar(filename, "TRANSPORT", "TravelMap"))
                .DestinyMap = Val(GetVar(filename, "TRANSPORT", "DestinyMap"))
                .DestinyX = Val(GetVar(filename, "TRANSPORT", "DestinyX"))
                .DestinyY = Val(GetVar(filename, "TRANSPORT", "DestinyY"))
                .AlterDestinyMap = Val(GetVar(filename, "TRANSPORT", "AlterDestinyMap"))
                .AlterDestinyX = Val(GetVar(filename, "TRANSPORT", "AlterDestinyX"))
                .AlterDestinyY = Val(GetVar(filename, "TRANSPORT", "AlterDestinyY"))
                .Tick = GetTickCount
                .IntervalWait = Val(GetVar(filename, "TRANSPORT", "WaitInterval"))
                .IntervalTravel = Val(GetVar(filename, "TRANSPORT", "TravelInterval"))
                .Embarque = Val(GetVar(filename, "TRANSPORT", "LoadInterval"))
                .Sound = Val(GetVar(filename, "TRANSPORT", "Sound"))
                .Passaporte = Val(GetVar(filename, "TRANSPORT", "Passport"))
            End With
        End If
    Next i
    
End Sub

' ************
' ** QUESTs **
' ************
Sub SaveQuests()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next

End Sub

Sub SaveQuest(ByVal questnum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\quests\quest" & questnum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Quest(questnum)
    Close #F
End Sub

Sub LoadQuests()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call Checkquests

    For i = 1 To MAX_QUESTS
        filename = App.path & "\data\quests\quest" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Quest(i)
        Close #F
    Next

End Sub

Sub Checkquests()
    Dim i As Long

    For i = 1 To MAX_QUESTS

        If Not FileExist("\Data\quests\quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If

    Next

End Sub

Sub Clearquest(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = vbNullString
    Quest(Index).Desc = vbNullString
End Sub

Sub Clearquests()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        Call Clearquest(i)
    Next

End Sub

Sub CheckEXP()
    Dim i As Long
    MAX_LEVELS = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "Levels"))
    MAX_STAT_LEVELS = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "StatLevels"))
    LevelUpBonus = Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "LevelUpPDLBonus"))
    ExpToPDL = 1 / Val(GetVar(App.path & "\data\options.ini", "OPTIONS", "ExpToPDLConversion"))
    ReDim Experience(1 To MAX_LEVELS) As Currency
    ReDim PDLBase(1 To MAX_LEVELS)
    ReDim StatExperience(1 To MAX_STAT_LEVELS) As Currency
    For i = 1 To MAX_LEVELS
        Experience(i) = Val(GetVar(App.path & "\data\exp.ini", "EXPERIENCE", "Exp" & i))
        If i >= 60 Then Experience(i) = Experience(i) * 2
        If i >= 90 Then Experience(i) = Experience(i) * 2
        If i >= 95 Then Experience(i) = Experience(i) * 2
        DoEvents
    Next i
    For i = 1 To MAX_STAT_LEVELS
        StatExperience(i) = Val(GetVar(App.path & "\data\stats.ini", "EXPERIENCE", "Exp" & i))
        DoEvents
    Next i
    For i = 1 To MAX_LEVELS
        PDLBase(i) = GetLevelPDL(i)
    Next i
End Sub

Sub MakePDLXP()
Dim Count As Currency, i As Long, F As Long
    For i = 1 To 200
        Dim mult As Double
        mult = 1
        For F = 1 To i
            mult = mult * 1.5
        Next F
        Call PutVar(App.path & "\data\stats.ini", "EXPERIENCE", "Exp" & i, Val((12 * i) * mult))
    Next i
End Sub

Sub CheckProvacoes()
    
    Dim n As Long
    n = 1
    Do While FileExist("data\provation\" & n & ".dat")
        n = n + 1
    Loop
    n = n - 1
    
    If n = 0 Then
        ReDim Provação(1 To 1) As ProvRec
        Exit Sub
    End If
    
    ReDim Provação(1 To n) As ProvRec
    
    Dim i As Long
    For i = 1 To n
        LoadProvation i
    Next i
    
    ProvaçãoCount = n
    
    Exit Sub
    
    'No need for this
    ReDim Provação(1).Wave(1 To 4) As WaveRec
    
    With Provação(1)
        .Map = 40
        .X = 12
        .Y = 12
        .Cost = 15000
        .MinLevel = 20
        .RewardXP = 15000
        .TradeItem = 0
        .RewardItem = 42
        
        Call AddWave(18, 5000, 5, 1)
        Call AddWave(18, 20000, 10)
        Call AddWave(17, 35000, 10)
        Call AddWave(17, 60000, 20)
        
    End With
    
    ReDim Provação(2).Wave(1 To 6) As WaveRec
    
    With Provação(2)
        .Map = 41
        .X = 12
        .Y = 12
        .Cost = 35000
        .MinLevel = 40
        .RewardXP = 90000
        .TradeItem = 42
        .RewardItem = 43
        
        Call AddWave(17, 5000, 15, 2)
        Call AddWave(17, 20000, 20)
        Call AddWave(17, 35000, 30)
        Call AddWave(19, 60000, 20)
        Call AddWave(20, 90000, 10)
        Call AddWave(21, 130000, 1)
        
    End With
    
    ReDim Provação(3).Wave(1 To 4) As WaveRec
    
    With Provação(3)
        .Map = 42
        .X = 12
        .Y = 12
        .Cost = 60000
        .MinLevel = 60
        .RewardXP = 200000
        .TradeItem = 43
        .RewardItem = 44
        
        Call AddWave(20, 5000, 15, 3)
        Call AddWave(20, 15000, 30)
        Call AddWave(21, 30000, 1)
        Call AddWave(16, 60000, 1)
        
    End With
    
    ProvaçãoCount = UBound(Provação)
    
End Sub

Sub AddWave(ByVal EnemyNum As Long, ByVal Interval As Long, Optional Quant As Byte = 1, Optional prov As Byte = 0)
    Static ProvNum As Byte
    If prov <> 0 Then ProvNum = prov
    
    Dim i As Long, n As Long, WaveNum As Long
    
    For i = 1 To UBound(Provação(ProvNum).Wave)
        If Provação(ProvNum).Wave(i).WaveTimer = 0 Then
            WaveNum = i
            Exit For
        End If
    Next i
    
    If WaveNum > 0 Then
        Provação(ProvNum).Wave(WaveNum).WaveTimer = Interval
        n = 0
        i = 1
        Do While i <= MAX_MAP_NPCS And n < Quant
            If Provação(ProvNum).Wave(WaveNum).Enemy(i).Num = 0 Then
                Provação(ProvNum).Wave(WaveNum).Enemy(i).Num = EnemyNum
                n = n + 1
            End If
            i = i + 1
        Loop
    Else
        Call SetStatus("Erro ao adicionar onda á provação " & prov & "!.")
    End If
    
End Sub

Sub LoadProvation(ByVal ProvNum As Long)
    Dim F As Long
    Dim filename As String
    filename = App.path & "\data\provation\" & ProvNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Get #F, , Provação(ProvNum)
    Close #F
End Sub

Sub LoadWishes()
    Dim i As Long, n As Long
    Dim filename As String
    
    Call SetStatus("Loading wishes...")
    
    filename = App.path & "\data\wishes.ini"
    
    n = 1
    Do While GetVar(filename, "WISHES", "Invoke" & n) <> ""
        n = n + 1
    Loop
    n = n - 1
    
    If n > 0 Then
        ReDim Wish(1 To n) As WishRec
        For i = 1 To n
            Wish(i).Phrase = LCase(GetVar(filename, "WISHES", "Invoke" & i))
            Wish(i).Event = Val(GetVar(filename, "WISHES", "Event" & i))
            Wish(i).Type = Val(GetVar(filename, "WISHES", "Type" & i))
            Wish(i).Item = Val(GetVar(filename, "WISHES", "Item" & i))
            Wish(i).ItemVal = Val(GetVar(filename, "WISHES", "ItemVal" & i))
        Next i
    Else
        ReDim Wish(1 To 1) As WishRec
    End If
End Sub

Sub LoadNPCBase()
    Dim rs As ADODB.RecordSet
    Dim SQL As String
    
    ReDim NPCBase(1 To MAX_LEVELS)
    
    If Not UZ Then Exit Sub
    
    SQL = "SELECT * FROM tb_npc_base"
    Query rs, SQL
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            NPCBase(rs.Fields("level")).Damage = Val(rs.Fields("damage"))
            NPCBase(rs.Fields("level")).HP = Val(rs.Fields("hp"))
            NPCBase(rs.Fields("level")).Exp = Val(rs.Fields("exp"))
            NPCBase(rs.Fields("level")).Acc = Val(rs.Fields("acc"))
            NPCBase(rs.Fields("level")).Esq = Val(rs.Fields("esq"))
            rs.MoveNext
        Loop
    End If
    
End Sub

Sub LoadDailyMission()
    Dim Total As Long
    Dim filename As String
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim n As Long
    
    filename = App.path & "\data\daily.ini"
    Total = Val(GetVar(filename, "INIT", "MaxDaily"))
    
    If Total < 1 Then
        ReDim DailyMission(1 To 1)
    Else
        ReDim DailyMission(1 To Total)
        
        Dim i As Long
        For i = 1 To Total
            DailyMission(i).Description = GetVar(filename, "DAILY" & i, "Description")
            DailyMission(i).NumberFactory = GetVar(filename, "DAILY" & i, "NumberFactory")
            
            tmpSprite = GetVar(filename, "DAILY" & i, "Subtype")
            
            If tmpSprite <> vbNullString Then
                tmpArray() = Split(tmpSprite, ",")
                ReDim DailyMission(i).Subtype(0 To UBound(tmpArray))
                For n = 0 To UBound(tmpArray)
                    DailyMission(i).Subtype(n) = tmpArray(n)
                Next n
            End If
        Next i
    End If
End Sub
Sub UpdatePlayer(ByVal Index As Long, ByVal UpdateNum As Long)
    Dim i As Long
    Select Case UpdateNum
    
        Case 1 'Retrocedendo a distribuição de pontos para o método anterior
            For i = 1 To Stats.Stat_Count - 1
                Player(Index).statPoints(i) = 0
                Player(Index).stat(i) = 1
            Next i
            Player(Index).Points = Player(Index).Level * 3
            Player(Index).Version = UpdateNum
        
    End Select
End Sub

Sub LoadConquistas()
    Dim ConquistaCount As Long
    Dim filename As String
    
    filename = App.path & "\data\Conquistas.ini"
    
    ConquistaCount = 1
    Do While GetVar(filename, "CONQUISTA" & ConquistaCount, "Name") <> vbNullString
        ConquistaCount = ConquistaCount + 1
    Loop
    ConquistaCount = ConquistaCount - 1
    
    If ConquistaCount > 0 Then
        ReDim Conquistas(1 To ConquistaCount) As ConquistaRec
        
        Dim i As Long
        For i = 1 To ConquistaCount
            Conquistas(i).Name = GetVar(filename, "CONQUISTA" & i, "Name")
            Conquistas(i).Desc = GetVar(filename, "CONQUISTA" & i, "Desc")
            Conquistas(i).Exp = Val(GetVar(filename, "CONQUISTA" & i, "Exp"))
            Conquistas(i).Progress = Val(GetVar(filename, "CONQUISTA" & i, "Progress"))
            
            Dim n As Long
            For n = 1 To 5
                Conquistas(i).Reward(n).Num = Val(GetVar(filename, "CONQUISTA" & i, "Reward" & n))
                Conquistas(i).Reward(n).Value = Val(GetVar(filename, "CONQUISTA" & i, "Value" & n))
            Next n
            
        Next i
    End If
End Sub
Sub LoadEspAmount()
    Dim i As Long
    For i = 1 To 3
        EspAmount(i) = Val(GetVar(App.path & "\data\EspAmount.ini", "AMOUNT", "Esp" & i))
    Next i
End Sub

Sub SaveEspAmount()
    Dim i As Long
    For i = 1 To 3
        Call PutVar(App.path & "\data\EspAmount.ini", "AMOUNT", "Esp" & i, Val(EspAmount(i)))
    Next i
End Sub
