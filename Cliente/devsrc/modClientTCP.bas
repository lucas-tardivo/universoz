Attribute VB_Name = "modClientTCP"
Option Explicit
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer

Sub TcpInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set PlayerBuffer = New clsBuffer

    ' connect
    frmMain.Socket.RemoteHost = Options.IP
    frmMain.Socket.RemotePort = Options.Port

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TcpInit", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DestroyTCP()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.Socket.Close
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyTCP", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.Socket.GetData Buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes Buffer()
    
    If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "IncomingData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConnectToServer(ByVal I As Long) As Boolean
Dim Wait As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMain.Socket.Close
    frmMain.Socket.Connect
    
    SetStatus "Connecting to server..."
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop
    
    ConnectToServer = IsConnected

    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConnectToServer", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsConnected() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If frmMain.Socket.State = sckConnected Then
        IsConnected = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsConnected", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlaying", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SendData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If IsConnected Then
        Set Buffer = New clsBuffer
                
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data()
        frmMain.Socket.SendData Buffer.ToArray()
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************

Public Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CLogin
    Buffer.WriteString Name
    Buffer.WriteString Password
    Buffer.WriteLong App.Major
    Buffer.WriteLong App.Minor
    Buffer.WriteLong App.Revision
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendLogin", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SayMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SayMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BroadcastMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBroadcastMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BroadcastMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EmoteMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CEmoteMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EmoteMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMsg
    Buffer.WriteString MsgTo
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerMove()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong GetPlayerDir(MyIndex)
    Buffer.WriteLong TempPlayer(MyIndex).Moving
    Buffer.WriteLong Player(MyIndex).X
    Buffer.WriteLong Player(MyIndex).Y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerMove", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerDir()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerDir
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerDir", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerRequestNewMap()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNewMap
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerRequestNewMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMap()
Dim packet As String
Dim X As Long
Dim Y As Long
Dim I As Long, Z As Long, w As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    CanMoveNow = False
    With Map
        Buffer.WriteLong CMapData
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteString Trim$(.Music)
        Buffer.WriteString Trim$(.BGS)
        Buffer.WriteByte .Moral
        Buffer.WriteLong .Up
        Buffer.WriteLong .Down
        Buffer.WriteLong .Left
        Buffer.WriteLong .Right
        Buffer.WriteLong .BootMap
        Buffer.WriteByte .BootX
        Buffer.WriteByte .BootY
        
        Buffer.WriteLong Map.Weather
        Buffer.WriteLong Map.WeatherIntensity
        
        Buffer.WriteLong Map.Fog
        Buffer.WriteLong Map.FogSpeed
        Buffer.WriteLong Map.FogOpacity
        Buffer.WriteLong Map.FogDir
        
        Buffer.WriteLong Map.Red
        Buffer.WriteLong Map.Green
        Buffer.WriteLong Map.Blue
        Buffer.WriteLong Map.Alpha
        
        Buffer.WriteByte .MaxX
        Buffer.WriteByte .MaxY
        
        Buffer.WriteByte .Fly
        Buffer.WriteByte .Ambiente
    End With

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY

            With Map.Tile(X, Y)
                For I = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .layer(I).X
                    Buffer.WriteLong .layer(I).Y
                    Buffer.WriteLong .layer(I).Tileset
                Next
                For Z = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Autotile(Z)
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteString .Data4
                Buffer.WriteByte .DirBlock
            End With

        Next
    Next

    With Map

        For X = 1 To MAX_MAP_NPCS
            Buffer.WriteLong .Npc(X)
            Buffer.WriteLong .NpcSpawnType(X)
        Next
        Buffer.WriteLong .Panorama

    End With

    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpMeTo(ByVal Name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpMeTo
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpMeTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpToMe(ByVal Name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpToMe
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarptoMe", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpTo
    Buffer.WriteLong MapNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetAccess
    Buffer.WriteString Name
    Buffer.WriteLong Access
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSetAccess", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetSprite
    Buffer.WriteLong SpriteNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSetSprite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendKick(ByVal Name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CKickPlayer
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendKick", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBan(ByVal Name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanPlayer
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBan", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanList()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanList
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBanList", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditItem()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditItem
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveItem(ByVal itemNum As Long)
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemNum)), ItemSize
    Buffer.WriteLong CSaveItem
    Buffer.WriteLong itemNum
    Buffer.WriteBytes ItemData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditAnimation()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditAnimation
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
Dim Buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    Buffer.WriteLong CSaveAnimation
    Buffer.WriteLong Animationnum
    Buffer.WriteBytes AnimationData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditNpc()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditNpc
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveNpc(ByVal npcNum As Long)
Dim Buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    NpcSize = LenB(Npc(npcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(npcNum)), NpcSize
    Buffer.WriteLong CSaveNpc
    Buffer.WriteLong npcNum
    Buffer.WriteBytes NpcData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSavequest(ByVal QuestNum As Long)
Dim Buffer As clsBuffer
Dim questSize As Long
Dim questData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    questSize = LenB(Quest(QuestNum))
    ReDim questData(questSize - 1)
    CopyMemory questData(0), ByVal VarPtr(Quest(QuestNum)), questSize
    Buffer.WriteLong CSavequest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes questData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSavequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditResource()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditResource
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
Dim Buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    Buffer.WriteLong CSaveResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapRespawn()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapRespawn
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapRespawn", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseItem(ByVal invNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseItem
    Buffer.WriteLong invNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUseItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDropItem(ByVal invNum As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InBank Or InShop Then Exit Sub
    
    ' do basic checks
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    If PlayerInv(invNum).Num < 1 Or PlayerInv(invNum).Num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
        If Amount < 1 Or Amount > PlayerInv(invNum).Value Then Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapDropItem
    Buffer.WriteLong invNum
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDropItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendWhosOnline()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWhosOnline
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendWhosOnline", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetMotd
    Buffer.WriteString MOTD
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMOTDChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditShop()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveShop(ByVal shopnum As Long)
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopnum)), ShopSize
    Buffer.WriteLong CSaveShop
    Buffer.WriteLong shopnum
    Buffer.WriteBytes ShopData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditSpell()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditSpell
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    Buffer.WriteLong CSaveSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditMap()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanDestroy()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanDestroy
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBanDestroy", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapInvSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapSpellSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub GetPing()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PingStart = GetTickCount
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCheckPing
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GetPing", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUnequip(ByVal eqNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUnequip
    Buffer.WriteLong eqNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUnequip", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestPlayerData()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPlayerData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestPlayerData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestItems()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestItems
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestItems", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestAnimations()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestAnimations
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestAnimations", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestNPCS()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNPCS
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestNPCS", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestResources()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestResources
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestResources", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestSpells()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestSpells
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestSpells", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestShops()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestShops
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestShops", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long, Optional ItemTo As String = "Drop", Optional Msg As Byte = 0)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSpawnItem
    Buffer.WriteLong tmpItem
    Buffer.WriteLong tmpAmount
    Buffer.WriteString ItemTo
    Buffer.WriteByte Msg
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSpawnItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTrainStat(ByVal StatNum As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseStatPoint
    Buffer.WriteByte StatNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub BuyItem(ByVal shopSlot As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBuyItem
    Buffer.WriteLong shopSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BuyItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SellItem(ByVal invSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSellItem
    Buffer.WriteLong invSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SellItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DepositItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDepositItem
    Buffer.WriteLong invSlot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DepositItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWithdrawItem
    Buffer.WriteLong bankslot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WithdrawItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseBank()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseBank
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    InBank = False
    GUIWindow(GUI_BANK).visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CloseBank", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChangeBankSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CChangeBankSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChangeBankSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AcceptTrade()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AcceptTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DeclineTrade()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DeclineTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub TradeItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeItem
    Buffer.WriteLong invSlot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UntradeItem(ByVal invSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUntradeItem
    Buffer.WriteLong invSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UntradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarChange
    Buffer.WriteLong sType
    Buffer.WriteLong Slot
    Buffer.WriteLong hotbarNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbarChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
Dim Buffer As clsBuffer, X As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check if spell
    If Hotbar(Slot).sType = 2 Then ' spell
        For X = 1 To MAX_PLAYER_SPELLS
            ' is the spell matching the hotbar?
            If PlayerSpells(X) = Hotbar(Slot).Slot Then
                ' found it, cast it
                CastSpell X
                Exit Sub
            End If
        Next
        ' can't find the spell, exit out
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarUse
    Buffer.WriteLong Slot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbarUse", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapReport()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapReport
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapReport", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptTradeRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAcceptTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineTradeRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyLeave()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyLeave
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyLeave", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptParty()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAcceptParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineParty()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Events_SendSaveEvent(ByVal EIndex As Long)
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Dim Buffer As clsBuffer
    Dim I As Long, d As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CSaveEventData
    Buffer.WriteLong EIndex
    Buffer.WriteString Events(EIndex).Name
    Buffer.WriteByte Events(EIndex).chkSwitch
    Buffer.WriteByte Events(EIndex).chkVariable
    Buffer.WriteByte Events(EIndex).chkHasItem
    Buffer.WriteLong Events(EIndex).SwitchIndex
    Buffer.WriteByte Events(EIndex).SwitchCompare
    Buffer.WriteLong Events(EIndex).VariableIndex
    Buffer.WriteByte Events(EIndex).VariableCompare
    Buffer.WriteLong Events(EIndex).VariableCondition
    Buffer.WriteLong Events(EIndex).HasItemIndex
    If Events(EIndex).HasSubEvents Then
        Buffer.WriteLong UBound(Events(EIndex).SubEvents)
        For I = 1 To UBound(Events(EIndex).SubEvents)
            With Events(EIndex).SubEvents(I)
                Buffer.WriteLong .Type
                If .HasText Then
                    Buffer.WriteLong UBound(.Text)
                    For d = 1 To UBound(.Text)
                        Buffer.WriteString .Text(d)
                    Next d
                Else
                    Buffer.WriteLong 0
                End If
                If .HasData Then
                    Buffer.WriteLong UBound(.Data)
                    For d = 1 To UBound(.Data)
                        Buffer.WriteLong .Data(d)
                    Next d
                Else
                    Buffer.WriteLong 0
                End If
            End With
        Next I
    Else
        Buffer.WriteLong 0
    End If
    
    Buffer.WriteByte Events(EIndex).Trigger
    Buffer.WriteByte Events(EIndex).WalkThrought
    
    SendData Buffer.ToArray

    Set Buffer = Nothing
End Sub

Public Sub Events_SendRequestEditEvents()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditEvents
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Public Sub Events_SendRequestEventsData()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEventsData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Public Sub Events_SendRequestEventData(ByVal I As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEventData
    Buffer.WriteLong I
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Public Sub Events_SendChooseEventOption(ByVal I As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CChooseEventOption
    Buffer.WriteLong I
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub RequestSwitchesAndVariables()
Dim I As Long, Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CRequestSwitchesAndVariables
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Sub SendSwitchesAndVariables()
Dim I As Long, Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwitchesAndVariables
    For I = 1 To MAX_SWITCHES
        Buffer.WriteString Switches(I)
    Next
    For I = 1 To MAX_VARIABLES
        Buffer.WriteString Variables(I)
    Next
    SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Public Sub SendRequestEditEffect()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditEffect
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditEffect", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveEffect(ByVal EffectNum As Long)
Dim Buffer As clsBuffer
Dim EffectSize As Long
Dim EffectData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    EffectSize = LenB(Effect(EffectNum))
    ReDim EffectData(EffectSize - 1)
    CopyMemory EffectData(0), ByVal VarPtr(Effect(EffectNum)), EffectSize
    Buffer.WriteLong CSaveEffect
    Buffer.WriteLong EffectNum
    Buffer.WriteBytes EffectData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveEffect", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Sub SendRequestEffects()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEffects
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEffects", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerTarget(ByVal Target As Long, ByVal TargetType As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If myTargetType = TargetType And myTarget = Target Then
        myTargetType = 0
        myTarget = 0
    Else
        myTarget = Target
        myTargetType = TargetType
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTarget
    Buffer.WriteLong Target
    Buffer.WriteLong TargetType
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    If ScouterOn Then PlaySound "checando pdl.mp3", -1, -1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerTarget", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendEditNews(ByVal Text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CEditNews
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Public Sub SendRequestEditNews()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditNews
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Public Sub SendRequestNews()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNews
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDevSuite()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDevSuite
    Buffer.WriteByte 1
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDevSuite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditQuest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditQuest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendQuestInfo(QuestNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuestInfo
    Buffer.WriteLong QuestNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

