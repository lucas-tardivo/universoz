Attribute VB_Name = "modClientTCP"
Option Explicit
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer

Sub TcpInit(Optional GetPìngs As Boolean = True)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    If Options.Debug = 2 Then On Error Resume Next
    
    Set PlayerBuffer = New clsBuffer
    
    'Pings
    If GetPìngs = True Then
        Dim i As Long, Tick As Long
        For i = 1 To UBound(Options.Servers)
            Tick = GetTickCount
            
            frmMain.Socket.RemoteHost = Options.Servers(i).IP
            frmMain.Socket.RemotePort = Options.Port
            frmMain.Socket.Connect
            
            Do While frmMain.Socket.State <> sckConnected And Tick + 3000 > GetTickCount
                DoEvents
            Loop
            
            Options.Servers(i).Ping = GetTickCount - Tick
            frmMain.Socket.Close
        Next i
    End If

    ' connect
    frmMain.Socket.Close
    frmMain.Socket.RemoteHost = Options.Servers(SelectedServer).IP
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    frmMain.Socket.Close
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyTCP", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    frmMain.Socket.GetData buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes buffer()
    
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

Public Function ConnectToServer(ByVal i As Long) As Boolean
Dim Wait As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMain.Socket.Close
    frmMain.Socket.Connect
    
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
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

Sub SendData(ByRef data() As Byte)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If IsConnected Then
        Set buffer = New clsBuffer
        buffer.WriteLong (UBound(data) - LBound(data)) + 1
        buffer.WriteBytes data()
        
        frmMain.Socket.SendData buffer.ToArray()
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
Public Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CNewAccount
    buffer.WriteString Name
    buffer.WriteString Password
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendNewAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDelAccount
    buffer.WriteString Name
    buffer.WriteString Password
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDelAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CLogin
    buffer.WriteString Name
    buffer.WriteString Password
    buffer.WriteLong App.Major
    buffer.WriteLong App.Minor
    buffer.WriteLong App.Revision
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendLogin", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Sprite As Long, ByVal Hair As Byte)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAddChar
    buffer.WriteString Name
    buffer.WriteLong Sex
    buffer.WriteLong ClassNum
    buffer.WriteLong Sprite
    buffer.WriteByte Hair
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAddChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SayMsg(ByVal Text As String)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    If InTutorial Then
        If Player(MyIndex).InTutorial = 0 And TutorialStep = 7 Then
            If LCase(Text) = "bom-dia" Then
                TutorialStep = 8
                Call PlaySound("Success2.wav", -1, -1)
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SayMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BroadcastMsg(ByVal Text As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBroadcastMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    If InTutorial Then
        If Player(MyIndex).InTutorial = 0 And TutorialStep = 8 Then
            If LCase(Text) = "bom-dia" Then
                TutorialStep = 9
                Call PlaySound("Success2.wav", -1, -1)
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BroadcastMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EmoteMsg(ByVal Text As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CEmoteMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EmoteMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMsg
    buffer.WriteString MsgTo
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerMove()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMove
    buffer.WriteLong GetPlayerDir(MyIndex)
    buffer.WriteLong TempPlayer(MyIndex).moving
    buffer.WriteLong Player(MyIndex).X
    buffer.WriteLong Player(MyIndex).Y
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    If Player(MyIndex).InTutorial = 0 And TutorialStep = 14 Then
        TutorialProgress = TutorialProgress + 1
    End If
    
    If Player(MyIndex).InTutorial = 0 And TutorialStep = 15 And ShiftDown = True Then
        TutorialProgress = TutorialProgress + 1
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerMove", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerDir()
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerDir
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerDir", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerRequestNewMap()
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNewMap
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
    
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
Dim i As Long, z As Long, w As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    CanMoveNow = False
    With Map
        buffer.WriteLong CMapData
        buffer.WriteString Trim$(.Name)
        buffer.WriteString Trim$(.Music)
        buffer.WriteString Trim$(.BGS)
        buffer.WriteByte .Moral
        buffer.WriteLong .Up
        buffer.WriteLong .Down
        buffer.WriteLong .Left
        buffer.WriteLong .Right
        buffer.WriteLong .BootMap
        buffer.WriteByte .BootX
        buffer.WriteByte .BootY
        
        buffer.WriteLong Map.Weather
        buffer.WriteLong Map.WeatherIntensity
        
        buffer.WriteLong Map.Fog
        buffer.WriteLong Map.FogSpeed
        buffer.WriteLong Map.FogOpacity
        buffer.WriteLong Map.FogDir
        
        buffer.WriteLong Map.Red
        buffer.WriteLong Map.Green
        buffer.WriteLong Map.Blue
        buffer.WriteLong Map.Alpha
        
        buffer.WriteByte .MaxX
        buffer.WriteByte .MaxY
        
        buffer.WriteByte .Fly
        buffer.WriteByte .Ambiente
    End With

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY

            With Map.Tile(X, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(i).X
                    buffer.WriteLong .Layer(i).Y
                    buffer.WriteLong .Layer(i).Tileset
                Next
                For z = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Autotile(z)
                Next
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteString .Data4
                buffer.WriteByte .DirBlock
            End With

        Next
    Next

    With Map

        For X = 1 To MAX_MAP_NPCS
            buffer.WriteLong .Npc(X)
            buffer.WriteLong .NpcSpawnType(X)
        Next
        buffer.WriteLong .Panorama

    End With

    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpMeTo(ByVal Name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpMeTo
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpMeTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpToMe(ByVal Name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpToMe
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarptoMe", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpTo(ByVal mapnum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpTo
    buffer.WriteLong mapnum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetAccess
    buffer.WriteString Name
    buffer.WriteLong Access
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSetAccess", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetSprite
    buffer.WriteLong SpriteNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSetSprite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendKick(ByVal Name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CKickPlayer
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendKick", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBan(ByVal Name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBanPlayer
    buffer.WriteString Name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBan", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanList()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBanList
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBanList", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapRespawn()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CMapRespawn
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapRespawn", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseItem(ByVal invNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUseItem
    buffer.WriteLong invNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUseItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDropItem(ByVal invNum As Long, ByVal Amount As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If InBank Or InShop Then Exit Sub
    
    ' do basic checks
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    If PlayerInv(invNum).Num < 1 Or PlayerInv(invNum).Num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
        If Amount < 1 Or Amount > PlayerInv(invNum).Value Then Exit Sub
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong CMapDropItem
    buffer.WriteLong invNum
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDropItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendWhosOnline()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWhosOnline
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendWhosOnline", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetMotd
    buffer.WriteString MOTD
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMOTDChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanDestroy()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBanDestroy
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBanDestroy", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSwapInvSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSwapSpellSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub GetPing()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    PingStart = GetTickCount
    Set buffer = New clsBuffer
    buffer.WriteLong CCheckPing
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GetPing", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUnequip(ByVal eqNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUnequip
    buffer.WriteLong eqNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUnequip", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTrainStat(ByVal StatNum As Byte)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUseStatPoint
    buffer.WriteByte StatNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Public Sub BuyItem(ByVal shopSlot As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBuyItem
    buffer.WriteLong shopSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BuyItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SellItem(ByVal invSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSellItem
    buffer.WriteLong invSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SellItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DepositItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDepositItem
    buffer.WriteLong invSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DepositItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWithdrawItem
    buffer.WriteLong bankslot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WithdrawItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseBank()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCloseBank
    SendData buffer.ToArray()
    Set buffer = Nothing
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
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CChangeBankSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChangeBankSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AcceptTrade()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTrade
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AcceptTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DeclineTrade()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTrade
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DeclineTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub TradeItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeItem
    buffer.WriteLong invSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UntradeItem(ByVal invSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUntradeItem
    buffer.WriteLong invSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UntradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarChange
    buffer.WriteLong sType
    buffer.WriteLong Slot
    buffer.WriteLong hotbarNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbarChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
Dim buffer As clsBuffer, X As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
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

    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarUse
    buffer.WriteLong Slot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbarUse", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptTradeRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAcceptTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineTradeRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyLeave()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyLeave
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyLeave", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptParty()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptParty
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAcceptParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineParty()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineParty
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Events_SendSaveEvent(ByVal EIndex As Long)
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Dim buffer As clsBuffer
    Dim i As Long, d As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong CSaveEventData
    buffer.WriteLong EIndex
    buffer.WriteString Events(EIndex).Name
    buffer.WriteByte Events(EIndex).chkSwitch
    buffer.WriteByte Events(EIndex).chkVariable
    buffer.WriteByte Events(EIndex).chkHasItem
    buffer.WriteLong Events(EIndex).SwitchIndex
    buffer.WriteByte Events(EIndex).SwitchCompare
    buffer.WriteLong Events(EIndex).VariableIndex
    buffer.WriteByte Events(EIndex).VariableCompare
    buffer.WriteLong Events(EIndex).VariableCondition
    buffer.WriteLong Events(EIndex).HasItemIndex
    If Events(EIndex).HasSubEvents Then
        buffer.WriteLong UBound(Events(EIndex).SubEvents)
        For i = 1 To UBound(Events(EIndex).SubEvents)
            With Events(EIndex).SubEvents(i)
                buffer.WriteLong .Type
                If .HasText Then
                    buffer.WriteLong UBound(.Text)
                    For d = 1 To UBound(.Text)
                        buffer.WriteString .Text(d)
                    Next d
                Else
                    buffer.WriteLong 0
                End If
                If .HasData Then
                    buffer.WriteLong UBound(.data)
                    For d = 1 To UBound(.data)
                        buffer.WriteLong .data(d)
                    Next d
                Else
                    buffer.WriteLong 0
                End If
            End With
        Next i
    Else
        buffer.WriteLong 0
    End If
    
    buffer.WriteByte Events(EIndex).Trigger
    buffer.WriteByte Events(EIndex).WalkThrought
    
    SendData buffer.ToArray

    Set buffer = Nothing
End Sub

Public Sub Events_SendRequestEditEvents()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditEvents
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
Public Sub Events_SendRequestEventsData()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEventsData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
Public Sub Events_SendRequestEventData(ByVal i As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEventData
    buffer.WriteLong i
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
Public Sub Events_SendChooseEventOption(ByVal i As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CChooseEventOption
    buffer.WriteLong i
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditEffect()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditEffect
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditEffect", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveEffect(ByVal EffectNum As Long)
Dim buffer As clsBuffer
Dim EffectSize As Long
Dim EffectData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    EffectSize = LenB(Effect(EffectNum))
    ReDim EffectData(EffectSize - 1)
    CopyMemory EffectData(0), ByVal VarPtr(Effect(EffectNum)), EffectSize
    buffer.WriteLong CSaveEffect
    buffer.WriteLong EffectNum
    buffer.WriteBytes EffectData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveEffect", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Sub SendRequestEffects()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEffects
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEffects", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerTarget(ByVal Target As Long, ByVal TargetType As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If myTargetType = TargetType And myTarget = Target Then
        myTargetType = 0
        myTarget = 0
    Else
        myTarget = Target
        myTargetType = TargetType
        
        If myTargetType = TARGET_TYPE_NPC Then
            If Npc(MapNpc(myTarget).Num).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER And Len(Trim$(Npc(MapNpc(myTarget).Num).AttackSay)) > 0 Then
                AddChatBubble myTarget, TARGET_TYPE_NPC, Npc(MapNpc(myTarget).Num).AttackSay, Black
            End If
        End If
    End If
    
    If ScouterOn Then PlaySound "checando pdl.mp3", -1, -1

    Set buffer = New clsBuffer
    buffer.WriteLong CTarget
    buffer.WriteLong Target
    buffer.WriteLong TargetType
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerTarget", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDevSuite()
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDevSuite
    buffer.WriteByte 0
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDevSuite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendOnDeath()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong COnDeath
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendQuestInfo(QuestNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CQuestInfo
    buffer.WriteLong QuestNum
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptSell()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSellPlanet
    buffer.WriteByte 1
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineSell()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSellPlanet
    buffer.WriteByte 0
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendEnterGravity()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CEnterGravity
    buffer.WriteLong SelectedGravity
    buffer.WriteLong SelectedHours
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendCompleteTutorial()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCompleteTutorial
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    TutorialBlockWalk = False
    
    Call PlaySound("Success2.wav", -1, -1)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAttack()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAttack
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    TempPlayer(MyIndex).Attacking = 1
    TempPlayer(MyIndex).AttackTimer = GetTickCount
    TempPlayer(MyIndex).AttackAnim = Rand(0, 1)
    
    Dim i As Long
    If Not InOwnPlanet Then
    For i = 1 To MAX_MAP_NPCS
        Dim NpcX As Long, NpcY As Long
        ' Check if at same coordinates
        Select Case GetPlayerDir(MyIndex)
            Case DIR_UP
                NpcX = MapNpc(i).X
                NpcY = MapNpc(i).Y + 1
            Case DIR_DOWN
                NpcX = MapNpc(i).X
                NpcY = MapNpc(i).Y - 1
            Case DIR_LEFT
                NpcX = MapNpc(i).X + 1
                NpcY = MapNpc(i).Y
            Case DIR_RIGHT
                NpcX = MapNpc(i).X - 1
                NpcY = MapNpc(i).Y
        End Select
        
        If NpcX = GetPlayerX(MyIndex) Then
            If NpcY = GetPlayerY(MyIndex) Then
                If MapNpc(i).Num > 0 Then
                If Npc(MapNpc(i).Num).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(MapNpc(i).Num).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
                    Dim AnimationNum As Long
                    AnimationNum = PlayerAttackAnim
                    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                        AnimationNum = Item(GetPlayerEquipment(MyIndex, Weapon)).Animation
                    End If
                    Call DoAnimation(AnimationNum, NpcX, NpcY, TARGET_TYPE_NPC, i, GetPlayerDir(MyIndex))
                    
                    If myTarget <> i And myTargetType <> TARGET_TYPE_NPC Then
                        myTarget = i
                        myTargetType = TARGET_TYPE_NPC
                    End If
                    MapNpc(i).Vital(Vitals.HP) = MapNpc(i).Vital(Vitals.HP) - 3
                    If MapNpc(i).Vital(Vitals.HP) < 1 Then MapNpc(i).Vital(Vitals.HP) = 1
                End If
                End If
            End If
        End If
    Next i
    End If
End Sub
Sub SendCreateGuild(ByVal GuildName As String, ByRef IconColor() As Byte)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCreateGuild
    buffer.WriteString GuildName
    Dim i As Long
    For i = 0 To 24
        buffer.WriteByte IconColor(i)
    Next i
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendGuildMOTD(ByVal MOTD As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildAction
    buffer.WriteByte 1
    buffer.WriteString MOTD
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendGuildInvite(ByVal PlayerName As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildAction
    buffer.WriteByte 2
    buffer.WriteString PlayerName
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendGuildInviteAnswer(ByVal Answer As Byte)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildAction
    buffer.WriteByte 3
    buffer.WriteByte Answer
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendGuildRevoke(ByVal MemberIndex As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildAction
    buffer.WriteByte 4
    buffer.WriteLong MemberIndex
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendGuildPromote(ByVal MemberIndex As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildAction
    buffer.WriteByte 5
    buffer.WriteLong MemberIndex
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendGuildKick(ByVal MemberIndex As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildAction
    buffer.WriteByte 6
    buffer.WriteLong MemberIndex
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendGuildLeave()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildAction
    buffer.WriteByte 7
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendGuildDonate(ByVal DonationType As Byte, ByVal Quant As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildAction
    buffer.WriteByte 8
    buffer.WriteByte DonationType
    buffer.WriteLong Quant
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendGuildUpBlock(ByVal UpBlock As Byte)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CGuildAction
    buffer.WriteByte 9
    buffer.WriteByte UpBlock
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendArenaDeny()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CChallengeArena
    buffer.WriteByte 2
    buffer.WriteByte 0
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendArenaAccept()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CChallengeArena
    buffer.WriteByte 2
    buffer.WriteByte 1
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRemoveBlock()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlanetChange
    buffer.WriteByte 0
    buffer.WriteLong EditTargetX
    buffer.WriteLong EditTargetY
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMoveBlock(ByVal X As Long, ByVal Y As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlanetChange
    buffer.WriteByte 3
    buffer.WriteLong EditTargetX
    buffer.WriteLong EditTargetY
    buffer.WriteLong X
    buffer.WriteLong Y
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendEvolute()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlanetChange
    buffer.WriteByte 4
    buffer.WriteLong EditTargetX
    buffer.WriteLong EditTargetY
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAccelerate()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlanetChange
    buffer.WriteByte 2
    buffer.WriteLong EditTargetX
    buffer.WriteLong EditTargetY
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAntiHackData(DLLFound As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAntiHackData
    If DLLFound = vbNullString Then DLLFound = "Nenhuma encontrada!"
    buffer.WriteString DLLFound
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlanetName()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlanetChange
    buffer.WriteByte 1
    buffer.WriteString sDialogue
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendOpenBuilding()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlanetChange
    buffer.WriteByte 5
    buffer.WriteLong EditTargetX
    buffer.WriteLong EditTargetY
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendProduceSementes(ByVal Nivel As Byte, ByVal Quant As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlanetChange
    buffer.WriteByte 6
    buffer.WriteLong EditTargetX
    buffer.WriteLong EditTargetY
    buffer.WriteByte Nivel
    buffer.WriteByte Quant
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcelerar()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlanetChange
    buffer.WriteByte 7
    buffer.WriteLong EditTargetX
    buffer.WriteLong EditTargetY
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSellEspeciaria(ByVal Number As Byte, ByVal Quant As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSellEsp
    buffer.WriteByte Number
    buffer.WriteLong Quant
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long, Optional ItemTo As String = "Drop", Optional Msg As Byte = 0)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSpawnItem
    buffer.WriteLong tmpItem
    buffer.WriteLong tmpAmount
    buffer.WriteString ItemTo
    buffer.WriteByte Msg
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSpawnItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSupportMsg(ByVal Msg As String, Optional ToPlayer As String = "admin")
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSupport
    buffer.WriteString Msg
    buffer.WriteString ToPlayer
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
