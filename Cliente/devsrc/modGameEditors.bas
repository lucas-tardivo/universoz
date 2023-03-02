Attribute VB_Name = "modGameEditors"
Option Explicit
Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
Dim i As Long
Dim smusic() As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set the width
    frmEditor_Map.Width = 9825
    
    ' we're in the map editor
    InMapEditor = True
    
    ' show the form
    frmEditor_Map.visible = True
    
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.max = NumTileSets
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.Value = 1
    
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    MapEditorTileScroll
    
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"
    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).Name
    Next
    
    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorProperties()
Dim X As Long
Dim Y As Long
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    For i = 1 To UBound(musicCache)
        frmEditor_MapProperties.lstMusic.AddItem musicCache(i)
    Next
    frmEditor_MapProperties.cmbSound.Clear
    frmEditor_MapProperties.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_MapProperties.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_MapProperties
        .txtName.Text = Trim$(Map.Name)
        
        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For i = 0 To .lstMusic.ListCount
                If .lstMusic.List(i) = Trim$(Map.Music) Then
                    .lstMusic.ListIndex = i
                End If
            Next
        End If
        
        If .cmbSound.ListCount >= 0 Then
            .cmbSound.ListIndex = 0
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Map.BGS) Then
                    .cmbSound.ListIndex = i
                End If
            Next
        End If
        
        ' rest of it
        .txtUp.Text = CStr(Map.Up)
        .txtDown.Text = CStr(Map.Down)
        .txtLeft.Text = CStr(Map.Left)
        .txtRight.Text = CStr(Map.Right)
        .cmbMoral.ListIndex = Map.Moral
        .txtBootMap.Text = CStr(Map.BootMap)
        .txtBootX.Text = CStr(Map.BootX)
        .txtBootY.Text = CStr(Map.BootY)
        
        .optAmbiente(Map.Ambiente).Value = True
        
        .chkFly.Value = Map.Fly
        
        .CmbWeather.ListIndex = Map.Weather
        .scrlWeatherIntensity.Value = Map.WeatherIntensity
        
        .ScrlFog.Value = Map.Fog
        .ScrlFogSpeed.Value = Map.FogSpeed
        .scrlFogOpacity.Value = Map.FogOpacity
        .cmbDir.ListIndex = Map.FogDir
        
        .ScrlR.Value = Map.Red
        .ScrlG.Value = Map.Green
        .ScrlB.Value = Map.Blue
        .scrlA.Value = Map.Alpha
        .scrlPanorama = Map.Panorama
        
        ' show the map npcs
        .lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS
            If Map.Npc(X) > 0 Then
            .lstNpcs.AddItem X & ": " & Trim$(Npc(Map.Npc(X)).Name)
            Else
                .lstNpcs.AddItem X & ": No NPC"
            End If
        Next
        .lstNpcs.ListIndex = 0
        
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"
        For X = 1 To MAX_NPCS
            .cmbNpc.AddItem X & ": " & Trim$(Npc(X).Name)
        Next
        
        ' set the combo box properly
        Dim tmpString() As String
        Dim npcNum As Long
        tmpString = Split(.lstNpcs.List(.lstNpcs.ListIndex))
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = Map.Npc(npcNum)
    
        ' show the current map
        .lblMap.Caption = "Mapa atual: " & GetPlayerMap(MyIndex)
        .txtMaxX.Text = Map.MaxX
        .txtMaxY.Text = Map.MaxY
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorProperties", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False, Optional ByVal theAutotile As Byte = 0)
Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If theAutotile > 0 Then
        With Map.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurLayer) = theAutotile
            CacheRenderState X, Y, CurLayer
        End With
        ' do a re-init so we can see our changes
        initAutotiles
        Exit Sub
    End If

    If Not multitile Then ' single
        With Map.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurLayer) = 0
            CacheRenderState X, Y, CurLayer
        End With
    Else ' multitile
        Y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            X2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(X, Y)
                            .Layer(CurLayer).X = EditorTileX + X2
                            .Layer(CurLayer).Y = EditorTileY + Y2
                            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                            .Autotile(CurLayer) = 0
                            CacheRenderState X, Y, CurLayer
                        End With
                    End If
                End If
                X2 = X2 + 1
            Next
            Y2 = Y2 + 1
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSetTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
Dim i As Long
Dim CurLayer As Long
Dim tmpDir As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.Value Then
            ' no autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
            Else ' multi tile!
                If frmEditor_Map.scrlAutotile.Value = 0 Then
                    MapEditorSetTile CurX, CurY, CurLayer, True
                Else
                    MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
                End If
            End If
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmEditor_Map.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = ""
                End If
                ' item spawn
                If frmEditor_Map.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' resource
                If frmEditor_Map.optResource.Value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' npc spawn
                If frmEditor_Map.optNpcSpawn.Value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' shop
                If frmEditor_Map.optShop.Value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' bank
                If frmEditor_Map.optBank.Value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' heal
                If frmEditor_Map.optHeal.Value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' trap
                If frmEditor_Map.optTrap.Value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' slide
                If frmEditor_Map.optSlide.Value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' sound
                If frmEditor_Map.optSound.Value Then
                    .Type = TILE_TYPE_SOUND
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = MapEditorSound
                End If
                ' Event
                If frmEditor_Map.optEvent.Value Then
                    .Type = TILE_TYPE_EVENT
                    .Data1 = MapEditorEventIndex
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' Arena
                If frmEditor_Map.optArena.Value Then
                    .Type = TILE_TYPE_ARENA
                    .Data1 = MapEditorArenaMap
                    .Data2 = MapEditorArenaX
                    .Data3 = MapEditorArenaY
                End If
            End With
        ElseIf frmEditor_Map.optBlock.Value Then
            If movedMouse Then Exit Sub
            ' find what tile it is
            X = X - ((X \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)
            ' see if it hits an arrow
            For i = 1 To 4
                If X >= DirArrowX(i) And X <= DirArrowX(i) + 8 Then
                    If Y >= DirArrowY(i) And Y <= DirArrowY(i) + 8 Then
                        ' flip the value.
                        setDirBlock Map.Tile(CurX, CurY).DirBlock, CByte(i), Not isDirBlocked(Map.Tile(CurX, CurY).DirBlock, CByte(i))
                        Exit Sub
                    End If
                End If
            Next
        ElseIf frmEditor_Map.optEyeDropper.Value Then
            Call MapEditorEyeDropper(CurX, CurY, CurLayer)
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.Value Then
            With Map.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).X = 0
                .Layer(CurLayer).Y = 0
                .Layer(CurLayer).Tileset = 0
                If .Autotile(CurLayer) > 0 Then
                    .Autotile(CurLayer) = 0
                    ' do a re-init so we can see our changes
                    initAutotiles
                End If
                CacheRenderState X, Y, CurLayer
            End With
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With

        End If
    End If

    CacheResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
        
        frmEditor_Map.lblTilePosition.Caption = "X: " & EditorTileX & " Y:" & EditorTileY
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub MapEditorEyeDropper(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With Map.Tile(X, Y)
        EditorTileX = .Layer(CurLayer).X
        EditorTileY = .Layer(CurLayer).Y
        If .Layer(CurLayer).Tileset > 0 Then
            frmEditor_Map.scrlTileSet.Value = .Layer(CurLayer).Tileset
        Else
            frmEditor_Map.scrlTileSet.Value = 1
        End If
    End With
    
    EditorTileWidth = 1
    EditorTileHeight = 1
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / PIC_X Then X = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / PIC_X
        If Y < 0 Then Y = 0
        If Y > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / PIC_Y Then Y = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / PIC_Y
        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorTileScroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' horizontal scrolling
    If Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width < frmEditor_Map.picBack.Width Then
        frmEditor_Map.scrlPictureX.Enabled = False
    Else
        frmEditor_Map.scrlPictureX.Enabled = True
    End If
    
    ' vertical scrolling
    If Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height < frmEditor_Map.picBack.Height Then
        frmEditor_Map.scrlPictureY.Enabled = False
    Else
        frmEditor_Map.scrlPictureY.Enabled = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorTileScroll", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSend()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SendMap
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSend", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorCancel()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong 1
    SendData Buffer.ToArray()
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearLayer()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next
    
    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Layer(CurLayer).X = 0
                Map.Tile(X, Y).Layer(CurLayer).Y = 0
                Map.Tile(X, Y).Layer(CurLayer).Tileset = 0
                CacheRenderState X, Y, CurLayer
            Next
        Next
        
        initAutotiles
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorFillLayer()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Layer(CurLayer).X = EditorTileX
                Map.Tile(X, Y).Layer(CurLayer).Y = EditorTileY
                Map.Tile(X, Y).Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                Map.Tile(X, Y).Autotile(CurLayer) = frmEditor_Map.scrlAutotile.Value
                CacheRenderState X, Y, CurLayer
            Next
        Next
        
        ' now cache the positions
        initAutotiles
    End If
    

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorFillLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearAttribs()
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, Options.Game_Name) = vbYes Then

        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Type = 0
            Next
        Next

    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearAttribs", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorLeaveMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorLeaveMap", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorPlaceRandomTile(ByVal X As Long, Y As Long)
Dim i As Long
Dim CurLayer As Long

' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub

    If frmEditor_Map.optLayers.Value Then
        If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
            MapEditorSetTile X, Y, CurLayer, , frmEditor_Map.scrlAutotile.Value
        Else ' multi tile!
            If frmEditor_Map.scrlAutotile.Value = 0 Then
                MapEditorSetTile X, Y, CurLayer, True
            Else
                MapEditorSetTile X, Y, CurLayer, , frmEditor_Map.scrlAutotile.Value
            End If
        End If
    End If

    CacheResources

' Error handler
Exit Sub
errorhandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Item.visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Item.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Item(EditorIndex)
        frmEditor_Item.txtName.Text = Trim$(.Name)
        If .Pic > frmEditor_Item.scrlPic.max Then .Pic = 0
        frmEditor_Item.scrlPic.Value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.Value = .Animation
        frmEditor_Item.scrlEffect.Value = .Effect
        frmEditor_Item.txtDesc.Text = Trim$(.Desc)
        frmEditor_Item.chkStackable.Value = .Stackable
        frmEditor_Item.chkDrop.Value = .CantDrop
        
        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Item.cmbSound.ListCount
                If frmEditor_Item.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Item.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.visible = True
            frmEditor_Item.txtDamage.Text = Trim$(.Data2)
            frmEditor_Item.txtDefence.Text = Trim$(.Data2)
            frmEditor_Item.cmbTool.ListIndex = .Data3

            If .speed < 100 Then .speed = 100
            frmEditor_Item.scrlSpeed.Value = .speed
            frmEditor_Item.scrlProjectileNum.Value = .Projectile
            frmEditor_Item.scrlProjectileRange.Value = .Range
            frmEditor_Item.scrlProjectileRotation.Value = .Rotation
            
            ' loop for stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(i).Value = .Add_Stat(i)
            Next
            
            frmEditor_Item.scrlPaperdoll = .Paperdoll
            If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
                frmEditor_Item.fraWeapon.visible = True
                frmEditor_Item.fraArmor.visible = False
            Else
                frmEditor_Item.fraWeapon.visible = False
                frmEditor_Item.fraArmor.visible = True
            End If
        Else
            frmEditor_Item.fraEquipment.visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.visible = True
            frmEditor_Item.scrlAddHp.Value = .AddHP
            frmEditor_Item.scrlAddMP.Value = .AddMP
            frmEditor_Item.scrlAddExp.Value = .AddEXP
            frmEditor_Item.scrlCastSpell.Value = .CastSpell
            frmEditor_Item.chkInstant.Value = .instaCast
        Else
            frmEditor_Item.fraVitals.visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.visible = True
            If .Data1 <= 0 Then .Data1 = 1
            frmEditor_Item.scrlSpell.Value = .Data1
            frmEditor_Item.scrlSpellQuant.Value = .Data2
        Else
            frmEditor_Item.fraSpell.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_ESOTERICA) Then
            frmEditor_Item.frmEsoterica.visible = True
            frmEditor_Item.txtBonus.Text = .EsotericaBonus
            frmEditor_Item.txtSecs.Text = .EsotericaTime
        Else
            frmEditor_Item.frmEsoterica.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_DRAGONBALL) Then
            frmEditor_Item.frmDragonball.visible = True
            If .Dragonball = 0 Then .Dragonball = 1
            frmEditor_Item.scrlDragonball.Value = .Dragonball
        Else
            frmEditor_Item.frmDragonball.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_TITULO) Then
            frmEditor_Item.scrlTituloCor = .Data1
            frmEditor_Item.chkIcon = .Data2
        Else
            frmEditor_Item.frmTitulo.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_EXTRATOR) Then
            frmEditor_Item.scrlExtratorNum.max = MAX_RESOURCES
            frmEditor_Item.scrlExtratorNum.Value = .Data2
        Else
            frmEditor_Item.frmExtrator.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_NAVE) Then
            frmEditor_Item.scrlNaveSprite.Value = .Data1
            frmEditor_Item.scrlNaveSpeed.Value = .Data2
        Else
            frmEditor_Item.frmNave.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_COMBUSTIVEL) Then
            frmEditor_Item.scrlBonus = .Data1
        Else
            frmEditor_Item.FrmCombustivel.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_VIP) Then
            frmEditor_Item.frmVIP.visible = True
            frmEditor_Item.txtDiasVIP.Text = Item(EditorIndex).Data1
        Else
            frmEditor_Item.frmVIP.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_PLANETCHANGE) Then
            frmEditor_Item.frmPlanetChange.visible = True
            frmEditor_Item.cmbPlanetChange.ListIndex = Item(EditorIndex).Data1
            
            LoadItemConfigs
        Else
            frmEditor_Item.frmPlanetChange.visible = False
        End If
        
        ' Basic requirements
        frmEditor_Item.scrlAccessReq.Value = .AccessReq
        frmEditor_Item.scrlLevelReq.Value = .LevelReq
        
        ' loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(i).Value = .Stat_Req(i)
        Next
        
        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(i).Name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.txtPrice.Text = Trim$(.Price)
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.Value = .Rarity
         
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With
    
    Item_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadItemConfigs()
    If Item(EditorIndex).Data1 = 2 Or Item(EditorIndex).Data1 = 3 Then
        frmEditor_Item.frmColor.visible = True
        frmEditor_Item.cmbCanal.ListIndex = Item(EditorIndex).Data2
        frmEditor_Item.scrlTom.Value = Item(EditorIndex).Data3
    Else
        frmEditor_Item.frmColor.visible = False
    End If
    
    If Item(EditorIndex).Data1 = 5 Then
        frmEditor_Item.frmAmbiente.visible = True
        frmEditor_Item.cmbAmbiente.ListIndex = Item(EditorIndex).Data2
    Else
        frmEditor_Item.frmAmbiente.visible = False
    End If
    
    If Item(EditorIndex).Data1 = 6 Then
        frmEditor_Item.frmClima.visible = True
        frmEditor_Item.cmbClima.ListIndex = Item(EditorIndex).Data2
    Else
        frmEditor_Item.frmClima.visible = False
    End If
    
    If Item(EditorIndex).Data1 = 7 Then
        frmEditor_Item.frmNPC.visible = True
        frmEditor_Item.scrlNPC.max = MAX_NPCS
        frmEditor_Item.scrlNPC.Value = Item(EditorIndex).Data2
        frmEditor_Item.txtLimit.Text = Val(Item(EditorIndex).Data3)
    Else
        frmEditor_Item.frmNPC.visible = False
    End If
    
    If Item(EditorIndex).Data1 = 8 Then
        frmEditor_Item.frmResource.visible = True
        frmEditor_Item.scrlResource.max = MAX_RESOURCES
        frmEditor_Item.scrlResource.Value = Item(EditorIndex).Data2
        frmEditor_Item.txtResourceLimit.Text = Val(Item(EditorIndex).Data3)
    Else
        frmEditor_Item.frmResource.visible = False
    End If
    
    If Item(EditorIndex).Data1 = 10 Then
        frmEditor_Item.frmSaibaman.visible = True
        frmEditor_Item.txtIndex.Text = Val(Item(EditorIndex).Data2)
    Else
        frmEditor_Item.frmSaibaman.visible = False
    End If
End Sub

Public Sub ItemEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If
    Next
    
    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Item()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Item", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Animation.visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Animation.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.Text = Trim$(.Name)
        
        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Animation.cmbSound.ListCount
                If frmEditor_Animation.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If
        
        For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).Value = .Sprite(i, 0)
            frmEditor_Animation.scrlFrameCount(i).Value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).Value = .LoopCount(i)
            
            If .looptime(i) > 0 Then
                frmEditor_Animation.scrlLoopTime(i).Value = .looptime(i)
            Else
                frmEditor_Animation.scrlLoopTime(i).Value = 45
            End If
        Next
        
        frmEditor_Animation.scrlTremor.Value = .Tremor
        frmEditor_Animation.scrlBuraco.Value = .Buraco
        
        frmEditor_Animation.optDir(0).Value = True
         
        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With
    
    Animation_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If
    Next
    
    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Animation()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Animation", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_NPC.visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_NPC
        .txtName.Text = Trim$(Npc(EditorIndex).Name)
        .txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
        If Npc(EditorIndex).Sprite < 0 Or Npc(EditorIndex).Sprite > .scrlSprite.max Then Npc(EditorIndex).Sprite = 0
        .scrlSprite.Value = Npc(EditorIndex).Sprite
        .txtSpawnSecs.Text = CStr(Npc(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = Npc(EditorIndex).Behaviour
        .scrlRange.Value = Npc(EditorIndex).Range
        .txtHP.Text = Npc(EditorIndex).HP
        .txtEXP.Text = Npc(EditorIndex).EXP
        .txtLevel.Text = Npc(EditorIndex).Level
        .txtDamage.Text = Npc(EditorIndex).Damage
        .scrlAnimation.Value = Npc(EditorIndex).Animation
        'If Npc(EditorIndex).speed = 0 Then Npc(EditorIndex).speed = 1
        .scrlMoveSpeed.Value = Npc(EditorIndex).speed
        .scrlEvent.Value = Npc(EditorIndex).Event
        .scrlEffect.Value = Npc(EditorIndex).Effect
        If Npc(EditorIndex).AttackSpeed = 0 Then Npc(EditorIndex).AttackSpeed = 1000
        .scrlAttackSpeed.Value = Npc(EditorIndex).AttackSpeed
        .scrlItemDrop.Value = 2
        .scrlItemDrop.Value = 1
        .txtLevelToPDL = Npc(EditorIndex).ND
        .chkRanged.Value = Npc(EditorIndex).Ranged
        .scrlDamage.Value = Npc(EditorIndex).ArrowDamage
        .scrlProjectile.Value = Npc(EditorIndex).ArrowAnim
        .scrlArrowAnim.Value = Npc(EditorIndex).ArrowAnimation
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Npc(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).Value = Npc(EditorIndex).Stat(i)
        Next
    End With
    
    NPC_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If
    Next
    
    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_NPC()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_NPC", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
Dim i As Long
Dim SoundSet As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Resource.visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Resource
        .scrlExhaustedPic.max = NumResources
        .scrlNormalPic.max = NumResources
        .scrlAnimation.max = MAX_ANIMATIONS
        
        .txtName.Text = Trim$(Resource(EditorIndex).Name)
        .txtMessage.Text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.Text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlTool.Value = Resource(EditorIndex).ToolRequired
        .scrlHealth.Value = Resource(EditorIndex).health
        .txtRespawnTime.Text = Resource(EditorIndex).RespawnTime
        .scrlAnimation.Value = Resource(EditorIndex).Animation
        .scrlEffect.Value = Resource(EditorIndex).Effect
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Resource(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    Resource_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If
    Next
    
    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Resource()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Resource", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Shop.visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    
    frmEditor_Shop.txtName.Text = Trim$(Shop(EditorIndex).Name)
    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.Value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.Value = 100
    End If
    
    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(Item(i).Name)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(Item(i).Name)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    
    UpdateShopTrade
    
    Shop_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
Dim i As Long, n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(i)
            ' if none, show as none
            If .Item = 0 Or .CostItem(1) = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Vazio"
            Else
                Dim Text As String
                If .CostItem(1) = 0 Then
                    Text = "[Error]"
                Else
                    Text = i & ": " & .ItemValue & "x " & Trim$(Item(.Item).Name) & " Por " & .CostValue(1) & "x " & Trim$(Item(.CostItem(1)).Name)
                End If
                For n = 2 To 5
                    If .CostItem(n) > 0 Then
                        Text = Text & " e " & .CostValue(n) & "x " & Trim$(Item(.CostItem(n)).Name)
                    End If
                Next n
                frmEditor_Shop.lstTradeItem.AddItem Text
            End If
        End With
    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateShopTrade", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If
    Next
    
    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Shop()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Shop", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Spell.visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Spell
        
        ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"
        For i = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(i).Name)
        Next
        
        If Spell(EditorIndex).ClassReq > -1 And Spell(EditorIndex).ClassReq <= Max_Classes Then
            .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        End If
        
        If Spell(EditorIndex).Type = SPELL_TYPE_TRANS Then
            .frmTrans.visible = True
        Else
            .frmTrans.visible = False
        End If
        
        ' set values
        .txtName.Text = Trim$(Spell(EditorIndex).Name)
        .txtDesc.Text = Trim$(Spell(EditorIndex).Desc)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.Value = Spell(EditorIndex).MPCost
        .scrlLevel.Value = Spell(EditorIndex).LevelReq
        .scrlAccess.Value = Spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        .scrlCast.Value = Spell(EditorIndex).CastTime
        .scrlCool.Value = Spell(EditorIndex).CDTime
        .scrlIcon.Value = Spell(EditorIndex).Icon
        .scrlMap.Value = Spell(EditorIndex).Map
        .scrlUpgrade.Value = Spell(EditorIndex).Upgrade
        .scrlRequisite.Value = Spell(EditorIndex).Requisite
        .scrlDir.Value = Spell(EditorIndex).Dir
        .scrlVital.Value = Spell(EditorIndex).Vital
        .scrlDuration.Value = Spell(EditorIndex).Item
        .txtPrice.Text = Spell(EditorIndex).Price
        .scrlRange.Value = Spell(EditorIndex).Range
        .scrlLargura.Value = Spell(EditorIndex).LinearRange
        If Spell(EditorIndex).IsAoE Then
            .chkAOE.Value = 1
            .frmAoE.visible = True
            If Spell(EditorIndex).AoEDuration > 0 Then
                .scrlAoeDuration.Value = (Spell(EditorIndex).AoEDuration / 100)
            Else
                .scrlAoeDuration.Value = 0
            End If
            If Spell(EditorIndex).AoETick > 0 Then
                .scrlEffectTick.Value = (Spell(EditorIndex).AoETick / 100)
            Else
                .scrlAoeDuration.Value = 0
            End If
        Else
            .chkAOE.Value = 0
            .frmAoE.visible = False
        End If
        .scrlAOE.Value = Spell(EditorIndex).AoE
        .scrlAnimCast.Value = Spell(EditorIndex).CastAnim
        .scrlAnim.Value = Spell(EditorIndex).SpellAnim
        .scrlStun.Value = Spell(EditorIndex).StunDuration
        .scrlEffect.Value = Spell(EditorIndex).Effect
        .scrlSprite.Value = Spell(EditorIndex).SpriteTrans
        .scrlTransAnim.Value = Spell(EditorIndex).TransAnim
        .scrlPDL.Value = Spell(EditorIndex).PDLBonus
        .scrlPlayerAnim.Value = Spell(EditorIndex).CastPlayerAnim
        .scrlChangeHair.Value = Spell(EditorIndex).HairChange
        .scrlProjectile.Value = Spell(EditorIndex).Projectile
        .scrlRotate.Value = Spell(EditorIndex).RotateSpeed
        .scrlImpact.Value = Spell(EditorIndex).Impact
        
        For i = 1 To 3
            .scrlSpellLinearAnim(i - 1).Value = Spell(EditorIndex).SpellLinearAnim(i)
        Next i
        
        For i = 0 To 4
            If Spell(EditorIndex).Add_Stat(i + 1) > .scrlStat(i).max Then Spell(EditorIndex).Add_Stat(i + 1) = 0
            .scrlStat(i).Value = Spell(EditorIndex).Add_Stat(i + 1)
        Next i
        
        For i = 1 To Vitals.Vital_Count - 1
            .scrlTransVital(i - 1).Value = Spell(EditorIndex).TransVital(i)
        Next i
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Spell(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    Spell_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If
    Next
    
    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Spell()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Spell", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAttributeDialogue()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Map.fraNpcSpawn.visible = False
    frmEditor_Map.fraResource.visible = False
    frmEditor_Map.fraMapItem.visible = False
    frmEditor_Map.fraMapWarp.visible = False
    frmEditor_Map.fraShop.visible = False
    frmEditor_Map.fraSoundEffect.visible = False
    frmEditor_Map.fraEvent.visible = False
    frmEditor_Map.fraArena.visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAttributeDialogue", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub Events_ClearChanged()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Event_Changed(i) = False
    Next i
End Sub

Public Sub EventEditorInit()
    Dim i As Long
    With frmEditor_Events
        If .visible = False Then Exit Sub
        EditorIndex = .lstIndex.ListIndex + 1
        .txtName = Trim$(Events(EditorIndex).Name)
        .chkPlayerSwitch.Value = Events(EditorIndex).chkSwitch
        .chkPlayerVar.Value = Events(EditorIndex).chkVariable
        .chkHasItem.Value = Events(EditorIndex).chkHasItem
        .cmbPlayerSwitch.ListIndex = Events(EditorIndex).SwitchIndex
        .cmbPlayerSwitchCompare.ListIndex = Events(EditorIndex).SwitchCompare
        .cmbPlayerVar.ListIndex = Events(EditorIndex).VariableIndex
        .cmbPlayerVarCompare.ListIndex = Events(EditorIndex).VariableCompare
        .txtPlayerVariable.Text = Events(EditorIndex).VariableCondition
        .cmbHasItem.ListIndex = Events(EditorIndex).HasItemIndex - 1
        .cmbTrigger.ListIndex = Events(EditorIndex).Trigger
        .chkWalkthrought.Value = Events(EditorIndex).WalkThrought
        Call .PopulateSubEventList
    End With
    Event_Changed(EditorIndex) = True
End Sub

Public Sub EventEditorOk()
Dim i As Long
    For i = 1 To MAX_EVENTS
        If Event_Changed(i) Then
            Call Events_SendSaveEvent(i)
        End If
    Next i
    
    Unload frmEditor_Events
    Events_ClearChanged
    Editor = 0
End Sub

Public Sub EventEditorCancel()
    Editor = 0
    Unload frmEditor_Events
    Events_ClearChanged
    ClearEvents
    Events_SendRequestEventsData
End Sub

' *********************
' ** Event Utilities **
' *********************
Public Function GetSubEventCount(ByVal Index As Long)
    GetSubEventCount = 0
    If Index <= 0 Or Index > MAX_EVENTS Then Exit Function
    If Events(Index).HasSubEvents Then
        GetSubEventCount = UBound(Events(Index).SubEvents)
    End If
End Function

' /////////////////
' // Effect Editor //
' /////////////////
Public Sub EffectEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Effect.visible = False Then Exit Sub
    EditorIndex = frmEditor_Effect.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Effect.cmbSound.Clear
    frmEditor_Effect.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Effect.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Effect(EditorIndex)
        frmEditor_Effect.txtName.Text = Trim$(.Name)
        
        ' find the sound we have set
        If frmEditor_Effect.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Effect.cmbSound.ListCount
                If frmEditor_Effect.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Effect.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Effect.cmbSound.ListIndex = -1 Then frmEditor_Effect.cmbSound.ListIndex = 0
        End If
        frmEditor_Effect.scrlSprite.Value = .Sprite
        frmEditor_Effect.cmbType.ListIndex = .Type - 1
        frmEditor_Effect.scrlParticles.Value = .Particles
        frmEditor_Effect.scrlSize.Value = .Size
        frmEditor_Effect.scrlAlpha.Value = .Alpha
        frmEditor_Effect.scrlDecay.Value = .Decay
        frmEditor_Effect.scrlRed.Value = .Red
        frmEditor_Effect.scrlGreen.Value = .Green
        frmEditor_Effect.scrlBlue.Value = .Blue
        frmEditor_Effect.scrlXSpeed.Value = .XSpeed
        frmEditor_Effect.scrlYSpeed.Value = .YSpeed
        frmEditor_Effect.scrlXAcc.Value = .XAcc
        frmEditor_Effect.scrlYAcc.Value = .YAcc
        frmEditor_Effect.optEffectType(.isMulti) = True
        If .isMulti = 1 Then
            frmEditor_Effect.fraMultiParticle.visible = True
            frmEditor_Effect.fraEffect.visible = False
            frmEditor_Effect.scrlEffect.Value = .MultiParticle(1)
        Else
            frmEditor_Effect.fraEffect.visible = True
            frmEditor_Effect.fraMultiParticle.visible = False
        End If
        frmEditor_Effect.scrlDuration = .Duration
        
        EditorIndex = frmEditor_Effect.lstIndex.ListIndex + 1
    End With
    
    Effect_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EffectEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EffectEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_EFFECTS
        If Effect_Changed(i) Then
            Call SendSaveEffect(i)
        End If
    Next
    
    Unload frmEditor_Effect
    Editor = 0
    ClearChanged_Effect
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EffectEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EffectEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Effect
    ClearChanged_Effect
    ClearEffects
    SendRequestEffects
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EffectEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Effect()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Effect_Changed(1), MAX_EFFECTS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Effect", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub QuestEditorInit()
    Dim i As Long

    With frmEditor_Quest
        i = .lstIndex.ListIndex + 1
        .txtName = Trim$(Quest(i).Name)
        .txtDesc = Trim$(Quest(i).Desc)
        .scrlIcon.Value = Quest(i).Icon
        .txtEvent = Val(Quest(i).EventNum)
        .chkDay.Value = Val(Quest(i).NotDay)
        .chkNight = Val(Quest(i).NotNight)
        .chkRepeat = Val(Quest(i).Repeat)
        .txtHours = Val(Quest(i).Cooldown)
        .cmbType.ListIndex = Val(Quest(i).Type)
        
        .scrlIcon.max = numitems
    End With
End Sub

Public Sub QuestEditorOk(QuestNum As Long)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SendSavequest(QuestNum)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
