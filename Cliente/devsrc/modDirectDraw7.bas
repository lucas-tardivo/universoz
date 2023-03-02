Attribute VB_Name = "modGraphics"
Option Explicit
' **********************
' ** Renders graphics **
' **********************
' DirectX8 Object
Private DirectX8 As DirectX8 'The master DirectX object.
Private Direct3D As Direct3D8 'Controls all things 3D.
Public Direct3D_Device As Direct3DDevice8 'Represents the hardware rendering.
Public Direct3DX As D3DX8

'The 2D (Transformed and Lit) vertex format.
Public Const FVF_TLVERTEX As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

'The 2D (Transformed and Lit) vertex format type.
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    RHW As Single
    color As Long
    TU As Single
    TV As Single
End Type

Private Vertex_List(3) As TLVERTEX '4 vertices will make a square.

'Some color depth constants to help make the DX constants more readable.
Private Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Private Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Private Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8

Public RenderingMode As Long

Private Direct3D_Window As D3DPRESENT_PARAMETERS 'Backbuffer and viewport description.
Private Display_Mode As D3DDISPLAYMODE

Public ScreenWidth As Long
Public ScreenHeight As Long

'Graphic Textures
Public Tex_GUI() As DX8TextureRec
Public Tex_Buttons() As DX8TextureRec
Public Tex_Buttons_h() As DX8TextureRec
Public Tex_Buttons_c() As DX8TextureRec
Public Tex_Item() As DX8TextureRec ' arrays
Public Tex_Character() As DX8TextureRec
Public Tex_Paperdoll() As DX8TextureRec
Public Tex_Tileset() As DX8TextureRec
Public Tex_Resource() As DX8TextureRec
Public Tex_Animation() As DX8TextureRec
Public Tex_SpellIcon() As DX8TextureRec
Public Tex_Face() As DX8TextureRec
Public Tex_Fog() As DX8TextureRec
Public Tex_Panorama() As DX8TextureRec
Public Tex_Particle() As DX8TextureRec
Public Tex_Projectile() As DX8TextureRec
Public Tex_Hair() As hairrec
Public Tex_Transportes() As DX8TextureRec
Public Tex_Blood As DX8TextureRec ' singes
Public Tex_Misc As DX8TextureRec
Public Tex_Direction As DX8TextureRec
Public Tex_Target As DX8TextureRec
Public Tex_Bars As DX8TextureRec
Public Tex_Selection As DX8TextureRec
Public Tex_White As DX8TextureRec
Public Tex_Weather As DX8TextureRec
Public Tex_Fade As DX8TextureRec
Public Tex_Shadow As DX8TextureRec
Public Tex_Ambiente As DX8TextureRec
Public Tex_Clouds As DX8TextureRec
Public Tex_Alerta As DX8TextureRec
Public Tex_Scouter As DX8TextureRec
Public Tex_Buraco As DX8TextureRec
Public Tex_HairBase As DX8TextureRec

' Number of graphic files
Public NumGUIs As Long
Public NumButtons As Long
Public NumButtons_c As Long
Public NumButtons_h As Long
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public numitems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Public NumFogs As Long
Public NumPanoramas As Long
Public NumParticles As Long
Public NumProjectiles As Long
Public NumTransportes As Long

Public Type DX8TextureRec
    Texture As Long
    Width As Long
    Height As Long
    filepath As String
    TexWidth As Long
    TexHeight As Long
    ImageData() As Byte
    HasData As Boolean
End Type

Public Type GlobalTextureRec
    Texture As Direct3DTexture8
    TexWidth As Long
    TexHeight As Long
    Timer As Long
End Type

Public Type RECT
    Top As Long
    Left As Long
    Bottom As Long
    Right As Long
End Type

Type hairrec
    TexHair() As DX8TextureRec
End Type

Public gTexture() As GlobalTextureRec
Public NumTextures As Long
Public CurrentTexture As Long
Public BubbleOpaque As Byte
Public AlertX As Long
Public ScouterOn As Boolean
Public Tremor As Long
Public TremorX As Long
Public NumHair(0 To TotalHairTypes) As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDX8() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set DirectX8 = New DirectX8 'Creates the DirectX object.
    Set Direct3D = DirectX8.Direct3DCreate() 'Creates the Direct3D object using the DirectX object.
    Set Direct3DX = New D3DX8
    
    ScreenWidth = 800
    ScreenHeight = 600
    
    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.
    
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_DISCARD 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'DispMode.Format = D3DFMT_X8R8G8B8
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = ScreenWidth ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = ScreenHeight 'frmMain.Scaleheight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hWnd 'Use frmMain as the device window.
    
    'we've already setup for Direct3D_Window.
    If TryCreateDirectX8Device = False Then
        MsgBox "Unable to initialize DirectX8. You may be missing dx8vb.dll or have incompatible hardware to use DirectX8."
        DestroyGame
    End If

    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    End With
    
    ' Initialise the surfaces
    LoadTextures
    Engine_Init_Particles
    ' We're done
    InitDX8 = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InitDX8", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function TryCreateDirectX8Device() As Boolean
Dim i As Long

On Error GoTo nexti

    For i = 1 To 4
        Select Case i
            Case 1
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 2
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 3
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_MIXED_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 4
                TryCreateDirectX8Device = False
                Exit Function
        End Select
nexti:
    Next

End Function

Function GetNearestPOT(Value As Long) As Long
Dim i As Long
    Do While 2 ^ i < Value
        i = i + 1
    Loop
    GetNearestPOT = 2 ^ i
End Function
Public Sub SetTexture(ByRef TextureRec As DX8TextureRec)
    If TextureRec.Texture <> CurrentTexture Then
        If TextureRec.Texture > NumTextures Then TextureRec.Texture = NumTextures
        If TextureRec.Texture < 0 Then TextureRec.Texture = 0
        
        If Not TextureRec.Texture = 0 Then
            If gTexture(TextureRec.Texture).Timer = 0 Then
                Call LoadTexture(TextureRec)
                gTexture(TextureRec.Texture).Timer = GetTickCount + 150000
                If DEBUG_MODE = True Then AddText "Loaded texture: " & TextureRec.Texture, White
            End If
        End If
        CurrentTexture = TextureRec.Texture
    End If
End Sub
Public Sub UnsetTexture(ByRef TextureNum As Long)
    If gTexture(TextureNum).Timer < GetTickCount Then
        Set gTexture(TextureNum).Texture = Nothing
        gTexture(TextureNum).Timer = 0
        If DEBUG_MODE = True Then AddText "Unloaded texture: " & TextureNum, White
    End If
End Sub
Public Sub LoadTexture(ByRef TextureRec As DX8TextureRec)
Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, i As Long
Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long
Dim BMU As clsBitmapUtils
Dim PathName As String
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If TextureRec.HasData = False Then
        Set GDIToken = New cGDIpToken
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        Set SourceBitmap = New cGDIpImage
        
        If Mid(TextureRec.filepath, Len(TextureRec.filepath) - Len(GFX_EXT) + 1, Len(GFX_EXT)) = GFX_EXT Then
            Set BMU = New clsBitmapUtils
            With BMU
                Call .LoadByteData(TextureRec.filepath)
                Call .DecryptByteData(GFX_PASSWORD)
                Call .DecompressByteData      'If you want to use zlib, you can change this to .DecompressByteData_ZLib
            End With
            
            TextureRec.Width = BMU.ImageWidth
            TextureRec.Height = BMU.ImageHeight
            
            PathName = WinDir & "temp.bmp"
            
LoadPathName:
            BMU.SaveBitmap PathName
            SourceBitmap.LoadPicture_FileName PathName, GDIToken
            
            On Error GoTo ReloadPathName
            
            Kill PathName
            
            If 0 > 1 Then
ReloadPathName:
                PathName = App.Path & "\temp.bmp"
                GoTo LoadPathName
            End If
        
        Else
            
            Call SourceBitmap.LoadPicture_FileName(TextureRec.filepath, GDIToken)
            
            TextureRec.Width = SourceBitmap.Width
            TextureRec.Height = SourceBitmap.Height
        
        End If
        
        SourceBitmap.ExtraTransparentColor = SourceBitmap.GetPixel(0, 0)
        
        newWidth = GetNearestPOT(TextureRec.Width)
        newHeight = GetNearestPOT(TextureRec.Height)
        'If newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height Then
        If Mid(TextureRec.filepath, Len(TextureRec.filepath) - 8) <> "white" & GFX_EXT Then
            Set ConvertedBitmap = New cGDIpImage
            Set GDIGraphics = New cGDIpRenderer
            i = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
            Call ConvertedBitmap.LoadPicture_FromNothing(newHeight, newWidth, i, GDIToken) 'I HAVE NO IDEA why this is backwards but it works.
            Call GDIGraphics.DestroyHGraphics(i)
            i = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
            Call GDIGraphics.AttachTokenClass(GDIToken)
            Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, i)
            Call ConvertedBitmap.SaveAsPNG(ImageData)
            GDIGraphics.DestroyHGraphics (i)
            TextureRec.ImageData = ImageData
            Set ConvertedBitmap = Nothing
            Set GDIGraphics = Nothing
            Set SourceBitmap = Nothing
        Else
            Call SourceBitmap.SaveAsPNG(ImageData)
            TextureRec.ImageData = ImageData
            Set SourceBitmap = Nothing
        End If
    Else
        ImageData = TextureRec.ImageData
    End If
    
    
    Set gTexture(TextureRec.Texture).Texture = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, _
                                                    ImageData(0), _
                                                    UBound(ImageData) + 1, _
                                                    newWidth, _
                                                    newHeight, _
                                                    D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
    
    gTexture(TextureRec.Texture).TexWidth = newWidth
    gTexture(TextureRec.Texture).TexHeight = newHeight
    Exit Sub
errorhandler:
    HandleError "LoadTexture", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub LoadTextures()
Dim i As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckGUIs
    Call CheckButtons
    Call CheckButtons_c
    Call CheckButtons_h
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    Call CheckFogs
    Call CheckPanoramas
    Call CheckParticles
    Call CheckProjectiles
    Call CheckHair
    Call CheckTransportes
    
    NumTextures = NumTextures + 16
    
    ReDim Preserve gTexture(NumTextures)
    Tex_HairBase.filepath = App.Path & "\data files\graphics\adm\Hair" & GFX_EXT
    Tex_HairBase.Texture = NumTextures - 15
    Tex_Buraco.filepath = App.Path & "\data files\graphics\misc\buraco" & GFX_EXT
    Tex_Buraco.Texture = NumTextures - 14
    Tex_Scouter.filepath = App.Path & "\data files\graphics\misc\scoutertarget" & GFX_EXT
    Tex_Scouter.Texture = NumTextures - 13
    Tex_Alerta.filepath = App.Path & "\data files\graphics\misc\alerta" & GFX_EXT
    Tex_Alerta.Texture = NumTextures - 12
    Tex_Clouds.filepath = App.Path & "\data files\graphics\misc\nuvens.png"
    Tex_Clouds.Texture = NumTextures - 11
    Tex_Ambiente.filepath = App.Path & "\data files\graphics\misc\ambiente" & GFX_EXT
    Tex_Ambiente.Texture = NumTextures - 10
    Tex_Shadow.filepath = App.Path & "\data files\graphics\misc\shadow.png"
    Tex_Shadow.Texture = NumTextures - 9
    Tex_Fade.filepath = App.Path & "\data files\graphics\misc\fader" & GFX_EXT
    Tex_Fade.Texture = NumTextures - 8
    Tex_Weather.filepath = App.Path & "\data files\graphics\misc\weather" & GFX_EXT
    Tex_Weather.Texture = NumTextures - 7
    Tex_White.filepath = App.Path & "\data files\graphics\misc\white" & GFX_EXT
    Tex_White.Texture = NumTextures - 6
    Tex_Direction.filepath = App.Path & "\data files\graphics\misc\direction" & GFX_EXT
    Tex_Direction.Texture = NumTextures - 5
    Tex_Target.filepath = App.Path & "\data files\graphics\misc\target" & GFX_EXT
    Tex_Target.Texture = NumTextures - 4
    Tex_Misc.filepath = App.Path & "\data files\graphics\misc\misc" & GFX_EXT
    Tex_Misc.Texture = NumTextures - 3
    Tex_Blood.filepath = App.Path & "\data files\graphics\misc\blood" & GFX_EXT
    Tex_Blood.Texture = NumTextures - 2
    Tex_Bars.filepath = App.Path & "\data files\graphics\misc\bars.png"
    Tex_Bars.Texture = NumTextures - 1
    Tex_Selection.filepath = App.Path & "\data files\graphics\misc\select" & GFX_EXT
    Tex_Selection.Texture = NumTextures
    
    EngineInitFontTextures
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTextures", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadTextures()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    On Error Resume Next
    
    For i = 1 To NumTextures
        Set gTexture(i).Texture = Nothing
        ZeroMemory ByVal VarPtr(gTexture(i)), LenB(gTexture(i))
    Next
    
    ReDim gTexture(1)

    
    For i = 1 To NumTileSets
        Tex_Tileset(i).Texture = 0
    Next

    For i = 1 To numitems
        Tex_Item(i).Texture = 0
    Next

    For i = 1 To NumCharacters
        Tex_Character(i).Texture = 0
    Next
    
    For i = 1 To NumPaperdolls
        Tex_Paperdoll(i).Texture = 0
    Next
    
    For i = 1 To NumResources
        Tex_Resource(i).Texture = 0
    Next
    
    For i = 1 To NumAnimations
        Tex_Animation(i).Texture = 0
    Next
    
    For i = 1 To NumSpellIcons
        Tex_SpellIcon(i).Texture = 0
    Next
    
    For i = 1 To NumFaces
        Tex_Face(i).Texture = 0
    Next
    
    For i = 1 To NumGUIs
        Tex_GUI(i).Texture = 0
    Next
    
    For i = 1 To NumButtons
        Tex_Buttons(i).Texture = 0
    Next
    
    For i = 1 To NumButtons_c
        Tex_Buttons_c(i).Texture = 0
    Next
    
    For i = 1 To NumButtons_h
        Tex_Buttons_h(i).Texture = 0
    Next
    
    Tex_Misc.Texture = 0
    Tex_Blood.Texture = 0
    Tex_Direction.Texture = 0
    Tex_Target.Texture = 0
    Tex_Selection.Texture = 0
    Tex_Bars.Texture = 0
    Tex_White.Texture = 0
    Tex_Weather.Texture = 0
    Tex_Fade.Texture = 0
    Tex_Shadow.Texture = 0
    
    UnloadFontTextures
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadTextures", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Drawing **
' **************
Public Sub RenderTexture(ByRef TextureRec As DX8TextureRec, ByVal dX As Single, ByVal dY As Single, ByVal sx As Single, ByVal sY As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional color As Long = -1, Optional ByVal Degrees As Single = 0)
    Dim TextureNum As Long
    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single
    Dim RadAngle As Single 'The angle in Radians
    Dim CenterX As Single
    Dim CenterY As Single
    Dim NewX As Single
    Dim NewY As Single
    Dim SinRad As Single
    Dim CosRad As Single
    Dim i As Long
    
    SetTexture TextureRec
    
    TextureNum = TextureRec.Texture
    
    textureWidth = gTexture(TextureNum).TexWidth
    textureHeight = gTexture(TextureNum).TexHeight
    
    If sY + sHeight > textureHeight Then Exit Sub
    If sx + sWidth > textureWidth Then Exit Sub
    If sx < 0 Then Exit Sub
    If sY < 0 Then Exit Sub

    sx = sx - 0.5
    sY = sY - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sx / textureWidth)
    sourceY = (sY / textureHeight)
    sourceWidth = ((sx + sWidth) / textureWidth)
    sourceHeight = ((sY + sHeight) / textureHeight)
    
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    'Check if a rotation is required
    If Degrees <> 0 And Degrees <> 360 Then

        'Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        'Set the CenterX and CenterY values
        CenterX = dX + (dWidth * 0.5)
        CenterY = dY + (dHeight * 0.5)

        'Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        'Loops through the passed vertex buffer
        For i = 0 To 3

            'Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (Vertex_List(i).X - CenterX) * CosRad - (Vertex_List(i).Y - CenterY) * SinRad
            NewY = CenterY + (Vertex_List(i).Y - CenterY) * CosRad + (Vertex_List(i).X - CenterX) * SinRad

            'Applies the new co-ordinates to the buffer
            Vertex_List(i).X = NewX
            Vertex_List(i).Y = NewY
        Next
    End If
    
    Call Direct3D_Device.SetTexture(0, gTexture(TextureNum).Texture)
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))
End Sub

Public Sub RenderTextureByRects(TextureRec As DX8TextureRec, sRECT As RECT, dRect As RECT)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    RenderTexture TextureRec, dRect.Left, dRect.Top, sRECT.Left, sRECT.Top, dRect.Right - dRect.Left, dRect.Bottom - dRect.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTextureByRects", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDirection(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' render dir blobs
    For i = 1 To 4
        rec.Left = (i - 1) * 8
        rec.Right = rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(X, Y).DirBlock, CByte(i)) Then
            rec.Top = 8
        Else
            rec.Top = 16
        End If
        rec.Bottom = rec.Top + 8
        'render!
        RenderTexture Tex_Direction, ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDirection", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTarget(ByVal X As Long, ByVal Y As Long)
Dim sRECT As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If
    
    If Not ScouterOn Then
        RenderTexture Tex_Target, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
    Else
        If GetTickCount Mod 1000 < 500 Then
        RenderTexture Tex_Scouter, X - 14, Y + 8, 0, 0, Tex_Scouter.Width, Tex_Scouter.Height, Tex_Scouter.Width, Tex_Scouter.Height
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTarget", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawHover(ByVal tType As Long, ByVal Target As Long, ByVal X As Long, ByVal Y As Long)
Dim sRECT As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    If ScouterOn Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = Width
        .Right = .Left + Width
    End With
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If
    
    RenderTexture Tex_Target, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawHover", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapTile(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT
Dim i As Long
Dim n As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Ground To MapLayer.Mask2
            If GetTickCount Mod 800 < 400 And i = MapLayer.Mask Then
                If (.Layer(MapLayer.MaskAnim).Tileset > 0 And .Layer(MapLayer.MaskAnim).Tileset <= NumTileSets) And (.Layer(MapLayer.MaskAnim).X > 0 Or .Layer(MapLayer.MaskAnim).Y > 0) Then
                    RenderTexture Tex_Tileset(.Layer(MapLayer.MaskAnim).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(MapLayer.MaskAnim).X * 32, .Layer(MapLayer.MaskAnim).Y * 32, 32, 32, 32, 32, -1
                    GoTo NextLayer
                End If
            End If
            If GetTickCount Mod 800 < 400 And i = MapLayer.Mask2 Then
                If (.Layer(MapLayer.Mask2Anim).Tileset > 0 And .Layer(MapLayer.Mask2Anim).Tileset <= NumTileSets) And (.Layer(MapLayer.Mask2Anim).X > 0 Or .Layer(MapLayer.Mask2Anim).Y > 0) Then
                    RenderTexture Tex_Tileset(.Layer(MapLayer.Mask2Anim).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(MapLayer.Mask2Anim).X * 32, .Layer(MapLayer.Mask2Anim).Y * 32, 32, 32, 32, 32, -1
                    GoTo NextLayer
                End If
            End If
            If Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                If i = MapLayer.Mask Then
                    If .Type = TILE_TYPE_EVENT Then
                        If .Data1 > 0 Then
                            If Events(.Data1).WalkThrought = NO Then
                                If Player(MyIndex).EventOpen(.Data1) = YES Then Exit Sub
                            End If
                        End If
                    End If
                End If
                RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32, -1
            ElseIf Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            End If
NextLayer:

        Next
    End With
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "DrawMapTile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapFringeTile(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            If GetTickCount Mod 800 < 400 And i = MapLayer.Fringe Then
                If (.Layer(MapLayer.FringeAnim).Tileset > 0 And .Layer(MapLayer.FringeAnim).Tileset <= NumTileSets) And (.Layer(MapLayer.FringeAnim).X > 0 Or .Layer(MapLayer.FringeAnim).Y > 0) Then
                    RenderTexture Tex_Tileset(.Layer(MapLayer.FringeAnim).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(MapLayer.FringeAnim).X * 32, .Layer(MapLayer.FringeAnim).Y * 32, 32, 32, 32, 32, -1
                    GoTo NextLayer
                End If
            End If
            If Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_NORMAL Then
                RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32, -1
            ElseIf Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            End If
NextLayer:
        Next
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapFringeTile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawBlood(ByVal Index As Long)
Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    'load blood then
    BloodCount = Tex_Blood.Width / 32
    
    With Blood(Index)
        If .Alpha <= 0 Then Exit Sub
        ' check if we should be seeing it
        If .Timer + 20000 < GetTickCount Then
            .Alpha = .Alpha - 1
        End If
        
        rec.Top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        RenderTexture Tex_Blood, ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorARGB(.Alpha, 255, 255, 255)
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBlood", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
Dim Sprite As Integer, sRECT As RECT, i As Long, Width As Long, Height As Long, looptime As Long, FrameCount As Long
Dim X As Long, Y As Long, lockindex As Long, Rotation As Integer
    
    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    On Error Resume Next
    'Animation Dir
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer, AnimInstance(Index).Dir)
    
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    ' pre-load texture for calculations
    'SetTexture Tex_Anim(Sprite)
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    ' total width divided by frame count
    Width = 192
    Height = 192
    
    With sRECT
        .Top = (Height * ((AnimInstance(Index).frameIndex(Layer) - 1) \ AnimColumns))
        .Bottom = .Top + Height
        .Left = (Width * (((AnimInstance(Index).frameIndex(Layer) - 1) Mod AnimColumns)))
        .Right = .Left + Width
    End With
    
    If AnimInstance(Index).IsLinear = 1 Then
        Select Case AnimInstance(Index).Dir
            Case DIR_UP: Rotation = 270
            Case DIR_DOWN: Rotation = 90
            Case DIR_LEFT: Rotation = 180
            Case DIR_RIGHT: Rotation = 0
        End Select
    Else
        Rotation = 0
    End If
    
    ' change x or y if locked
    If AnimInstance(Index).locktype > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).locktype = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + TempPlayer(lockindex).xOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + TempPlayer(lockindex).YOffset
                End If
            End If
        ElseIf AnimInstance(Index).locktype = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).Num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + TempMapNpc(lockindex).xOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + TempMapNpc(lockindex).YOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If
    
    X = X + Animation(AnimInstance(Index).Animation).XAxis(AnimInstance(Index).Dir)
    Y = Y + Animation(AnimInstance(Index).Animation).YAxis(AnimInstance(Index).Dir)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    'EngineRenderRectangle Tex_Anim(sprite), x, y, sRECT.left, sRECT.top, sRECT.width, sRECT.height, sRECT.width, sRECT.height, sRECT.width, sRECT.height
    RenderTexture Tex_Animation(Sprite), X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, , Rotation
End Sub

Public Sub DrawMapResource(ByVal Resource_num As Long)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As RECT
Dim X As Long, Y As Long
Dim i As Long, Alpha As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).X > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).Data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' cut down everything if we're editing
    If InMapEditor And frmEditor_Map.chkResources.Value = 0 Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' src rect
    With rec
        .Top = 0
        .Bottom = Tex_Resource(Resource_sprite).Height
        .Left = 0
        .Right = Tex_Resource(Resource_sprite).Width
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (Tex_Resource(Resource_sprite).Width / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - Tex_Resource(Resource_sprite).Height + 32
    

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If ConvertMapY(GetPlayerY(i)) < ConvertMapY(MapResource(Resource_num).Y) And ConvertMapY(GetPlayerY(i)) > ConvertMapY(MapResource(Resource_num).Y) - (Tex_Resource(Resource_sprite).Height) / 32 Then
                If ConvertMapX(GetPlayerX(i)) >= ConvertMapX(MapResource(Resource_num).X) - ((Tex_Resource(Resource_sprite).Width / 2) / 32) And ConvertMapX(GetPlayerX(i)) <= ConvertMapX(MapResource(Resource_num).X) + ((Tex_Resource(Resource_sprite).Width / 2) / 32) Then
                    Alpha = 150
                Else
                    Alpha = 255
                End If
            Else
                Alpha = 255
            End If
        End If
    Next

    
    ' render it
    Call DrawResource(Resource_sprite, Alpha, X, Y, rec)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawResource(ByVal Resource As Long, ByVal Alpha As Long, ByVal dX As Long, dY As Long, rec As RECT)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long
Dim destRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub

    X = ConvertMapX(dX)
    Y = ConvertMapY(dY)
    
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    
    RenderTexture Tex_Resource(Resource), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, Alpha)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRECT As RECT
Dim barWidth As Long
Dim i As Long, npcNum As Long, partyIndex As Long, HPMod As Byte
Dim Totalbarwidth As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SetTexture Tex_Bars
    ' dynamic bar calculations
    
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(i).Num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) <= MapNpc(i).MaxHP And i = myTarget And myTargetType = TARGET_TYPE_NPC Then
                If Npc(MapNpc(i).Num).ND = 0 Then
                    sWidth = Tex_Bars.Width
                    sHeight = 19
                    
                    ' lock to npc
                    tmpX = MapNpc(i).X * PIC_X + TempMapNpc(i).xOffset + 16 - (sWidth / 2)
                    tmpY = MapNpc(i).Y * PIC_Y + TempMapNpc(i).YOffset + 35
                    
                    sWidth = 96
                    ' calculate the width to fill
                    If sWidth > 0 Then BarWidth_NpcHP_Max(i) = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (MapNpc(i).MaxHP / sWidth)) * sWidth
                    sWidth = Tex_Bars.Width
                    ' draw bar background
                    With sRECT
                        .Top = 19 ' HP bar background
                        .Left = 0
                        .Right = .Left + 125
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 200)
                    
                    ' draw the bar proper
                    With sRECT
                        .Top = 10 ' HP bar
                        .Left = 0
                        .Right = .Left + BarWidth_NpcHP(i)
                        .Bottom = .Top + 9
                    End With
                    
                    tmpX = MapNpc(i).X * PIC_X + TempMapNpc(i).xOffset + 29 - (sWidth / 2)
                    tmpY = MapNpc(i).Y * PIC_Y + TempMapNpc(i).YOffset + 42
                    
                    HPMod = 255 * ((MapNpc(i).Vital(Vitals.HP) / MapNpc(i).MaxHP))
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 0, 0, 200)
                Else
                    sWidth = 48
                    sHeight = 10
                    
                    ' lock to npc
                    tmpX = MapNpc(i).X * PIC_X + TempMapNpc(i).xOffset + 16 - (sWidth / 2)
                    tmpY = MapNpc(i).Y * PIC_Y + TempMapNpc(i).YOffset + 35
                    
                    ' calculate the width to fill
                    If sWidth > 0 Then BarWidth_NpcHP_Max(i) = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (MapNpc(i).MaxHP / sWidth)) * sWidth
                    ' draw bar background
                    With sRECT
                        .Top = 0 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 200)
                    
                    ' draw the bar proper
                    With sRECT
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + BarWidth_NpcHP(i)
                        .Bottom = .Top + sHeight
                    End With
                    
                    HPMod = 255 * ((MapNpc(i).Vital(Vitals.HP) / MapNpc(i).MaxHP))
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255 - HPMod, 255, 0, 200)
                End If
            End If
        End If
    Next
    
    sWidth = 48
    sHeight = 10

    ' check for casting time bar
    If TempPlayer(MyIndex).SpellBuffer > 0 Then
        If Spell(PlayerSpells(TempPlayer(MyIndex).SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + TempPlayer(MyIndex).xOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + TempPlayer(MyIndex).YOffset + 24 + sHeight + 1
            
            ' calculate the width to fill
            barWidth = (GetTickCount - TempPlayer(MyIndex).SpellBufferTimer) / ((Spell(PlayerSpells(TempPlayer(MyIndex).SpellBuffer)).CastTime * 1000)) * sWidth
            
            ' draw bar background
            With sRECT
                .Top = 0 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 200)
            
            ' draw the bar proper
            With sRECT
                .Top = 0 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .Top + sHeight
            End With
            If barWidth < sWidth Then
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(15, 180, 255, 200)
            Else
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 0, 0, 200)
            End If
        End If
    End If
    
    ' draw own health bar
    'If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
    '    tmpX = GetPlayerX(MyIndex) * PIC_X + TempPlayer(MyIndex).XOffSet + 16 - (sWidth / 2)
    '    tmpY = GetPlayerY(MyIndex) * PIC_X + TempPlayer(MyIndex).YOffSet + 35
       
        ' calculate the width to fill
    '    If sWidth > 0 Then BarWidth_PlayerHP_Max(MyIndex) = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
        ' draw bar background
    '    With sRECT
    '        .Top = sHeight * 1 ' HP bar background
    '        .Left = 0
    '        .Right = .Left + sWidth
    '        .Bottom = .Top + sHeight
    '    End With
    '    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
       
        ' draw the bar proper
    '    With sRECT
    '        .Top = 0 ' HP bar
    '        .Left = 0
    '        .Right = .Left + BarWidth_PlayerHP(MyIndex)
    '        .Bottom = .Top + sHeight
    '    End With
    '    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
    'End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + TempPlayer(partyIndex).xOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + TempPlayer(partyIndex).YOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' draw bar background
                    With sRECT
                        .Top = 0 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 200)
                    
                    ' draw the bar proper
                    With sRECT
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(0, 255, 0, 150)
                End If
            End If
        Next
    End If
    
    'Targeto
    If myTargetType = TARGET_TYPE_PLAYER And myTarget > 0 Then
                ' player exists
                If GetPlayerVital(myTarget, Vitals.HP) > 0 And GetPlayerVital(myTarget, Vitals.HP) <= GetPlayerMaxVital(myTarget, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(myTarget) * PIC_X + TempPlayer(myTarget).xOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(myTarget) * PIC_X + TempPlayer(myTarget).YOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(myTarget, Vitals.HP) / sWidth) / (GetPlayerMaxVital(myTarget, Vitals.HP) / sWidth)) * sWidth
                    
                    ' draw bar background
                    With sRECT
                        .Top = 0 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 200)
                    
                    ' draw the bar proper
                    With sRECT
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(0, 255, 0, 150)
                End If
    End If
                    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBars", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub DrawPlayer(ByVal Index As Long)
Dim Anim As Byte, i As Long, X As Long, Y As Long
Dim Sprite As Long, spritetop As Long, Hair As Byte
Dim rec As RECT
Dim AttackSpeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = GetPlayerSprite(Index)
    Hair = Player(Index).Hair
    Hair = 1

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    If frmPaperdoll.visible = True Then
        Call DrawPaperdollTest
        Exit Sub
    End If

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        AttackSpeed = Item(GetPlayerEquipment(Index, Weapon)).speed - (GetPlayerStat(Index, Agility) * 10)
    Else
        AttackSpeed = 1000 - (GetPlayerStat(Index, Agility) * 10)
    End If
    
    If AttackSpeed < 200 Then AttackSpeed = 200

    If VXFRAME = False Then
        ' Reset frame
        If TempPlayer(Index).Step = 3 Then
            Anim = 0
        ElseIf TempPlayer(Index).Step = 1 Then
            Anim = 2
        End If
    Else
        Anim = 1
    End If
    
    ' Check for attacking animation
    If TempPlayer(Index).AttackTimer + (AttackSpeed / 2) > GetTickCount Then
        If TempPlayer(Index).Attacking = 1 Then
            Dim Porc As Long
            Porc = (TempPlayer(Index).AttackTimer + (AttackSpeed / 2)) - GetTickCount
            Porc = (Porc / (AttackSpeed / 2)) * 100
            
            If TempPlayer(Index).AttackAnim = 0 Then
                If Porc < 50 Then Anim = 7
                If Porc >= 50 Then Anim = 6
            Else
                If Porc < 50 Then Anim = 9
                If Porc >= 50 Then Anim = 8
            End If
        End If
    Else
        ' If not attacking, walk normally
        If TempPlayer(Index).Fly = 0 Then
            If TempPlayer(Index).Step = 0 Then TempPlayer(Index).Step = 2
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (TempPlayer(Index).YOffset > 8) Then Anim = TempPlayer(Index).Step
                Case DIR_DOWN
                    If (TempPlayer(Index).YOffset < -8) Then Anim = TempPlayer(Index).Step
                Case DIR_LEFT
                    If (TempPlayer(Index).xOffset > 8) Then Anim = TempPlayer(Index).Step
                Case DIR_RIGHT
                    If (TempPlayer(Index).xOffset < -8) Then Anim = TempPlayer(Index).Step
            End Select
            If TempPlayer(Index).Moving = MOVING_RUNNING Or TempPlayer(Index).Moving = -MOVING_RUNNING Then
                If TempPlayer(Index).Step = 1 Then Anim = 4
                If TempPlayer(Index).Step = 3 Then Anim = 5
            End If
        Else
            Anim = 18
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (TempPlayer(Index).YOffset > 8) Then Anim = 19
                Case DIR_DOWN
                    If (TempPlayer(Index).YOffset < -8) Then Anim = 19
                Case DIR_LEFT
                    If (TempPlayer(Index).xOffset > 8) Then Anim = 19
                Case DIR_RIGHT
                    If (TempPlayer(Index).xOffset < -8) Then Anim = 19
            End Select
            If TempPlayer(Index).Moving = MOVING_RUNNING Or TempPlayer(Index).Moving = -MOVING_RUNNING Then
                Anim = 20
            End If
            
            If Index = MyIndex Then
                If DirUp = True Or DirDown = True Or DirLeft = True Or DirRight = True Then
                Anim = 19
                If ShiftDown = True Then Anim = 20
                End If
            End If
        End If
    End If
    
    If TempPlayer(Index).SpellBufferNum > 0 Then
        If TempPlayer(Index).SpellBufferTimer + (Spell(TempPlayer(Index).SpellBufferNum).CastTime * 1000) > GetTickCount Then
            If Spell(TempPlayer(Index).SpellBufferNum).CastPlayerAnim > 0 Then
                
                Select Case Spell(TempPlayer(Index).SpellBufferNum).CastPlayerAnim
                    
                    Case 1 'kamehameha
                        Anim = 10
                        
                        If ((TempPlayer(Index).SpellBufferTimer + (Spell(TempPlayer(Index).SpellBufferNum).CastTime * 1000))) - GetTickCount < 100 Then
                            TempPlayer(Index).KamehamehaLast = GetTickCount
                        End If
                    
                    Case 2 'Spirit bomb
                        Anim = 15
                        
                        If ((TempPlayer(Index).SpellBufferTimer + (Spell(TempPlayer(Index).SpellBufferNum).CastTime * 1000))) - GetTickCount < 100 Then
                            TempPlayer(Index).SpiritBombLast = GetTickCount
                        End If
                
                    Case 3 'big bang attack
                        Anim = 14
                    
                    Case 4 'Super sayan
                    
                    If (GetTickCount - TempPlayer(Index).SpellBufferTimer) < 300 Then
                        Anim = 12
                        Else
                        Anim = 13
                    End If
                
                End Select
                
            End If
        End If
    End If
    
    If TempPlayer(Index).SpiritBombLast + 300 > GetTickCount Then
        Anim = 14
    End If
    
    If TempPlayer(Index).KamehamehaLast + 300 > GetTickCount Then
        Anim = 11
    End If

    ' Check to see if we want to stop making him attack
    With TempPlayer(Index)
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = spritetop * (Tex_Character(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Character(Sprite).Height / 4)
        If VXFRAME = False Then
            .Left = Anim * (Tex_Character(Sprite).Width / 22)
            .Right = .Left + (Tex_Character(Sprite).Width / 22)
        Else
            .Left = Anim * (Tex_Character(Sprite).Width / 3)
            .Right = .Left + (Tex_Character(Sprite).Width / 3)
        End If
    End With

    ' Calculate the X
    If VXFRAME = False Then
        X = GetPlayerX(Index) * PIC_X + TempPlayer(Index).xOffset - ((Tex_Character(Sprite).Width / 22 - 32) / 2)
    Else
        X = GetPlayerX(Index) * PIC_X + TempPlayer(Index).xOffset - ((Tex_Character(Sprite).Width / 3 - 32) / 2)
    End If
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + TempPlayer(Index).YOffset - ((Tex_Character(Sprite).Height / 4) - (Tex_Character(Sprite).Height / 8))
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + TempPlayer(Index).YOffset - ((Tex_Character(Sprite).Height / 4) - 64)
    End If
    
    'novas sprites
    Y = Y + 16
    
    ' render player shadow
    If Not hideNames Then RenderTexture Tex_Shadow, ConvertMapX(X), ConvertMapY(Y + 18), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    
    If TempPlayer(Index).Fly = 1 Then
        Y = Y - 16
    End If
    
    If Player(Index).Trans > 0 Then
        If Spell(Player(Index).Trans).SpriteTrans > 0 Then
            Dim Reded As Byte
            Reded = Spell(Player(Index).Trans).SpriteTrans
        End If
    End If
    
    If Not GetPlayerDir(Index) = DIR_UP Then
        If TempPlayer(Index).HairChange < 5 Then
            If Reded < 255 Then
                RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), ConvertMapX(X), ConvertMapY(Y) + ((rec.Bottom - rec.Top) / 2), rec.Left, rec.Top + ((rec.Bottom - rec.Top) / 2), rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(255, 255 - Reded, 255 - Reded, 255)
            Else
                RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), ConvertMapX(X), ConvertMapY(Y) + ((rec.Bottom - rec.Top) / 2), rec.Left, rec.Top + ((rec.Bottom - rec.Top) / 2), rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(0, 150, 255, 255)
            End If
        End If
    End If
    
    ' render the actual sprite
    If GetTickCount > TempPlayer(Index).StartFlash Then
        Call DrawSprite(Sprite, X, Y, rec, False, Reded)
        TempPlayer(Index).StartFlash = 0
    Else
        Call DrawSprite(Sprite, X, Y, rec, True, Reded)
    End If
    
    ' check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
            End If
        End If
    Next
    
    If ScouterOn And Index = MyIndex Then Call DrawPaperdoll(X, Y, ScouterPaperdoll, Anim, spritetop)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    If TempPlayer(Index).HairChange < 5 Then
        If Reded < 255 Then
            RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(255, 255 - Reded, 255 - Reded, 255)
        Else
            RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(0, 150, 255, 255)
        End If
    End If
    
    If GetPlayerDir(Index) = DIR_UP Then
        If TempPlayer(Index).HairChange < 5 Then
            If Reded < 255 Then
                RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), X, Y + ((rec.Bottom - rec.Top) / 2), rec.Left, rec.Top + ((rec.Bottom - rec.Top) / 2), rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(255, 255 - Reded, 255 - Reded, 255)
            Else
                RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), X, Y + ((rec.Bottom - rec.Top) / 2), rec.Left, rec.Top + ((rec.Bottom - rec.Top) / 2), rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(0, 150, 255, 255)
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayer", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte, i As Long, X As Long, Y As Long, Sprite As Long, spritetop As Long
Dim rec As RECT
Dim AttackSpeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).Num = 0 Or TempMapNpc(MapNpcNum).SpawnDelay > 0 Then Exit Sub ' no npc set
    
    If Npc(MapNpc(MapNpcNum).Num).GFXPack > 0 Then
        HandleGFXPack MapNpcNum
        Exit Sub
    End If
    
    Sprite = Npc(MapNpc(MapNpcNum).Num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    AttackSpeed = Npc(MapNpc(MapNpcNum).Num).AttackSpeed
    
    If AttackSpeed < 100 Then AttackSpeed = 1000

    ' Reset frame
    Anim = 0
    
    If Npc(MapNpc(MapNpcNum).Num).Fly = 1 Then
        If MapNpc(MapNpcNum).FlyOffSetTick + 100 < GetTickCount Then
            If MapNpc(MapNpcNum).FlyOffsetDir = 0 Then
                MapNpc(MapNpcNum).FlyOffSet = MapNpc(MapNpcNum).FlyOffSet - 1
                If MapNpc(MapNpcNum).FlyOffSet <= -5 Then MapNpc(MapNpcNum).FlyOffsetDir = 1
            Else
                MapNpc(MapNpcNum).FlyOffSet = MapNpc(MapNpcNum).FlyOffSet + 1
                If MapNpc(MapNpcNum).FlyOffSet >= 5 Then MapNpc(MapNpcNum).FlyOffsetDir = 0
            End If
            MapNpc(MapNpcNum).FlyOffSetTick = GetTickCount
        End If
        If GetTickCount Mod Npc(MapNpc(MapNpcNum).Num).FlyTick < Npc(MapNpc(MapNpcNum).Num).FlyTick / 2 Then
            Anim = 1
        Else
            Anim = 0
        End If
    End If
    
    ' Check for attacking animation
    If TempMapNpc(MapNpcNum).AttackTimer + (AttackSpeed / 2) > GetTickCount Then
        If TempMapNpc(MapNpcNum).Attacking = 1 Then
            If VXFRAME = False Then
                Anim = 3
            Else
                Anim = 2
            End If
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (TempMapNpc(MapNpcNum).YOffset > 8) Then Anim = TempMapNpc(MapNpcNum).Step
            Case DIR_DOWN
                If (TempMapNpc(MapNpcNum).YOffset < -8) Then Anim = TempMapNpc(MapNpcNum).Step
            Case DIR_LEFT
                If (TempMapNpc(MapNpcNum).xOffset > 8) Then Anim = TempMapNpc(MapNpcNum).Step
            Case DIR_RIGHT
                If (TempMapNpc(MapNpcNum).xOffset < -8) Then Anim = TempMapNpc(MapNpcNum).Step
        End Select
        
        If Npc(MapNpc(MapNpcNum).Num).Fly = 1 Then
            If GetTickCount Mod Npc(MapNpc(MapNpcNum).Num).FlyTick < Npc(MapNpc(MapNpcNum).Num).FlyTick / 2 Then
                    Anim = 1
                Else
                    Anim = 0
            End If
        End If
    End If

    ' Check to see if we want to stop making him attack
    With TempMapNpc(MapNpcNum)
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (Tex_Character(Sprite).Height / 4) * spritetop
        .Bottom = .Top + Tex_Character(Sprite).Height / 4
        If VXFRAME = False Then
            .Left = Anim * (Tex_Character(Sprite).Width / 4)
            .Right = .Left + (Tex_Character(Sprite).Width / 4)
        Else
            .Left = Anim * (Tex_Character(Sprite).Width / 3)
            .Right = .Left + (Tex_Character(Sprite).Width / 3)
        End If
    End With

    ' Calculate the X
    If VXFRAME = False Then
        X = MapNpc(MapNpcNum).X * PIC_X + TempMapNpc(MapNpcNum).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)
    Else
        X = MapNpc(MapNpcNum).X * PIC_X + TempMapNpc(MapNpcNum).xOffset - ((Tex_Character(Sprite).Width / 3 - 32) / 2)
    End If
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + TempMapNpc(MapNpcNum).YOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + TempMapNpc(MapNpcNum).YOffset
    End If
    
    ' render player shadow
    If Npc(MapNpc(MapNpcNum).Num).Shadow = 1 Then RenderTexture Tex_Shadow, ConvertMapX(X) - (MapNpc(MapNpcNum).FlyOffSet / 2), ConvertMapY(Y + 18), 0, 0, (Tex_Character(Sprite).Width / 4) + MapNpc(MapNpcNum).FlyOffSet, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    
    If Npc(MapNpc(MapNpcNum).Num).Fly = 1 Then
        Y = Y + MapNpc(MapNpcNum).FlyOffSet
    End If
    
    ' render the actual sprite
    If GetTickCount > TempMapNpc(MapNpcNum).StartFlash Then
        If TempMapNpc(MapNpcNum).StunDuration > 0 Then
            If GetTickCount < TempMapNpc(MapNpcNum).StunTick + (TempMapNpc(MapNpcNum).StunDuration * 1000) Then
                Call DrawSprite(Sprite, X, Y, rec, True)
                Exit Sub
            End If
        End If
        Call DrawSprite(Sprite, X, Y, rec)
        TempMapNpc(MapNpcNum).StartFlash = 0
    Else
        Call DrawSprite(Sprite, X, Y, rec, True)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpc", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPaperdoll(ByVal X2 As Long, ByVal Y2 As Long, ByVal Sprite As Long, ByVal Anim As Long, ByVal spritetop As Long, Optional NotInMap As Boolean = False)
Dim rec As RECT
Dim X As Long, Y As Long
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    With rec
        .Top = spritetop * (Tex_Paperdoll(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Paperdoll(Sprite).Height / 4)
        If VXFRAME = False Then
            .Left = Anim * (Tex_Paperdoll(Sprite).Width / 22)
            .Right = .Left + (Tex_Paperdoll(Sprite).Width / 22)
        Else
            .Left = Anim * (Tex_Paperdoll(Sprite).Width / 3)
            .Right = .Left + (Tex_Paperdoll(Sprite).Width / 3)
        End If
    End With
    
    ' clipping
    If Not NotInMap Then
        X = ConvertMapX(X2)
        Y = ConvertMapY(Y2)
    Else
        X = X2
        Y = Y2
    End If
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' Clip to screen
    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If
    
    RenderTexture Tex_Paperdoll(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawSprite(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, rec As RECT, Optional Flash As Boolean = False, Optional Reded As Byte = 0)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    
    If Flash = True Then
        RenderTexture Tex_Character(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255 - Reded, 255 - Reded, 150)
    Else
        RenderTexture Tex_Character(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255 - Reded, 255 - Reded, 255)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawFog()
Dim fogNum As Long, color As Long, X As Long, Y As Long, RenderState As Long
    
    If InMapEditor = True Then
        CurrentFog = Map.Fog
        CurrentFogOpacity = Map.FogOpacity
        CurrentFogSpeed = Map.FogSpeed
    End If
    
    fogNum = CurrentFog
    If fogNum <= 0 Or fogNum > NumFogs Then Exit Sub
    color = D3DColorRGBA(255, 255, 255, 255 - CurrentFogOpacity)

    RenderState = 0
    ' render state
    Select Case RenderState
        Case 1 ' Additive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    
    For X = 0 To ((Map.MaxX * 32) / 256) + 1
        For Y = 0 To ((Map.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), ConvertMapX((X * 256) + fogOffsetX), ConvertMapY((Y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, color
        Next
    Next
    
    ' reset render state
    If RenderState > 0 Then
        Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Public Sub DrawTint()
Dim color As Long, ModG As Byte, ModA As Byte
    If InMapEditor = True Then
        CurrentTintR = Map.Red
        CurrentTintG = Map.Green
        CurrentTintB = Map.Blue
        CurrentTintA = Map.Alpha
    End If
    
    If ScouterOn = True Then
        ModA = 50
        ModG = 150
    End If
    
    If CurrentTintG + ModG > 255 Then ModG = 255 - CurrentTintG
    If CurrentTintA + ModA > 255 Then ModA = 255 - CurrentTintA
    
    color = D3DColorRGBA(CurrentTintR, CurrentTintG + ModG, CurrentTintB, CurrentTintA + ModA)
    RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, color
End Sub

Public Sub DrawWeather()
Dim color As Long, i As Long, SpriteLeft As Long, X As Long, Y As Long
    If Map.Weather = WEATHER_TYPE_CLOUDS Then
        For i = 1 To 10
            If Cloud(i).Use = 0 Then
                If Rand(1, 500) = 1 Then
                    Cloud(i).Use = 1
                    Cloud(i).X = 0
                    Cloud(i).Y = Rand(0, Map.MaxY * 32)
                    Cloud(i).Anim = Rand(0, 1)
                    Cloud(i).speed = Rand(1, 2)
                    Cloud(i).SizeX = Rand(0, 48)
                    Cloud(i).SizeY = Rand(0, 48)
                    Cloud(i).Alpha = Rand(100, 200)
                End If
            End If
            If Cloud(i).Use = 1 Then
                Cloud(i).X = Cloud(i).X + Cloud(i).speed
                
                X = ConvertMapX(Cloud(i).X)
                Y = ConvertMapY(Cloud(i).Y)
                
                If X > ScreenWidth Then Cloud(i).Use = 0
                
                RenderTexture Tex_Clouds, X, Y, 0, Cloud(i).Anim * 96, 96 + Cloud(i).SizeX, 96 + Cloud(i).SizeY, 96, 96, D3DColorRGBA(255, 255, 255, Cloud(i).Alpha)
            End If
        Next i
        Exit Sub
    End If
    
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).Type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(i).Type - 1
            End If
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(i).X), ConvertMapY(WeatherParticle(i).Y), SpriteLeft * 32, 0, WeatherParticle(i).Size, WeatherParticle(i).Size, 32, 32, -1
        End If
    Next
End Sub

Sub DrawAnimatedInvItems()
Dim i As Long
Dim itemNum As Long, ItemPic As Long
Dim X As Long, Y As Long
Dim MaxFrames As Byte
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For i = 1 To MAX_MAP_ITEMS

        If MapItem(i).Num > 0 Then
            ItemPic = Item(MapItem(i).Num).Pic

            If ItemPic < 1 Or ItemPic > numitems Then Exit Sub
            MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(i).Frame < MaxFrames - 1 Then
                MapItem(i).Frame = MapItem(i).Frame + 1
            Else
                MapItem(i).Frame = 1
            End If
        End If

    Next

    For i = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, i)

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic

            If ItemPic > 0 And ItemPic <= numitems Then
                If Tex_Item(ItemPic).Width > 64 Then
                    MaxFrames = (Tex_Item(ItemPic).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(i) < MaxFrames - 1 Then
                        InvItemFrame(i) = InvItemFrame(i) + 1
                    Else
                        InvItemFrame(i) = 1
                    End If

                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = (Tex_Item(ItemPic).Width / 2) + (InvItemFrame(i) * 32) ' middle to get the start of inv gfx, then +32 for each frame
                        .Right = .Left + 32
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' We'll now re-Draw the item, and place the currency value over it again :P
                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                        ' Draw currency but with k, m, b etc. using a convertion function
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, Yellow, 0
                    End If
                End If
            End If
        End If

    Next

    'frmMain.picInventory.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawAnimatedInvItems", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_DrawTileset()
Dim Height As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Width As Long
Dim Tileset As Long
Dim sRECT As RECT
Dim dRect As RECT, scrlX As Long, scrlY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    scrlX = frmEditor_Map.scrlPictureX.Value * PIC_X
    scrlY = frmEditor_Map.scrlPictureY.Value * PIC_Y
    
    Height = Tex_Tileset(Tileset).Height - scrlY
    Width = Tex_Tileset(Tileset).Width - scrlX
    
    sRECT.Left = frmEditor_Map.scrlPictureX.Value * PIC_X
    sRECT.Top = frmEditor_Map.scrlPictureY.Value * PIC_Y
    sRECT.Right = sRECT.Left + Width
    sRECT.Bottom = sRECT.Top + Height
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    RenderTextureByRects Tex_Tileset(Tileset), sRECT, dRect
    
    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.Value > 0 Then
        Select Case frmEditor_Map.scrlAutotile.Value
            Case 1 ' autotile
                EditorTileWidth = 2
                EditorTileHeight = 3
            Case 2 ' fake autotile
                EditorTileWidth = 1
                EditorTileHeight = 1
            Case 3 ' animated
                EditorTileWidth = 6
                EditorTileHeight = 3
            Case 4 ' cliff
                EditorTileWidth = 2
                EditorTileHeight = 2
            Case 5 ' waterfall
                EditorTileWidth = 2
                EditorTileHeight = 3
        End Select
    End If
    
    With destRect
        .X1 = (EditorTileX * 32) - sRECT.Left
        .X2 = (EditorTileWidth * 32) + .X1
        .Y1 = (EditorTileY * 32) - sRECT.Top
        .Y2 = (EditorTileHeight * 32) + .Y1
    End With
    
    DrawSelectionBox destRect
        
    With srcRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Map.picBack.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Map.picBack.ScaleHeight
    End With
    
    'Now render the selection tiles and we are done!
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picBack.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawTileset", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawSelectionBox(dRect As D3DRECT)
Dim Width As Long, Height As Long, X As Long, Y As Long
    Width = dRect.X2 - dRect.X1
    Height = dRect.Y2 - dRect.Y1
    X = dRect.X1
    Y = dRect.Y1
    If Width > 6 And Height > 6 Then
        'Draw Box 32 by 32 at graphicselx and graphicsely
        RenderTexture Tex_Selection, X, Y, 1, 1, 2, 2, 2, 2, -1 'top left corner
        RenderTexture Tex_Selection, X + 2, Y, 3, 1, Width - 4, 2, 32 - 6, 2, -1 'top line
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y, 29, 1, 2, 2, 2, 2, -1 'top right corner
        RenderTexture Tex_Selection, X, Y + 2, 1, 3, 2, Height - 4, 2, 32 - 6, -1 'Left Line
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2, 32 - 3, 3, 2, Height - 4, 2, 32 - 6, -1 'right line
        RenderTexture Tex_Selection, X, Y + 2 + (Height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 'bottom left corner
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2 + (Height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 'bottom right corner
        RenderTexture Tex_Selection, X + 2, Y + 2 + (Height - 4), 3, 32 - 3, Width - 4, 2, 32 - 6, 2, -1 'bottom line
    End If
End Sub

Public Sub DrawTileOutline()
Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optBlock.Value Then Exit Sub

    With rec
        .Top = EditorTileY * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = EditorTileX * PIC_X
        .Right = .Left + PIC_X
    End With
    
    RenderTexture Tex_Tileset(frmEditor_Map.scrlTileSet.Value), ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorARGB(200, 255, 255, 255)
    RenderTexture Tex_Misc, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), 0, 0, 32, 32, 32, 32

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTileOutline", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawMapItem()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = Item(frmEditor_Map.scrlMapItem.Value).Pic

    If itemNum < 1 Or itemNum > numitems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(itemNum), sRECT, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picMapItem.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawMapItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawItem()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = frmEditor_Item.scrlPic.Value

    If itemNum < 1 Or itemNum > numitems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(itemNum), sRECT, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picItem.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawProjectile()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = frmEditor_Item.scrlProjectileNum.Value

    If itemNum < 1 Or itemNum > NumProjectiles Then
        frmEditor_Item.picProjectile.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Projectile(itemNum), sRECT, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picProjectile.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawProjectile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawPaperdoll()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRECT As RECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'frmEditor_Item.picPaperdoll.Cls
    
    Sprite = frmEditor_Item.scrlPaperdoll.Value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = Tex_Paperdoll(Sprite).Height / 4
    sRECT.Left = 0
    If VXFRAME = False Then
        sRECT.Right = Tex_Paperdoll(Sprite).Width / 4
    Else
        sRECT.Right = Tex_Paperdoll(Sprite).Width / 3
    End If
    ' same for destination as source
    dRect = sRECT
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Paperdoll(Sprite), sRECT, dRect
                    
    With destRect
        .X1 = 0
        If VXFRAME = False Then
            .X2 = Tex_Paperdoll(Sprite).Width / 4
        Else
            .X2 = Tex_Paperdoll(Sprite).Width / 3
        End If
        .Y1 = 0
        .Y2 = Tex_Paperdoll(Sprite).Height / 4
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picPaperdoll.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_DrawIcon()
Dim iconNum As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    iconNum = frmEditor_Spell.scrlIcon.Value
    
    If iconNum < 1 Or iconNum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_SpellIcon(iconNum), sRECT, dRect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Spell.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_DrawIcon", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_DrawProjectile()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = frmEditor_Spell.scrlProjectile.Value

    If itemNum < 1 Or itemNum > NumProjectiles Then
        frmEditor_Spell.picProjectile.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Projectile(itemNum), sRECT, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Spell.picProjectile.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawProjectile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Public Sub EditorAnim_DrawAnim()
Dim i As Long, Animationnum As Long, ShouldRender As Boolean, Width As Long, Height As Long, looptime As Long, FrameCount As Long
Dim sx As Long, sY As Long, sRECT As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    sRECT.Top = 0
    sRECT.Bottom = 192
    sRECT.Left = 0
    sRECT.Right = 192

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).Value
        
        If Animationnum <= 0 Or Animationnum > NumAnimations Then
            ' don't render lol
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then

                    Width = 192
                    Height = 192

                    sY = (Height * ((AnimEditorFrame(i) - 1) \ AnimColumns))
                    sx = (Width * (((AnimEditorFrame(i) - 1) Mod AnimColumns)))

                    ' Start Rendering
                    Call Direct3D_Device.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call Direct3D_Device.BeginScene
                    
                    'EngineRenderRectangle Tex_Anim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    RenderTexture Tex_Animation(Animationnum), 0, 0, sx, sY, Width, Height, Width, Height
                    
                    ' Finish Rendering
                    Call Direct3D_Device.EndScene
                    Call Direct3D_Device.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(i).hWnd, ByVal 0)
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_DrawAnim", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_DrawProjectile()
Dim itemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = frmEditor_NPC.scrlProjectile.Value

    If itemNum < 1 Or itemNum > NumProjectiles Then
        frmEditor_NPC.picProjectile.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Projectile(itemNum), sRECT, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_NPC.picProjectile.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_DrawProjectile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_DrawSprite()
Dim Sprite As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_NPC.scrlSprite.Value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    sRECT.Top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = PIC_X * 3 ' facing down
    sRECT.Right = sRECT.Left + SIZE_X
    dRect.Top = 0
    dRect.Bottom = SIZE_Y
    dRect.Left = 0
    dRect.Right = SIZE_X
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(Sprite), sRECT, dRect
    
    With destRect
        .X1 = 0
        .X2 = SIZE_X
        .Y1 = 0
        .Y2 = SIZE_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_NPC.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_DrawSprite()
Dim Sprite As Long
Dim sRECT As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(Sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRECT, dRect
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(Sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(Sprite).Height
        End With
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picNormalPic.hWnd, ByVal (0)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(Sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRECT, dRect
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
        
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(Sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(Sprite).Height
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picExhaustedPic.hWnd, ByVal (0)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub Render_Graphics()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim rec As RECT
Dim rec_pos As RECT, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
   If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'Check for device lost.
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
    
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    
    If GettingMap Then Exit Sub
    
    ' update the viewpoint
    UpdateCamera

    ' unload any textures we need to unload
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
        
        On Error Resume Next
        Direct3D_Device.BeginScene
        
            If Map.Panorama > 0 Then
                RenderTexture Tex_Panorama(Map.Panorama), ParallaxX, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, frmMain.ScaleWidth, frmMain.ScaleHeight
                RenderTexture Tex_Panorama(Map.Panorama), ParallaxX + frmMain.ScaleWidth, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, frmMain.ScaleWidth, frmMain.ScaleHeight
            End If
            ' blit lower tiles
            If NumTileSets > 0 Then
                For X = TileView.Left To TileView.Right
                    For Y = TileView.Top To TileView.Bottom
                        If IsValidMapPoint(X, Y) Then
                            Call DrawMapTile(X, Y)
                        End If
                    Next
                Next
            End If
            
            DrawBuracos
        
            ' render the decals
            For i = 1 To MAX_BYTE
                Call DrawBlood(i)
            Next
        
            ' Blit out the items
            If numitems > 0 Then
                For i = 1 To MAX_MAP_ITEMS
                    If MapItem(i).Num > 0 Then
                        Call DrawItem(i)
                    End If
                Next
            End If
            
            UpdateEffectAll
            
            ' draw animations
            If NumAnimations > 0 Then
                For i = 1 To MAX_BYTE
                    If AnimInstance(i).Used(0) Then
                        DrawAnimation i, 0
                    End If
                Next
            End If
            
            ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
            For Y = 0 To Map.MaxY
                If NumCharacters > 0 Then
                    
                    ' Players
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                            If Player(i).Y = Y Then
                                If TempPlayer(i).Fly = 0 Then
                                    Call DrawPlayer(i)
                                    If Map.Tile(GetPlayerX(i), GetPlayerY(i)).Type = TILE_TYPE_RESOURCE Then
                                        If i <> MyIndex Then Call DrawFishAlert(i)
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                
                    ' Npcs
                    For i = 1 To Npc_HighIndex
                        If MapNpc(i).Y = Y Then
                            Call DrawNpc(i)
                        End If
                    Next
                End If
            Next
            
            If NumProjectiles > 0 Then
                Call DrawProjectile
            End If
            
            ' animations
            If NumAnimations > 0 Then
                For i = 1 To MAX_BYTE
                    If AnimInstance(i).Used(1) Then
                        DrawAnimation i, 1
                    End If
                Next
            End If
        
            ' blit out upper tiles
            If NumTileSets > 0 Then
                For X = TileView.Left To TileView.Right
                    For Y = TileView.Top To TileView.Bottom
                        If IsValidMapPoint(X, Y) Then
                            Call DrawMapFringeTile(X, Y)
                        End If
                    Next
                Next
            End If
            
            ' Resources
            For Y = 1 To Map.MaxY
                If NumResources > 0 Then
                    If Resources_Init Then
                        If Resource_Index > 0 Then
                            For i = 1 To Resource_Index
                                If MapResource(i).Y = Y Then
                                    Call DrawMapResource(i)
                                End If
                            Next
                        End If
                    End If
                End If
            Next Y
            
            If Transporte.Tipo <> 0 Then Call DrawTransporte
            
            If NumCharacters > 0 Then
                ' Players flying
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If TempPlayer(i).Fly = 1 Then
                            Call DrawPlayer(i)
                        End If
                    End If
                Next
            End If
            
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_RESOURCE Then
                Call DrawFishAlert(MyIndex)
            Else
                FishingTime = 0
                BubbleOpaque = 0
            End If
            
            DrawWeather
            DrawFog
            DrawTint
            DrawAmbient
            
            ' blit out a square at mouse cursor
            If InMapEditor Then
                If frmEditor_Map.optBlock.Value = True Then
                    For X = TileView.Left To TileView.Right
                        For Y = TileView.Top To TileView.Bottom
                            If IsValidMapPoint(X, Y) Then
                                Call DrawDirection(X, Y)
                            End If
                        Next
                    Next
                End If
                If frmEditor_Map.chkGrid.Value = YES Or frmEditor_Map.optBlock.Value = True Then
                    For X = TileView.Left To TileView.Right
                        For Y = TileView.Top To TileView.Bottom
                            If IsValidMapPoint(X, Y) Then
                                Call DrawMapGrid(X, Y)
                            End If
                        Next
                    Next
                End If
                Call DrawTileOutline
            End If
            
            ' Render the bars
            DrawBars
            
            ' Draw the target icon
            If myTarget > 0 Then
                If myTargetType = TARGET_TYPE_PLAYER Then
                    DrawTarget (Player(myTarget).X * 32) + TempPlayer(myTarget).xOffset, (Player(myTarget).Y * 32) + TempPlayer(myTarget).YOffset
                ElseIf myTargetType = TARGET_TYPE_NPC Then
                    DrawTarget (MapNpc(myTarget).X * 32) + TempMapNpc(myTarget).xOffset, (MapNpc(myTarget).Y * 32) + TempMapNpc(myTarget).YOffset
                End If
            End If
            
            ' Draw the hover icon
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If Player(i).Map = Player(MyIndex).Map Then
                        If CurX = Player(i).X And CurY = Player(i).Y Then
                            If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                                ' dont render lol
                            Else
                                DrawHover TARGET_TYPE_PLAYER, i, (Player(i).X * 32) + TempPlayer(i).xOffset, (Player(i).Y * 32) + TempPlayer(i).YOffset
                            End If
                        End If
                    End If
                End If
            Next
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Num > 0 Then
                    If CurX = MapNpc(i).X And CurY = MapNpc(i).Y Then
                        If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                            ' dont render lol
                        Else
                            DrawHover TARGET_TYPE_NPC, i, (MapNpc(i).X * 32) + TempMapNpc(i).xOffset, (MapNpc(i).Y * 32) + TempMapNpc(i).YOffset
                        End If
                    End If
                End If
            Next
            
            If DrawThunder > 0 Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160): DrawThunder = DrawThunder - 1
            
            ' Get rec
            With rec
                .Top = Camera.Top
                .Bottom = .Top + ScreenY
                .Left = Camera.Left
                .Right = .Left + ScreenX
            End With
                
            ' rec_pos
            With rec_pos
                .Bottom = ScreenY
                .Right = ScreenX
            End With
                
            With srcRect
                .X1 = 0
                .X2 = frmMain.ScaleWidth
                .Y1 = 0
                .Y2 = frmMain.ScaleHeight
            End With
            
            If Not hideNames Then
                ' draw player names
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        Call DrawPlayerName(i)
                    End If
                Next
                
                ' draw npc names
                For i = 1 To Npc_HighIndex
                    If MapNpc(i).Num > 0 Then
                        Call DrawNpcName(i)
                    End If
                Next
            End If
            
                ' draw the messages
            For i = 1 To MAX_BYTE
                If chatBubble(i).active Then
                    DrawChatBubble i
                End If
            Next
            
            For i = 1 To Action_HighIndex
                Call DrawActionMsg(i)
            Next i
            If Not hideGUI Then DrawGUI
            
            If BFPS = True Then
                RenderText Font_Default, "FPS: " & CStr(GameFPS) & " Ping: " & CStr(Ping), 12, 100, Yellow, 0
                RenderText Font_Default, Trim$("cur x: " & CurX & " y: " & CurY), 12, 114, Yellow, 0
                RenderText Font_Default, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 12, 128, Yellow, 0
                RenderText Font_Default, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 12, 142, Yellow, 0
            End If
            
            If GettingMap Then
                RenderText Font_Default, "Recebendo mapa...", 64, 100, Yellow, 0
            End If
            
            If InMapEditor Then Call DrawMapAttributes
            
            If FadeAmount > 0 Then RenderTexture Tex_Fade, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, FadeAmount)
            If FlashTimer > GetTickCount Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, -1
        Direct3D_Device.EndScene
        
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        Direct3D_Device.Present srcRect, ByVal 0, 0, ByVal 0
        DrawGDI
    End If

    ' Error handler
    Exit Sub
    
errorhandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        If Options.Debug = 1 Then
            HandleError "Render_Graphics", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
            Err.Clear
        End If
        MsgBox "Unrecoverable DX8 error."
        DestroyGame
    End If
End Sub

Sub HandleDeviceLost()
'Do a loop while device is lost
   Do While Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST
       Exit Sub
   Loop
   
   UnloadTextures
   
   'Reset the device
   Direct3D_Device.Reset Direct3D_Window
   
   DirectX_ReInit
    
   LoadTextures
   
End Sub

Private Function DirectX_ReInit() As Boolean

    On Error GoTo Error_Handler

    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
        
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.

    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'we've already setup for Direct3D_Window.
    'Creates the rendering device with some useful info, along with the info
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = 800 ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = 600 'frmMain.Scaleheight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hWnd 'Use frmMain as the device window.
    
    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    End With
    
    DirectX_ReInit = True

    Exit Function
    
Error_Handler:
    MsgBox "An error occured while initializing DirectX", vbCritical
    
    DestroyGame
    
    DirectX_ReInit = False
End Function

Public Sub UpdateCamera()
Dim offsetX As Long
Dim offsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    offsetX = TempPlayer(MyIndex).xOffset + PIC_X
    offsetY = TempPlayer(MyIndex).YOffset + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        offsetX = 0
        If StartX = -1 Then
            If TempPlayer(MyIndex).xOffset > 0 Then
                offsetX = TempPlayer(MyIndex).xOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0
        If StartY = -1 Then
            If TempPlayer(MyIndex).YOffset > 0 Then
                offsetY = TempPlayer(MyIndex).YOffset
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.MaxX Then
        offsetX = 32
        If EndX = Map.MaxX + 1 Then
            If TempPlayer(MyIndex).xOffset < 0 Then
                offsetX = TempPlayer(MyIndex).xOffset + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.MaxY Then
        offsetY = 32
        If EndY = Map.MaxY + 1 Then
            If TempPlayer(MyIndex).YOffset < 0 Then
                offsetY = TempPlayer(MyIndex).YOffset + PIC_Y
            End If
        End If
        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .Bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = X - (TileView.Left * PIC_X) - Camera.Left
    
    If Tremor > GetTickCount Then
        ConvertMapX = ConvertMapX + TremorX
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = Y - (TileView.Top * PIC_Y) - Camera.Top
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.Top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    InViewPort = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTilesets()
Dim X As Long
Dim Y As Long
Dim i As Long
'Dim tilesetInUse() As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'ReDim tilesetInUse(0 To NumTileSets)
    
    'For X = 0 To Map.MaxX
    '    For Y = 0 To Map.MaxY
    '        For I = 1 To MapLayer.Layer_Count - 1
                ' check exists
    '            If Map.Tile(X, Y).layer(I).Tileset > 0 And Map.Tile(X, Y).layer(I).Tileset <= NumTileSets Then
    '                tilesetInUse(Map.Tile(X, Y).layer(I).Tileset) = True
    '            End If
    '        Next
    '    Next
    'Next
    
    'For I = 1 To NumTileSets
    '    If tilesetInUse(I) Then
        
    '    Else
            ' unload tileset
            'Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            'Set Tex_Tileset(i) = Nothing
    '    End If
    'Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'This function will make it much easier to setup the vertices with the info it needs.
Private Function Create_TLVertex(X As Single, Y As Single, Z As Single, RHW As Single, color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX

    Create_TLVertex.X = X
    Create_TLVertex.Y = Y
    Create_TLVertex.Z = Z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.color = color
    'Create_TLVertex.Specular = Specular
    Create_TLVertex.TU = TU
    Create_TLVertex.TV = TV
    
End Function

Public Function Ceiling(dblValIn As Double, dblCeilIn As Double) As Double
' round it
Ceiling = Round(dblValIn / dblCeilIn, 0) * dblCeilIn
' if it rounded down, force it up
If Ceiling < dblValIn Then Ceiling = Ceiling + dblCeilIn
End Function

Public Sub DestroyDX8()
    UnloadTextures
    Set Direct3DX = Nothing
    Set Direct3D_Device = Nothing
    Set Direct3D = Nothing
    Set DirectX8 = Nothing
End Sub

Public Sub DrawGDI()
    'Cycle Through in-game stuff before cycling through editors
    
    If frmEditor_Animation.visible Then
        EditorAnim_DrawAnim
    End If
    
    If frmEditor_Item.visible Then
        EditorItem_DrawItem
        EditorItem_DrawPaperdoll
        EditorItem_DrawProjectile
    End If
    
    If frmEditor_Map.visible Then
        EditorMap_DrawTileset
        If frmEditor_Map.fraMapItem.visible Then EditorMap_DrawMapItem
    End If
    
    If frmEditor_NPC.visible Then
        EditorNpc_DrawSprite
        EditorNpc_DrawProjectile
    End If
    
    If frmEditor_Resource.visible Then
        EditorResource_DrawSprite
    End If
    
    If frmEditor_Spell.visible Then
        EditorSpell_DrawIcon
        EditorSpell_DrawProjectile
    End If
    
    If frmEditor_Quest.visible Then
        EditorQuest_DrawIcon
    End If
End Sub
Public Sub DrawGUI()
Dim i As Long, X As Long, Y As Long
Dim Width As Long, Height As Long

    ' render shadow
    'EngineRenderRectangle Tex_GUI(27), 0, 0, 0, 0, 800, 64, 1, 64, 800, 64
    'EngineRenderRectangle Tex_GUI(26), 0, 600 - 64, 0, 0, 800, 64, 1, 64, 800, 64
    RenderTexture Tex_GUI(23), 0, 0, 0, 0, 800, 64, 1, 64
    RenderTexture Tex_GUI(22), 0, 600 - 64, 0, 0, 800, 64, 1, 64
    ' render chatbox
        If Not inChat Then
            If chatOn Then
                Width = 412
                Height = 145
                RenderTexture Tex_GUI(1), GUIWindow(GUI_CHAT).X, GUIWindow(GUI_CHAT).Y, 0, 0, Width, Height, Width, Height
                RenderText Font_Default, RenderChatText & chatShowLine, GUIWindow(GUI_CHAT).X + 38, GUIWindow(GUI_CHAT).Y + 126, White
                ' draw buttons
                For i = 34 To 35
                    ' set co-ordinate
                    X = GUIWindow(GUI_CHAT).X + Buttons(i).X
                    Y = GUIWindow(GUI_CHAT).Y + Buttons(i).Y
                    Width = Buttons(i).Width
                    Height = Buttons(i).Height
                    ' check for state
                    If Buttons(i).State = 2 Then
                        ' we're clicked boyo
                        'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                        RenderTexture Tex_Buttons_c(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                    ElseIf (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
                        ' we're hoverin'
                        'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                        RenderTexture Tex_Buttons_h(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                        ' play sound if needed
                        If Not lastButtonSound = i Then
                            PlaySound Sound_ButtonHover, -1, -1
                            lastButtonSound = i
                        End If
                    Else
                        ' we're normal
                        'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                        RenderTexture Tex_Buttons(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                        ' reset sound if needed
                        If lastButtonSound = i Then lastButtonSound = 0
                    End If
                Next
            Else
                RenderTexture Tex_GUI(1), GUIWindow(GUI_CHAT).X, GUIWindow(GUI_CHAT).Y + 123, 0, 123, 412, 22, 412, 22
            End If
            RenderChatTextBuffer
        Else
            If GUIWindow(GUI_CURRENCY).visible Then DrawCurrency
            If GUIWindow(GUI_EVENTCHAT).visible Then DrawEventChat
        End If
    
    DrawGUIBars
    DrawMapName
    
    ' render menu
    If GUIWindow(GUI_MENU).visible Then DrawMenu
    
    ' render hotbar
    If GUIWindow(GUI_HOTBAR).visible Then DrawHotbar
    
    If GUIWindow(GUI_NEWS).visible Then DrawNews
    
    ' render menus
    If GUIWindow(GUI_INVENTORY).visible Then DrawInventory
    If GUIWindow(GUI_SPELLS).visible Then DrawSkills
    If GUIWindow(GUI_CHARACTER).visible Then DrawCharacter
    If GUIWindow(GUI_OPTIONS).visible Then DrawOptions
    If GUIWindow(GUI_PARTY).visible Then DrawParty
    If GUIWindow(GUI_SHOP).visible Then DrawShop
    If GUIWindow(GUI_BANK).visible Then DrawBank
    If GUIWindow(GUI_TRADE).visible Then DrawTrade
    If GUIWindow(GUI_DIALOGUE).visible Then DrawDialogue
    If GUIWindow(GUI_QUESTS).visible Then DrawQuests
    
    ' Drag and drop
    DrawDragItem
    DrawDragSpell
    
    ' Descriptions
    DrawInventoryItemDesc
    DrawCharacterItemDesc
    DrawPlayerSpellDesc
    DrawBankItemDesc
    DrawTradeItemDesc
    DrawPlayerQuestDesc
End Sub


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   All of this code is for auto tiles and the math behind generating them.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub placeAutotile(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(X, Y).Layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .X = autoInner(1).X
                .Y = autoInner(1).Y
            Case "b"
                .X = autoInner(2).X
                .Y = autoInner(2).Y
            Case "c"
                .X = autoInner(3).X
                .Y = autoInner(3).Y
            Case "d"
                .X = autoInner(4).X
                .Y = autoInner(4).Y
            Case "e"
                .X = autoNW(1).X
                .Y = autoNW(1).Y
            Case "f"
                .X = autoNW(2).X
                .Y = autoNW(2).Y
            Case "g"
                .X = autoNW(3).X
                .Y = autoNW(3).Y
            Case "h"
                .X = autoNW(4).X
                .Y = autoNW(4).Y
            Case "i"
                .X = autoNE(1).X
                .Y = autoNE(1).Y
            Case "j"
                .X = autoNE(2).X
                .Y = autoNE(2).Y
            Case "k"
                .X = autoNE(3).X
                .Y = autoNE(3).Y
            Case "l"
                .X = autoNE(4).X
                .Y = autoNE(4).Y
            Case "m"
                .X = autoSW(1).X
                .Y = autoSW(1).Y
            Case "n"
                .X = autoSW(2).X
                .Y = autoSW(2).Y
            Case "o"
                .X = autoSW(3).X
                .Y = autoSW(3).Y
            Case "p"
                .X = autoSW(4).X
                .Y = autoSW(4).Y
            Case "q"
                .X = autoSE(1).X
                .Y = autoSE(1).Y
            Case "r"
                .X = autoSE(2).X
                .Y = autoSE(2).Y
            Case "s"
                .X = autoSE(3).X
                .Y = autoSE(3).Y
            Case "t"
                .X = autoSE(4).X
                .Y = autoSE(4).Y
        End Select
    End With
End Sub

Public Sub initAutotiles()
Dim X As Long, Y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).X = 32
    autoInner(1).Y = 0
    
    ' NE - b
    autoInner(2).X = 48
    autoInner(2).Y = 0
    
    ' SW - c
    autoInner(3).X = 32
    autoInner(3).Y = 16
    
    ' SE - d
    autoInner(4).X = 48
    autoInner(4).Y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).X = 0
    autoNW(1).Y = 32
    
    ' NE - f
    autoNW(2).X = 16
    autoNW(2).Y = 32
    
    ' SW - g
    autoNW(3).X = 0
    autoNW(3).Y = 48
    
    ' SE - h
    autoNW(4).X = 16
    autoNW(4).Y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).X = 32
    autoNE(1).Y = 32
    
    ' NE - g
    autoNE(2).X = 48
    autoNE(2).Y = 32
    
    ' SW - k
    autoNE(3).X = 32
    autoNE(3).Y = 48
    
    ' SE - l
    autoNE(4).X = 48
    autoNE(4).Y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).X = 0
    autoSW(1).Y = 64
    
    ' NE - n
    autoSW(2).X = 16
    autoSW(2).Y = 64
    
    ' SW - o
    autoSW(3).X = 0
    autoSW(3).Y = 80
    
    ' SE - p
    autoSW(4).X = 16
    autoSW(4).Y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).X = 32
    autoSE(1).Y = 64
    
    ' NE - r
    autoSE(2).X = 48
    autoSE(2).Y = 64
    
    ' SW - s
    autoSE(3).X = 32
    autoSE(3).Y = 80
    
    ' SE - t
    autoSE(4).X = 48
    autoSE(4).Y = 80
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                CalculateAutotile X, Y, layerNum
                ' cache the rendering state of the tiles and set them
                CacheRenderState X, Y, layerNum
            Next
        Next
    Next
End Sub

Public Sub CacheRenderState(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early
    If X < 0 Or X > Map.MaxX Or Y < 0 Or Y > Map.MaxY Then Exit Sub

    With Map.Tile(X, Y)
        ' check if the tile can be rendered
        If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > NumTileSets Then
            Autotile(X, Y).Layer(layerNum).RenderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Then
            ' default to... default
            Autotile(X, Y).Layer(layerNum).RenderState = RENDER_STATE_NORMAL
        Else
            Autotile(X, Y).Layer(layerNum).RenderState = RENDER_STATE_AUTOTILE
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(X, Y).Layer(layerNum).srcX(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).X * 32) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).X
                Autotile(X, Y).Layer(layerNum).srcY(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).Y * 32) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).Y
            Next
        End If
    End With
End Sub

Public Sub CalculateAutotile(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(X, Y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(X, Y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, X, Y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, X, Y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, X, Y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MaxX Or Y2 < 0 Or Y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' fakes ALWAYS return true
    If Map.Tile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If Map.Tile(X2, Y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If Map.Tile(X1, Y1).Layer(layerNum).Tileset <> Map.Tile(X2, Y2).Layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If Map.Tile(X1, Y1).Layer(layerNum).X <> Map.Tile(X2, Y2).Layer(layerNum).X Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(X1, Y1).Layer(layerNum).Y <> Map.Tile(X2, Y2).Layer(layerNum).Y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long)
Dim YOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case Map.Tile(X, Y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            YOffset = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            YOffset = -32
    End Select
    
    ' Draw the quarter
    'EngineRenderRectangle Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, 16, 16
    RenderTexture Tex_Tileset(Map.Tile(X, Y).Layer(layerNum).Tileset), destX, destY, Autotile(X, Y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(X, Y).Layer(layerNum).srcY(quarterNum) + YOffset, 16, 16, 16, 16, -1
End Sub

Public Sub DrawItem(ByVal itemNum As Long)
Dim PicNum As Integer, dontRender As Boolean, i As Long, tmpIndex As Long
Dim X As Long, Y As Long
Dim Left As Long

    

    If MapItem(itemNum).Gravity < 10 Then MapItem(itemNum).Gravity = MapItem(itemNum).Gravity + 1
    
    If MapItem(itemNum).YOffset + MapItem(itemNum).Gravity > 0 Then
        MapItem(itemNum).YOffset = 0
    Else
        MapItem(itemNum).YOffset = MapItem(itemNum).YOffset + MapItem(itemNum).Gravity
    End If
    
    X = MapItem(itemNum).X * PIC_X
    Y = (MapItem(itemNum).Y * PIC_Y) + MapItem(itemNum).YOffset

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    PicNum = Item(MapItem(itemNum).Num).Pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

     ' if it's not us then don't render
    If MapItem(itemNum).playerName <> vbNullString Then
        If Trim$(MapItem(itemNum).playerName) <> Trim$(GetPlayerName(MyIndex)) Then
            dontRender = True
        End If
        ' make sure it's not a party drop
        If Party.Leader > 0 Then
            For i = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party.Member(i)
                If tmpIndex > 0 Then
                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(itemNum).playerName) Then
                        dontRender = False
                    End If
                End If
            Next
        End If
    End If
    
    If Tex_Item(PicNum).Width > 96 Then
        If GetTickCount Mod 1000 <= 500 Then
            Left = 32
        Else
            Left = 0
        End If
    End If
    
    'If Not dontRender Then EngineRenderRectangle Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32, 32, 32
    If Not dontRender Then
        RenderTexture Tex_Item(PicNum), X, Y, Left, 0, 32, 32, 32, 32
    End If
End Sub

Public Sub DrawDragItem()
    Dim PicNum As Integer, itemNum As Long
    
    If DragInvSlotNum = 0 Then Exit Sub
    
    itemNum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)
    If Not itemNum > 0 Then Exit Sub
    
    PicNum = Item(itemNum).Pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

    'EngineRenderRectangle Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawDragSpell()
    Dim PicNum As Integer, spellnum As Long
    
    If DragSpell = 0 Then Exit Sub
    
    spellnum = PlayerSpells(DragSpell)
    If Not spellnum > 0 Then Exit Sub
    
    PicNum = Spell(spellnum).Icon

    If PicNum < 1 Or PicNum > NumSpellIcons Then Exit Sub

    'EngineRenderRectangle Tex_Spellicon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_SpellIcon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawHotbar()
Dim i As Long, X As Long, Y As Long, t As Long, sS As String
Dim Width As Long, Height As Long, color As Long

    X = GUIWindow(GUI_HOTBAR).X - 3
    Y = GUIWindow(GUI_HOTBAR).Y - 3
    Width = 493
    Height = 43
    RenderTexture Tex_GUI(31), X, Y, 0, 0, Width, Height, Width, Height

    For i = 1 To MAX_HOTBAR
        ' draw the box
        X = GUIWindow(GUI_HOTBAR).X + ((i - 1) * (5 + 36))
        Y = GUIWindow(GUI_HOTBAR).Y
        Width = 36
        Height = 36
        'EngineRenderRectangle Tex_GUI(2), x, y, 0, 0, width, height, width, height, width, heigh
        RenderTexture Tex_GUI(2), X, Y, 0, 0, Width, Height, Width, Height
        ' draw the icon
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(i).Slot).Name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 Then
                        'EngineRenderRectangle Tex_Item(Item(Hotbar(i).Slot).Pic), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        RenderTexture Tex_Item(Item(Hotbar(i).Slot).Pic), X + 2, Y + 2, 0, 0, 32, 32, 32, 32
                    End If
                End If
            Case 2 ' spell
                If Len(Spell(Hotbar(i).Slot).Name) > 0 Then
                    If Spell(Hotbar(i).Slot).Icon > 0 Then
                        ' render normal icon
                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        RenderTexture Tex_SpellIcon(Spell(Hotbar(i).Slot).Icon), X + 2, Y + 2, 0, 0, 32, 32, 32, 32
                        ' we got the spell?
                        For t = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(t) > 0 Then
                                If PlayerSpells(t) = Hotbar(i).Slot Then
                                    If SpellCD(t) > 0 Then
                                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                                        RenderTexture Tex_SpellIcon(Spell(Hotbar(i).Slot).Icon), X + 2, Y + 2, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
        End Select
        ' draw the numbers
        sS = str(i)
        If i = 10 Then sS = "0"
        If i = 11 Then sS = " -"
        If i = 12 Then sS = " ="
        RenderText Font_Default, sS, X + 4, Y + 20, White
    Next
End Sub
Public Sub DrawInventory()
Dim i As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long
Dim Amount As String
Dim colour As Long
Dim Top As Long, Left As Long
Dim Width As Long, Height As Long
Dim RenderLeft As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_INVENTORY).x, GUIWindow(GUI_INVENTORY).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_INVENTORY).X, GUIWindow(GUI_INVENTORY).Y, 0, 0, Width, Height, Width, Height
    
    For i = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, i)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    If TradeYourOffer(X).Num = i Then
                        GoTo NextLoop
                    End If
                Next
            End If
            
            ' exit out if dragging
            If DragInvSlotNum = i Then GoTo NextLoop

            If ItemPic > 0 And ItemPic <= numitems Then
                Top = GUIWindow(GUI_INVENTORY).Y + InvTop - 2 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                Left = GUIWindow(GUI_INVENTORY).X + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                
                RenderLeft = 0
                
                If Tex_Item(ItemPic).Width > 96 Then
                    If GetTickCount Mod 1000 <= 500 Then
                        RenderLeft = 32
                    Else
                        RenderLeft = 0
                    End If
                End If
                
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                RenderTexture Tex_Item(ItemPic), Left, Top, RenderLeft, 0, 32, 32, 32, 32
                ' If item is a stack - draw the amount you have
                If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                    Y = Top + 21
                    X = Left - 4
                    Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour
                End If
            End If
        End If
NextLoop:
    Next
End Sub

Public Sub DrawInventoryItemDesc()
Dim invSlot As Long, isSB As Boolean
    
    If Not GUIWindow(GUI_INVENTORY).visible Then Exit Sub
    If DragInvSlotNum > 0 Then Exit Sub
    
    invSlot = IsInvItem(GlobalX, GlobalY)
    If invSlot > 0 Then
        If GetPlayerInvItemNum(MyIndex, invSlot) > 0 Then
            'If Item(GetPlayerInvItemNum(MyIndex, invSlot)).BindType > 0 And PlayerInv(invSlot).bound > 0 Then isSB = True
            DrawItemDesc GetPlayerInvItemNum(MyIndex, invSlot), GUIWindow(GUI_INVENTORY).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_INVENTORY).Y, isSB
            ' value
            If InShop > 0 Then
                DrawItemCost False, invSlot, GUIWindow(GUI_INVENTORY).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_INVENTORY).Y + GUIWindow(GUI_DESCRIPTION).Height + 10
            End If
        End If
    End If
End Sub

Public Sub DrawShopItemDesc()
Dim shopSlot As Long
    
    If Not GUIWindow(GUI_SHOP).visible Then Exit Sub
    
    shopSlot = IsShopItem(GlobalX, GlobalY)
    If shopSlot > 0 Then
        If Shop(InShop).TradeItem(shopSlot).Item > 0 Then
            DrawItemDesc Shop(InShop).TradeItem(shopSlot).Item, GUIWindow(GUI_SHOP).X + GUIWindow(GUI_SHOP).Width + 10, GUIWindow(GUI_SHOP).Y
            DrawItemCost True, shopSlot, GUIWindow(GUI_SHOP).X + GUIWindow(GUI_SHOP).Width + 10, GUIWindow(GUI_SHOP).Y + GUIWindow(GUI_DESCRIPTION).Height + 90
        End If
    End If
End Sub

Public Sub DrawCharacterItemDesc()
Dim eqSlot As Long, isSB As Boolean
    
    If Not GUIWindow(GUI_CHARACTER).visible Then Exit Sub
    
    eqSlot = IsEqItem(GlobalX, GlobalY)
    If eqSlot > 0 Then
        If GetPlayerEquipment(MyIndex, eqSlot) > 0 Then
            If Item(GetPlayerEquipment(MyIndex, eqSlot)).BindType > 0 Then isSB = True
            DrawItemDesc GetPlayerEquipment(MyIndex, eqSlot), GUIWindow(GUI_CHARACTER).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_CHARACTER).Y, isSB
        End If
    End If
End Sub

Public Sub DrawItemCost(ByVal isShop As Boolean, ByVal slotNum As Long, ByVal X As Long, ByVal Y As Long)
Dim CostItem As Long, CostValue As Long, itemNum As Long, sString As String, Width As Long, Height As Long, i As Long

    If slotNum = 0 Then Exit Sub
    
    If InShop <= 0 Then Exit Sub
    
    For i = 1 To 5
    ' draw the window
    Width = 190
    Height = 36
    
    ' find out the cost
    If Not isShop Then
        ' inventory - default to gold
        itemNum = GetPlayerInvItemNum(MyIndex, slotNum)
        If itemNum = 0 Then Exit Sub
        CostItem = MoedaZ
        CostValue = (Item(itemNum).Price / 100) * Shop(InShop).BuyRate
        sString = "Será comprado por"
        If Item(itemNum).Price = 0 Then
            sString = "Este item não pode ser vendido!"
            RenderTexture Tex_GUI(24), X, Y + 80, 0, 0, Width, Height, Width, Height
            RenderText Font_Default, sString, X + 4, Y + 83, BrightRed
            Exit Sub
        End If
        Y = Y + 80
    Else
        itemNum = Shop(InShop).TradeItem(slotNum).Item
        If itemNum = 0 Then Exit Sub
        CostItem = Shop(InShop).TradeItem(slotNum).CostItem(i)
        CostValue = Shop(InShop).TradeItem(slotNum).CostValue(i)
        If CostItem = 0 Then Exit Sub
        If i = 1 Then
            sString = "Será trocado por"
        Else
            sString = "Tambem é necessário"
        End If
    End If
    
    RenderTexture Tex_GUI(24), X, Y + (36 * (i - 1)), 0, 0, Width, Height, Width, Height
    
    'EngineRenderRectangle Tex_Item(Item(CostItem).Pic), x + 155, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(Item(CostItem).Pic), X + 155, Y + 2 + (36 * (i - 1)), 0, 0, 32, 32, 32, 32
    
    RenderText Font_Default, sString, X + 4, Y + 3 + (36 * (i - 1)), DarkGrey
    
    RenderText Font_Default, ConvertCurrency(CostValue) & " " & Trim$(Item(CostItem).Name), X + 4, Y + 18 + (36 * (i - 1)), White
    Next i
End Sub

Public Sub DrawItemDesc(ByVal itemNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal soulBound As Boolean = False)
Dim colour As Long, descString As String, theName As String, className As String, levelTxt As String, sInfo() As String, i As Long, Width As Long, Height As Long

    ' get out
    If itemNum = 0 Then Exit Sub

    ' render the window
    Width = 190
    If Not Trim$(Item(itemNum).Desc) = vbNullString Then
        Height = 210
    Else
        Height = 126
    End If
    'EngineRenderRectangle Tex_GUI(6), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(8), X, Y, 0, 0, Width, Height, Width, Height
    
    ' make sure it has a sprite
    If Item(itemNum).Pic > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Item(Item(itemnum).Pic), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 64, 64
        RenderTexture Tex_Item(Item(itemNum).Pic), X + 16, Y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not Trim$(Item(itemNum).Desc) = vbNullString Then
        RenderText Font_Default, WordWrap(Trim$(Item(itemNum).Desc), Width - 10), X + 10, Y + 128, White
    End If
    ' work out name colour
    Select Case Item(itemNum).Rarity
        Case 0 ' white
            colour = White
        Case 1 ' green
            colour = Green
        Case 2 ' blue
            colour = Blue
        Case 3 ' maroon
            colour = Red
        Case 4 ' purple
            colour = Pink
        Case 5 ' orange
            colour = Brown
    End Select
    
    If Not soulBound Then
        theName = Trim$(Item(itemNum).Name)
    Else
        theName = "(SB) " & Trim$(Item(itemNum).Name)
    End If
    
    ' render name
    RenderText Font_Default, theName, X + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), Y + 6, colour
    
    ' class req
    If Item(itemNum).ClassReq > 0 Then
        className = Trim$(Class(Item(itemNum).ClassReq).Name)
        ' do we match it?
        If GetPlayerClass(MyIndex) = Item(itemNum).ClassReq Then
            colour = Green
        Else
            colour = BrightRed
        End If
    Else
        className = "Todas as raças."
        colour = Green
    End If
    RenderText Font_Default, className, X + 48 - (EngineGetTextWidth(Font_Default, className) \ 2), Y + 92, colour
    
    ' level
    If Item(itemNum).LevelReq > 0 Then
        levelTxt = "Nível " & Item(itemNum).LevelReq
        ' do we match it?
        If GetPlayerLevel(MyIndex) >= Item(itemNum).LevelReq Then
            colour = Green
        Else
            colour = BrightRed
        End If
    Else
        levelTxt = "Todos os níveis."
        colour = Green
    End If
    RenderText Font_Default, levelTxt, X + 48 - (EngineGetTextWidth(Font_Default, levelTxt) \ 2), Y + 107, colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    i = 1
    ReDim Preserve sInfo(1 To i) As String
    Select Case Item(itemNum).Type
        Case ITEM_TYPE_NONE
            sInfo(i) = "No type"
        Case ITEM_TYPE_WEAPON
            sInfo(i) = "Arma"
        Case ITEM_TYPE_ARMOR
            sInfo(i) = "Peitoral"
        Case ITEM_TYPE_HELMET
            sInfo(i) = "Calças"
        Case ITEM_TYPE_SHIELD
            sInfo(i) = "Botas"
        Case ITEM_TYPE_CONSUME
            sInfo(i) = "Consumo"
        Case ITEM_TYPE_CURRENCY
            sInfo(i) = "Dinheiro"
        Case ITEM_TYPE_SPELL
            sInfo(i) = "Tecnica"
    End Select
    
    ' more info
    Select Case Item(itemNum).Type
        Case ITEM_TYPE_NONE, ITEM_TYPE_CURRENCY
            ' binding
            If Item(itemNum).BindType = 1 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Trava ao pegar"
            ElseIf Item(itemNum).BindType = 2 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Trava ao equipar"
            End If
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Valor: " & Item(itemNum).Price & "z"
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SHIELD
            ' binding
            If Item(itemNum).BindType = 1 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Trava ao pegar"
            ElseIf Item(itemNum).BindType = 2 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Trava ao equipar"
            End If
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Valor: " & Item(itemNum).Price & "z"
            ' damage/defence
            If Item(itemNum).Type = ITEM_TYPE_WEAPON Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Dano: " & Item(itemNum).Data2
                ' speed
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Veloc.: " & (Item(itemNum).speed / 1000) & "s"
            Else
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Defesa: " & Item(itemNum).Data2
            End If
            ' stat bonuses
            If Item(itemNum).Add_Stat(Stats.Strength) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemNum).Add_Stat(Stats.Strength) & " FOR"
            End If
            If Item(itemNum).Add_Stat(Stats.Endurance) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemNum).Add_Stat(Stats.Endurance) & " CON"
            End If
            If Item(itemNum).Add_Stat(Stats.Intelligence) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemNum).Add_Stat(Stats.Intelligence) & " KI"
            End If
            If Item(itemNum).Add_Stat(Stats.Agility) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemNum).Add_Stat(Stats.Agility) & " DES"
            End If
            If Item(itemNum).Add_Stat(Stats.Willpower) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemNum).Add_Stat(Stats.Willpower) & " TEC"
            End If
        Case ITEM_TYPE_CONSUME
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Valor: " & Item(itemNum).Price & "z"
            If Item(itemNum).CastSpell > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Casts Spell"
            End If
            If Item(itemNum).AddHP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemNum).AddHP & " HP"
            End If
            If Item(itemNum).AddMP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemNum).AddMP & " MP"
            End If
            If Item(itemNum).AddEXP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemNum).AddEXP & " EXP"
            End If
        Case ITEM_TYPE_SPELL
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Valor: " & Item(itemNum).Price & "z"
    End Select
    
    ' go through and render all this shit
    Y = Y + 12
    For i = 1 To UBound(sInfo)
        Y = Y + 12
        RenderText Font_Default, sInfo(i), X + 141 - (EngineGetTextWidth(Font_Default, sInfo(i)) \ 2), Y, White
    Next
End Sub
Public Sub DrawPlayerSpellDesc()
Dim spellSlot As Long
    
    If Not GUIWindow(GUI_SPELLS).visible Then Exit Sub
    If DragSpell > 0 Then Exit Sub
    
    spellSlot = IsPlayerSpell(GlobalX, GlobalY)
    If spellSlot > 0 Then
        If PlayerSpells(spellSlot) > 0 Then
            DrawSpellDesc PlayerSpells(spellSlot), GlobalX + 32, GlobalY, spellSlot 'DrawSpellDesc PlayerSpells(spellSlot), GUIWindow(GUI_SPELLS).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_SPELLS).Y, spellSlot
        End If
    End If
    
    spellSlot = IsPlayerEvoluteSpell(GlobalX, GlobalY)
    If spellSlot > 0 Then
        DrawSpellDesc spellSlot, GlobalX + 32, GlobalY
    End If
End Sub

Public Sub DrawPlayerQuestDesc()
Dim spellSlot As Long
    
    If Not GUIWindow(GUI_QUESTS).visible Then Exit Sub
    
    spellSlot = IsPlayerQuest(GlobalX, GlobalY)
    If spellSlot > 0 Then
        DrawQuestDesc spellSlot, GUIWindow(GUI_SPELLS).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_SPELLS).Y
    End If
End Sub

Public Sub DrawSpellDesc(ByVal spellnum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal spellSlot As Long = 0)
Dim colour As Long, theName As String, sUse As String, sInfo() As String, i As Long, tmpWidth As Long, barWidth As Long
Dim Width As Long, Height As Long
    
    ' don't show desc when dragging
    If DragSpell > 0 Then Exit Sub
    
    ' get out
    If spellnum = 0 Then Exit Sub

    ' render the window
    Width = 190
    If Not Trim$(Spell(spellnum).Desc) = vbNullString Then
        Height = 210
    Else
        Height = 126
    End If
    'EngineRenderRectangle Tex_GUI(29), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(8), X, Y, 0, 0, Width, Height, Width, Height
    
    If spellSlot = 0 Then
        RenderTexture Tex_GUI(25), X, Y - 60, 0, 0, 190, 128, 128, 64
        RenderText Font_Default, "Requisitos", X + 95 - (getWidth(Font_Default, "Requisitos") / 2), Y - 80, Yellow
        If Spell(spellnum).Requisite > 0 Then
            RenderTexture Tex_Item(Item(Spell(spellnum).Requisite).Pic), X - 6, Y - 66, 0, 0, 32, 32, 32, 32
            colour = BrightRed
            If Player(MyIndex).Titulo > 0 Then
                If Item(Player(MyIndex).Titulo).LevelReq < Item(Spell(spellnum).Requisite).LevelReq Then
                    colour = BrightRed
                Else
                    colour = White
                End If
            Else
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(MyIndex, i) > 0 Then
                        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_TITULO Then
                            If Item(GetPlayerInvItemNum(MyIndex, i)).LevelReq < Item(Spell(spellnum).Requisite).LevelReq Then
                                colour = BrightRed
                            Else
                                colour = White
                            End If
                        End If
                    End If
                Next i
            End If
            RenderText Font_Default, "Cargo:" & Trim$(Item(Spell(spellnum).Requisite).Name), X + 16, Y - 52, colour
        Else
            RenderText Font_Default, "Habilidade especial!", X + 16, Y - 52, BrightCyan
        End If
        If Spell(spellnum).Item > 0 Then
            colour = BrightRed
            If HasItem(Spell(spellnum).Item) > Spell(spellnum).Price Then colour = White
            If Spell(spellnum).Requisite = 0 Then RenderTexture Tex_Item(Item(Spell(spellnum).Item).Pic), X - 6, Y - 40, 0, 0, 32, 32, 32, 32
            RenderText Font_Default, "Preço:" & Spell(spellnum).Price & " " & Trim$(Item(Spell(spellnum).Item).Name), X + 16, Y - 32, colour
        End If
    End If
    
    ' make sure it has a sprite
    If Spell(spellnum).Icon > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Spellicon(Spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 32, 32
        RenderTexture Tex_SpellIcon(Spell(spellnum).Icon), X + 16, Y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not Trim$(Spell(spellnum).Desc) = vbNullString Then
        RenderText Font_Default, WordWrap(Trim$(Spell(spellnum).Desc), Width - 10), X + 10, Y + 128, White
    End If
    
    ' render name
    colour = White
    theName = Trim$(Spell(spellnum).Name)
    RenderText Font_Default, theName, X + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), Y + 6, colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    i = 1
    ReDim Preserve sInfo(1 To i) As String
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP
            sInfo(i) = "Damage HP"
        Case SPELL_TYPE_DAMAGEMP
            sInfo(i) = "Damage SP"
        Case SPELL_TYPE_HEALHP
            sInfo(i) = "Heal HP"
        Case SPELL_TYPE_HEALMP
            sInfo(i) = "Heal SP"
        Case SPELL_TYPE_WARP
            sInfo(i) = "Warp"
    End Select
    
    ' more info
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
            ' damage
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Vital: " & Spell(spellnum).Vital
            
            ' mp cost
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Cost: " & Spell(spellnum).MPCost & " SP"
            
            ' cast time
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Cast Time: " & Spell(spellnum).CastTime & "s"
            
            ' cd time
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Cooldown: " & Spell(spellnum).CDTime & "s"
            
            ' aoe
            If Spell(spellnum).AoE > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "AoE: " & Spell(spellnum).AoE
            End If
            
            ' stun
            If Spell(spellnum).StunDuration > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Stun: " & Spell(spellnum).StunDuration & "s"
            End If
    End Select
    
    ' go through and render all this shit
    Y = Y + 12
    For i = 1 To UBound(sInfo)
        Y = Y + 12
        RenderText Font_Default, sInfo(i), X + 141 - (EngineGetTextWidth(Font_Default, sInfo(i)) \ 2), Y, White
    Next
End Sub

Public Sub DrawSkills()
Dim i As Long, X As Long, Y As Long, spellnum As Long, spellpic As Long
Dim Top As Long, Left As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 480
    Height = 384
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_SPELLS).x, GUIWindow(GUI_SPELLS).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(26), GUIWindow(GUI_SPELLS).X, GUIWindow(GUI_SPELLS).Y, 0, 0, Width, Height, Width, Height
    
    
    ' render skills
    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i)

        ' make sure not dragging it
        If DragSpell = i Then GoTo NextLoop
        
        ' actually render
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellpic = Spell(spellnum).Icon

            If spellpic > 0 And spellpic <= NumSpellIcons Then
                Top = GUIWindow(GUI_SPELLS).Y + SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                If SpellCD(i) > 0 Then
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                    RenderTexture Tex_SpellIcon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                Else
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                    RenderTexture Tex_SpellIcon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
        End If
NextLoop:
    Next
    
    RenderText Font_Default, ("Próximas skills"), GUIWindow(GUI_SPELLS).X + 48, GUIWindow(GUI_SPELLS).Y + 180, Yellow
    
    Dim Z As Long
    Z = 0
    For i = 1 To MAX_SPELLS
        spellnum = i

        ' actually render
        If ShowSpell(spellnum) Then
            spellpic = Spell(spellnum).Icon

            If spellpic > 0 And spellpic <= NumSpellIcons Then
                Top = GUIWindow(GUI_SPELLS).Y + 160 + SpellTop + ((SpellOffsetY + 32) * ((Z) \ SpellColumns))
                Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((Z) Mod SpellColumns)))
                'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                If Not CanBuySpell(spellnum) Then
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                    RenderTexture Tex_SpellIcon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                Else
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                    RenderTexture Tex_SpellIcon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
            Z = Z + 1
        End If
    Next
End Sub

Function CanBuySpell(ByVal spellnum As Long) As Boolean
    If GetPlayerLevel(MyIndex) >= Spell(spellnum).LevelReq Then
        
        Dim itemNum As Long, ItemValue As Long
        
        itemNum = Spell(spellnum).Item
        ItemValue = Spell(spellnum).Price
        
        If Spell(spellnum).Requisite > 0 Then
            If Item(Spell(spellnum).Requisite).Type <> ITEM_TYPE_TITULO Then
                If Not HasItem(Spell(spellnum).Requisite) > 0 Then
                    Exit Function
                End If
            Else
                If Player(MyIndex).Titulo > 0 Then
                    If Item(Player(MyIndex).Titulo).LevelReq < Item(Spell(spellnum).Requisite).LevelReq Then
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            End If
        End If
        
        If itemNum > 0 Then
            If HasItem(itemNum) < ItemValue Then
                Exit Function
            End If
        End If
        
        CanBuySpell = True
    End If
End Function

Function HasItem(ByVal itemNum As Long) As Long
    Dim i As Long
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) = itemNum Then
            HasItem = GetPlayerInvItemValue(MyIndex, i)
            Exit Function
        End If
    Next i
End Function

Function ShowSpell(ByVal spellnum As Long) As Boolean
    If Spell(spellnum).AccessReq > 0 Then Exit Function
    If HasSpell(spellnum) Then Exit Function
    If Spell(spellnum).Upgrade > 0 Then
        'Primeiras skills
        If Not HasSpell(Spell(spellnum).Upgrade) Then
            If Not HasAntecessor(spellnum) Then
                ShowSpell = True
                Exit Function
            End If
        End If
    Else
        If Not HasAntecessor(spellnum) Then
            ShowSpell = True
            Exit Function
        End If
    End If
            
    Dim i As Long
    For i = 1 To MAX_PLAYER_SPELLS
        If PlayerSpells(i) > 0 Then
            If Spell(PlayerSpells(i)).Upgrade = spellnum Then
                ShowSpell = True
                Exit For
            End If
        End If
    Next i
End Function

Function HasAntecessor(ByVal spellnum As Long) As Boolean
    Dim i As Long
    For i = 1 To MAX_SPELLS
        If Trim$(Spell(i).Name) = vbNullString Then
            Exit For
        Else
            If Spell(i).Upgrade = spellnum Then
                HasAntecessor = True
                Exit For
            End If
        End If
    Next i
End Function

Function HasSpell(ByVal spellnum As Long) As Boolean
    Dim i As Long, SpellRealName As String
    Dim SpellLength As Byte

    'Evolutions
    SpellRealName = Trim$(Spell(spellnum).Name)
    SpellLength = Len(SpellRealName)

    For i = 1 To MAX_PLAYER_SPELLS

        If PlayerSpells(i) > 0 Then
            If PlayerSpells(i) = spellnum Then
                HasSpell = True
                Exit Function
            End If
        
            If Mid(Trim$(Spell(PlayerSpells(i)).Name), 1, SpellLength) = SpellRealName Then
                HasSpell = True
                Exit Function
            End If
        End If

    Next

End Function

Public Sub DrawQuestDesc(ByVal QuestNum As Long, ByVal X As Long, ByVal Y As Long)
Dim colour As Long, theName As String, sUse As String, sInfo() As String, i As Long, tmpWidth As Long, barWidth As Long
Dim Width As Long, Height As Long
    
    ' get out
    If QuestNum = 0 Then Exit Sub

    ' render the window
    Width = 190
    If Not Trim$(Quest(QuestNum).Desc) = vbNullString Then
        Height = 210
    Else
        Height = 126
    End If
    'EngineRenderRectangle Tex_GUI(29), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(8), X, Y, 0, 0, Width, Height, Width, Height
    
    ' make sure it has a sprite
    If Quest(QuestNum).Icon > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Spellicon(Spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 32, 32
        RenderTexture Tex_Item(Quest(QuestNum).Icon), X + 16, Y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not Trim$(Quest(QuestNum).Desc) = vbNullString Then
        RenderText Font_Default, WordWrap(Trim$(Quest(QuestNum).Desc), Width - 10), X + 10, Y + 128, White
    End If
    
    ' render name
    colour = White
    theName = Trim$(Quest(QuestNum).Name)
    RenderText Font_Default, theName, X + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), Y + 6, colour
    
    colour = Yellow
    theName = "Duplo clique para mais informações"
    RenderText Font_Default, WordWrap(theName, (Width - 10) / 2), X + 105, Y + 40, colour
End Sub

Public Sub DrawQuests()
Dim i As Long, X As Long, Y As Long, questpic As Long
Dim Top As Long, Left As Long
Dim Width As Long, Height As Long
Dim ActualQuest As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_SPELLS).x, GUIWindow(GUI_SPELLS).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_SPELLS).X, GUIWindow(GUI_SPELLS).Y, 0, 0, Width, Height, Width, Height
    
    ' render skills
    For i = 1 To MAX_QUESTS
        ' make sure are doing it
        If Player(MyIndex).QuestState(i).State <> 1 Then GoTo NextLoop
        
        ' actually render
            questpic = Quest(i).Icon
            
            ActualQuest = ActualQuest + 1

            If questpic > 0 And questpic <= numitems Then
                Top = GUIWindow(GUI_SPELLS).Y + SpellTop + ((SpellOffsetY + 32) * ((ActualQuest - 1) \ SpellColumns))
                Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((ActualQuest - 1) Mod SpellColumns)))
                RenderTexture Tex_Item(questpic), Left, Top, 0, 0, 32, 32, 32, 32
            End If
NextLoop:
    Next
End Sub

Public Sub DrawEquipment()
Dim X As Long, Y As Long, i As Long
Dim itemNum As Long, ItemPic As DX8TextureRec

    For i = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(MyIndex, i)

        ' get the item sprite
        If itemNum > 0 Then
            ItemPic = Tex_Item(Item(itemNum).Pic)
        Else
            ' no item equiped - use blank image
            ItemPic = Tex_GUI(8 + i)
        End If
        
        Y = GUIWindow(GUI_CHARACTER).Y + EqTop
        X = GUIWindow(GUI_CHARACTER).X + EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))

        'EngineRenderRectangle itempic, x, y, 0, 0, 32, 32, 32, 32, 32, 32
        RenderTexture ItemPic, X, Y, 0, 0, 32, 32, 32, 32
    Next
End Sub

Public Sub DrawCharacter()
Dim X As Long, Y As Long, i As Long, dX As Long, dY As Long, tmpString As String, buttonnum As Long
Dim Width As Long, Height As Long, sWidth As Long, sHeight As Long

    SetTexture Tex_Bars
    ' dynamic bar calculations
    sWidth = 96
    sHeight = 8
    
    X = GUIWindow(GUI_CHARACTER).X
    Y = GUIWindow(GUI_CHARACTER).Y
    
    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(5), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(6), X, Y, 0, 0, Width, Height, Width, Height
    
    ' render name
    tmpString = Trim$(GetPlayerName(MyIndex)) & " - Level " & GetPlayerLevel(MyIndex)
    RenderText Font_Default, tmpString, X + 7 + (187 / 2) - (EngineGetTextWidth(Font_Default, tmpString) / 2), Y + 9, White
    
    ' render stats
    dX = X + 20
    dY = Y + 48
    RenderText Font_Default, "Força", dX, dY, White
    RenderText Font_Default, "LEVEL " & GetPlayerStat(MyIndex, Strength), dX + 155 - EngineGetTextWidth(Font_Default, "LEVEL " & GetPlayerStat(MyIndex, Strength)), dY, Yellow
    dY = dY + 30
    RenderText Font_Default, "Constituição", dX, dY, White
    RenderText Font_Default, "LEVEL " & GetPlayerStat(MyIndex, Endurance), dX + 155 - EngineGetTextWidth(Font_Default, "LEVEL " & GetPlayerStat(MyIndex, Endurance)), dY, Yellow
    dY = dY + 30
    RenderText Font_Default, "KI", dX, dY, White
    RenderText Font_Default, "LEVEL " & GetPlayerStat(MyIndex, Intelligence), dX + 155 - EngineGetTextWidth(Font_Default, "LEVEL " & GetPlayerStat(MyIndex, Intelligence)), dY, Yellow
    dY = dY + 30
    RenderText Font_Default, "Destreza", dX, dY, White
    RenderText Font_Default, "LEVEL " & GetPlayerStat(MyIndex, Agility), dX + 155 - EngineGetTextWidth(Font_Default, "LEVEL " & GetPlayerStat(MyIndex, Agility)), dY, Yellow
    dY = dY + 30
    RenderText Font_Default, "Técnica", dX, dY, White
    RenderText Font_Default, "LEVEL " & GetPlayerStat(MyIndex, Willpower), dX + 155 - EngineGetTextWidth(Font_Default, "LEVEL " & GetPlayerStat(MyIndex, Willpower)), dY, Yellow
    dY = Y + 28
    RenderText Font_Default, "Pontos: " & GetPlayerPOINTS(MyIndex) & "/" & Player(MyIndex).Level * 3, dX, dY, Yellow
    
    'dY = Y + 64
    'For i = 1 To 5
    '    Dim sRECT As RECT
    '    With sRECT
    '        .Top = 10 ' HP bar background
    '        .Left = 0
    '        .Right = .Left + sWidth
    '        .Bottom = .Top + sHeight
    '    End With
    '    RenderTexture Tex_Bars, dX, dY + (30 * (i - 1)), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
    '    With sRECT
    '        .Top = 10 ' HP bar background
    '        .Left = 0
    '        .Right = .Left + (sWidth * ((GetPlayerStatPoints(MyIndex, i) - GetPlayerStatPrevLevel(MyIndex, i)) / (GetPlayerStatNextLevel(MyIndex, i) - GetPlayerStatPrevLevel(MyIndex, i))))
    '        .Bottom = .Top + sHeight
    '    End With
    '    RenderTexture Tex_Bars, dX, dY + (30 * (i - 1)), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(100, 255, 100, 255)
    'Next i
    
    ' draw the face
    'If GetPlayerSprite(MyIndex) > 0 And GetPlayerSprite(MyIndex) <= NumFaces Then
        'EngineRenderRectangle Tex_Face(GetPlayerSprite(MyIndex)), x + 49, y + 38, 0, 0, 96, 96, 96, 96, 96, 96
        'RenderTexture Tex_Character(GetPlayerSprite(MyIndex)), X + 49, Y + 38, 0, 0, 96, 96, 32, 34
    'End If
    
    If GetPlayerPOINTS(MyIndex) > 0 Then
        ' draw the buttons
        For buttonnum = 16 To 20
            X = GUIWindow(GUI_CHARACTER).X + Buttons(buttonnum).X
            Y = GUIWindow(GUI_CHARACTER).Y + Buttons(buttonnum).Y
            Width = Buttons(buttonnum).Width
            Height = Buttons(buttonnum).Height
            ' render accept button
            If Buttons(buttonnum).State = 2 Then
                ' we're clicked boyo
                Width = Buttons(buttonnum).Width
                Height = Buttons(buttonnum).Height
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(buttonnum).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(buttonnum).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = buttonnum Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = buttonnum
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = buttonnum Then lastButtonSound = 0
            End If
        Next
    End If
    
    ' draw the equipment
    DrawEquipment
End Sub

Public Sub DrawOptions()
Dim i As Long, X As Long, Y As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(24), GUIWindow(GUI_OPTIONS).x, GUIWindow(GUI_OPTIONS).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_OPTIONS).X, GUIWindow(GUI_OPTIONS).Y, 0, 0, Width, Height, Width, Height
    
    RenderText Font_Default, "Music: ", GUIWindow(GUI_OPTIONS).X + 20, GUIWindow(GUI_OPTIONS).Y + 17, White
    RenderText Font_Default, "Sound: ", GUIWindow(GUI_OPTIONS).X + 20, GUIWindow(GUI_OPTIONS).Y + 40, White
    RenderText Font_Default, "Debug: ", GUIWindow(GUI_OPTIONS).X + 20, GUIWindow(GUI_OPTIONS).Y + 63, White
    RenderText Font_Default, "FPS Cap: ", GUIWindow(GUI_OPTIONS).X + 20, GUIWindow(GUI_OPTIONS).Y + 103, White
    RenderText Font_Default, "Volume: ", GUIWindow(GUI_OPTIONS).X + 20, GUIWindow(GUI_OPTIONS).Y + 122, White
    
    Select Case Options.FPS
        Case 15
            RenderText Font_Default, "64", GUIWindow(GUI_OPTIONS).X + 120, GUIWindow(GUI_OPTIONS).Y + 103, Yellow
        Case 20
            RenderText Font_Default, "32", GUIWindow(GUI_OPTIONS).X + 120, GUIWindow(GUI_OPTIONS).Y + 103, Yellow
        Case Else
            RenderText Font_Default, "XX", GUIWindow(GUI_OPTIONS).X + 120, GUIWindow(GUI_OPTIONS).Y + 103, BrightRed
    End Select
    
    RenderText Font_Default, Options.volume, GUIWindow(GUI_OPTIONS).X + 120, GUIWindow(GUI_OPTIONS).Y + 122, Yellow
    ' draw buttons
    For i = 26 To 31
        ' set co-ordinate
        X = GUIWindow(GUI_OPTIONS).X + Buttons(i).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(i).Y
        Width = Buttons(i).Width
        Height = Buttons(i).Height
        ' check for state
        If Buttons(i).State = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = i Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = i
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
    For i = 42 To 45
    ' set co-ordinate
        X = GUIWindow(GUI_OPTIONS).X + Buttons(i).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(i).Y
        Width = Buttons(i).Width
        Height = Buttons(i).Height
        ' check for state
        If Buttons(i).State = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = i Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = i
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawParty()
Dim i As Long, X As Long, Y As Long, Width As Long, playerNum As Long, theName As String
Dim Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_PARTY).x, GUIWindow(GUI_PARTY).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(7), GUIWindow(GUI_PARTY).X, GUIWindow(GUI_PARTY).Y, 0, 0, Width, Height, Width, Height
    
    ' draw the bars
    If Party.Leader > 0 Then ' make sure we're in a party
        ' draw leader
        playerNum = Party.Leader
        ' name
        theName = Trim$(GetPlayerName(playerNum))
        ' draw name
        Y = GUIWindow(GUI_PARTY).Y + 12
        X = GUIWindow(GUI_PARTY).X + 7 + 90 - (EngineGetTextWidth(Font_Default, theName) / 2)
        RenderText Font_Default, theName, X, Y, White
        ' draw hp
        Y = GUIWindow(GUI_PARTY).Y + 29
        X = GUIWindow(GUI_PARTY).X + 6
        ' make sure we actually have the data before rendering
        If GetPlayerVital(playerNum, Vitals.HP) > 0 And GetPlayerMaxVital(playerNum, Vitals.HP) > 0 Then
            Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
        End If
        'EngineRenderRectangle Tex_GUI(13), x, y, 0, 0, width, 9, width, 9, width, 9
        RenderTexture Tex_GUI(13), X, Y, 0, 0, Width, 9, Width, 9
        ' draw mp
        Y = GUIWindow(GUI_PARTY).Y + 38
        ' make sure we actually have the data before rendering
        If GetPlayerVital(playerNum, Vitals.MP) > 0 And GetPlayerMaxVital(playerNum, Vitals.MP) > 0 Then
            Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
        End If
        'EngineRenderRectangle Tex_GUI(14), x, y, 0, 0, width, 9, width, 9, width, 9
        RenderTexture Tex_GUI(14), X, Y, 0, 0, Width, 9, Width, 9
        
        ' draw members
        For i = 1 To MAX_PARTY_MEMBERS
            If Party.Member(i) > 0 Then
                If Party.Member(i) <> Party.Leader Then
                    ' cache the index
                    playerNum = Party.Member(i)
                    ' name
                    theName = Trim$(GetPlayerName(playerNum))
                    ' draw name
                    Y = GUIWindow(GUI_PARTY).Y + 12 + ((i - 1) * 49)
                    X = GUIWindow(GUI_PARTY).X + 7 + 90 - (EngineGetTextWidth(Font_Default, theName) / 2)
                    RenderText Font_Default, theName, X, Y, White
                    ' draw hp
                    Y = GUIWindow(GUI_PARTY).Y + 29 + ((i - 1) * 49)
                    X = GUIWindow(GUI_PARTY).X + 6
                    ' make sure we actually have the data before rendering
                    If GetPlayerVital(playerNum, Vitals.HP) > 0 And GetPlayerMaxVital(playerNum, Vitals.HP) > 0 Then
                        Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
                    End If
                    'EngineRenderRectangle Tex_GUI(13), x, y, 0, 0, width, 9, width, 9, width, 9
                    RenderTexture Tex_GUI(13), X, Y, 0, 0, Width, 9, Width, 9
                    ' draw mp
                    Y = GUIWindow(GUI_PARTY).Y + 38 + ((i - 1) * 49)
                    ' make sure we actually have the data before rendering
                    If GetPlayerVital(playerNum, Vitals.MP) > 0 And GetPlayerMaxVital(playerNum, Vitals.MP) > 0 Then
                        Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
                    End If
                    'EngineRenderRectangle Tex_GUI(14), x, y, 0, 0, width, 9, width, 9, width, 9
                    RenderTexture Tex_GUI(14), X, Y, 0, 0, Width, 9, Width, 9
                End If
            End If
        Next
    End If
    
    ' draw buttons
    For i = 24 To 25
        ' set co-ordinate
        X = GUIWindow(GUI_PARTY).X + Buttons(i).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(i).Y
        Width = Buttons(i).Width
        Height = Buttons(i).Height
        ' check for state
        If Buttons(i).State = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = i Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = i
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
End Sub
Public Sub DrawCurrency()
Dim X As Long, Y As Long, buttonnum As Long
Dim Width As Long, Height As Long

    X = GUIWindow(GUI_CURRENCY).X
    Y = GUIWindow(GUI_CURRENCY).Y
    ' render chatbox
    Width = GUIWindow(GUI_CURRENCY).Width
    Height = GUIWindow(GUI_CURRENCY).Height
    RenderTexture Tex_GUI(21), X, Y, 0, 0, Width, Height, Width, Height
    Width = EngineGetTextWidth(Font_Default, CurrencyText)
    RenderText Font_Default, CurrencyText, X + 87 + (123 - (Width / 2)), Y + 40, White
    RenderText Font_Default, sDialogue & chatShowLine, X + 90, Y + 65, White
    
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If CurrencyAcceptState = 2 Then
        ' clicked
        RenderText Font_Default, "[Accept]", X, Y, Grey
    Else
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' hover
            RenderText Font_Default, "[Accept]", X, Y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 1 Then
                PlaySound Sound_ButtonHover, -1, -1
                lastNpcChatsound = 1
            End If
        Else
            ' normal
            RenderText Font_Default, "[Accept]", X, Y, Green
            ' reset sound if needed
            If lastNpcChatsound = 1 Then lastNpcChatsound = 0
        End If
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If CurrencyCloseState = 2 Then
        ' clicked
        RenderText Font_Default, "[Close]", X, Y, Grey
    Else
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' hover
            RenderText Font_Default, "[Close]", X, Y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 2 Then
                PlaySound Sound_ButtonHover, -1, -1
                lastNpcChatsound = 2
            End If
        Else
            ' normal
            RenderText Font_Default, "[Close]", X, Y, Yellow
            ' reset sound if needed
            If lastNpcChatsound = 2 Then lastNpcChatsound = 0
        End If
    End If
End Sub
Public Sub DrawDialogue()
Dim i As Long, X As Long, Y As Long, Sprite As Long, Width As Long
Dim Height As Long

    ' draw background
    X = GUIWindow(GUI_DIALOGUE).X
    Y = GUIWindow(GUI_DIALOGUE).Y
    
    ' render chatbox
    Width = GUIWindow(GUI_DIALOGUE).Width
    Height = GUIWindow(GUI_DIALOGUE).Height
    RenderTexture Tex_GUI(19), X, Y, 0, 0, Width, Height, Width, Height
    
    ' Draw the text
    RenderText Font_Default, WordWrap(Dialogue_TitleCaption, 392), X + 10, Y + 10, White
    RenderText Font_Default, WordWrap(Dialogue_TextCaption, 392), X + 10, Y + 25, White
    
    If Dialogue_ButtonVisible(1) Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 90
            If Dialogue_ButtonState(1) = 2 Then
                ' clicked
                RenderText Font_Default, "[Accept]", X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Default, "[Accept]", X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = 1 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 1
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[Accept]", X, Y, Green
                    ' reset sound if needed
                    If lastNpcChatsound = 1 Then lastNpcChatsound = 0
                End If
            End If
    End If
    If Dialogue_ButtonVisible(2) Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 105
            If Dialogue_ButtonState(2) = 2 Then
                ' clicked
                RenderText Font_Default, "[Okay]", X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Default, "[Okay]", X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = 2 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 2
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[Okay]", X, Y, BrightRed
                    ' reset sound if needed
                    If lastNpcChatsound = 2 Then lastNpcChatsound = 0
                End If
            End If
    End If
    If Dialogue_ButtonVisible(3) Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        X = GUIWindow(GUI_DIALOGUE).X + 10 + (155 - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 120
        If Dialogue_ButtonState(3) = 2 Then
            ' clicked
            RenderText Font_Default, "[Close]", X, Y, Grey
        Else
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                ' hover
                RenderText Font_Default, "[Close]", X, Y, Cyan
                ' play sound if needed
                If Not lastNpcChatsound = 3 Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastNpcChatsound = 3
                End If
            Else
                ' normal
                RenderText Font_Default, "[Close]", X, Y, Yellow
                ' reset sound if needed
                If lastNpcChatsound = 3 Then lastNpcChatsound = 0
            End If
        End If
    End If
End Sub

Public Sub DrawShop()
Dim i As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = GUIWindow(GUI_SHOP).Width
    Height = GUIWindow(GUI_SHOP).Height
    'EngineRenderRectangle Tex_GUI(23), GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(20), GUIWindow(GUI_SHOP).X, GUIWindow(GUI_SHOP).Y, 0, 0, Width, Height, Width, Height
     
    RenderText Font_Default, "Duplo clique para efetuar a compra", GUIWindow(GUI_SHOP).X + 22, GUIWindow(GUI_SHOP).Y + 8, White
    If GUIWindow(GUI_INVENTORY).visible = True Then RenderText Font_Default, "Duplo clique para efetuar a venda", GUIWindow(GUI_INVENTORY).X, GUIWindow(GUI_INVENTORY).Y - 16, White
     
    ' render the shop items
    For i = 1 To MAX_TRADES
        itemNum = Shop(InShop).TradeItem(i).Item
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            If ItemPic > 0 And ItemPic <= numitems Then
                
                Top = GUIWindow(GUI_SHOP).Y + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                Left = GUIWindow(GUI_SHOP).X + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    Y = GUIWindow(GUI_SHOP).Y + Top + 22
                    X = GUIWindow(GUI_SHOP).X + Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), Left + 4, Top + 4, colour
                End If
            End If
        End If
    Next
    
    ' draw buttons
    For i = 23 To 23
        ' set co-ordinate
        X = GUIWindow(GUI_SHOP).X + Buttons(i).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(i).Y
        Width = Buttons(i).Width
        Height = Buttons(i).Height
        ' check for state
        If Buttons(i).State = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = i Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = i
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
    
    ' draw item descriptions
    DrawShopItemDesc
End Sub

Public Sub DrawMenu()
Dim i As Long, X As Long, Y As Long
Dim Width As Long, Height As Long

    ' draw background
    X = GUIWindow(GUI_MENU).X
    Y = GUIWindow(GUI_MENU).Y
    Width = GUIWindow(GUI_MENU).Width
    Height = GUIWindow(GUI_MENU).Height
    'EngineRenderRectangle Tex_GUI(3), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(3), X, Y, 0, 0, Width, Height, Width, Height
    
    'RenderText Font_Default, MoedasZ & " Z", GUIWindow(GUI_MENU).X + 30, GUIWindow(GUI_MENU).Y + 33, Grey, , True
    
    ' draw buttons
    For i = 1 To 6
        If Buttons(i).visible Then
            ' set co-ordinate
            X = GUIWindow(GUI_MENU).X + Buttons(i).X
            Y = GUIWindow(GUI_MENU).Y + Buttons(i).Y
            Width = Buttons(i).Width
            Height = Buttons(i).Height
            ' check for state
            If Buttons(i).State = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                If i <> 6 Then
                    RenderTexture Tex_GUI(25), X - 46, Y - 40, 0, 0, 128, 64, 128, 64, D3DColorRGBA(255, 255, 255, 100)
                Else
                    RenderTexture Tex_GUI(25), X - 23, Y - 40, 0, 0, 128, 64, 64, 64, D3DColorRGBA(255, 255, 255, 100)
                    'RenderTexture Tex_GUI(25), X - 46 + 76, Y - 40, 0, 0, 6, 64, 122, 64, D3DColorRGBA(255, 255, 255, 100)
                End If
                RenderText Font_Default, MenuButtonName(i), X - (getWidth(Font_Default, MenuButtonName(i)) / 2) + 18, Y - 35, Grey, , True
                ' play sound if needed
                If Not lastButtonSound = i Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = i
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = i Then lastButtonSound = 0
            End If
        End If
    Next
End Sub


Public Sub DrawBank()
Dim i As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long, Width As Long
Dim Height As Long

    Width = GUIWindow(GUI_BANK).Width
    Height = GUIWindow(GUI_BANK).Height
    
    RenderTexture Tex_GUI(26), GUIWindow(GUI_BANK).X, GUIWindow(GUI_BANK).Y, 0, 0, Width, Height, Width, Height
    
    ' render the bank items' are you serous? that is it??? maybe... one sec :D :Polol
        For i = 1 To MAX_BANK
            itemNum = GetBankItemNum(i)
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                        
                     Top = GUIWindow(GUI_BANK).Y + BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                     Left = GUIWindow(GUI_BANK).X + BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))

                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                       
                    ' If the bank item is in a stack, draw the amount...
                    If GetBankItemValue(i) > 1 Then
                        Y = Top + 22
                        X = Left - 4
                        Amount = CStr(GetBankItemValue(i))
                            
                        ' Draw the currency
                        If CLng(Amount) < 1000000 Then
                            colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            colour = BrightGreen
                        End If
                    
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour
                    End If
                End If
            End If
    Next
    
            
             DrawBankItemDesc
                            
                        
End Sub
Public Sub DrawBankItemDesc()
Dim bankNum As Long
    If Not GUIWindow(GUI_BANK).visible Then Exit Sub
        
        bankNum = IsBankItem(GlobalX, GlobalY)
     
        
    If bankNum > 0 Then
        If bankNum > 0 Then
            If GetBankItemNum(bankNum) > 0 Then
                DrawItemDesc GetBankItemNum(bankNum), GUIWindow(GUI_BANK).X + 480, GUIWindow(GUI_BANK).Y
           End If
        End If
    End If
            
End Sub

Public Sub DrawTrade()
Dim i As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long, Width As Long
Dim Height As Long

    Width = GUIWindow(GUI_TRADE).Width
    Height = GUIWindow(GUI_TRADE).Width
    RenderTexture Tex_GUI(18), GUIWindow(GUI_TRADE).X, GUIWindow(GUI_TRADE).Y, 0, 0, Width, Height, Width, Height
        For i = 1 To MAX_INV
            ' render your offer
            itemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                    Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).X + 29 + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    ' If item is a stack - draw the amount you have
                    If TradeYourOffer(i).Value > 1 Then
                        Y = Top + 21
                        X = Left - 4
                            
                        Amount = CStr(TradeYourOffer(i).Value)
                            
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            colour = BrightGreen
                        End If
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour
                    End If
                End If
            End If
            
            ' draw their offer
            itemNum = TradeTheirOffer(i).Num
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                
                    Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop - 2 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).X + 257 + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    ' If item is a stack - draw the amount you have
                    If TradeTheirOffer(i).Value > 1 Then
                        Y = Top + 21
                        X = Left - 4
                                
                        Amount = CStr(TradeTheirOffer(i).Value)
                                
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            colour = BrightGreen
                        End If
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour
                    End If
                End If
            End If
        Next
        ' draw buttons
    For i = 40 To 41
        ' set co-ordinate
        X = Buttons(i).X
        Y = Buttons(i).Y
        Width = Buttons(i).Width
        Height = Buttons(i).Height
        ' check for state
        If Buttons(i).State = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = i Then
                PlaySound Sound_ButtonHover, -1, -1
                lastButtonSound = i
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
    RenderText Font_Default, "Você tem: " & YourWorth, GUIWindow(GUI_TRADE).X + 21, GUIWindow(GUI_TRADE).Y + 299, White
    RenderText Font_Default, "Ele tem: " & TheirWorth, GUIWindow(GUI_TRADE).X + 250, GUIWindow(GUI_TRADE).Y + 299, White
    RenderText Font_Default, TradeStatus, (GUIWindow(GUI_TRADE).Width / 2) - (EngineGetTextWidth(Font_Default, TradeStatus) / 2), GUIWindow(GUI_TRADE).Y + 317, Yellow
    DrawTradeItemDesc
End Sub

Public Sub DrawTradeItemDesc()
Dim tradeNum As Long
    If Not GUIWindow(GUI_TRADE).visible Then Exit Sub
        
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum > 0 Then
        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).Num) > 0 Then
            DrawItemDesc GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).Num), GUIWindow(GUI_TRADE).X + 480 + 10, GUIWindow(GUI_TRADE).Y
        End If
    End If
End Sub

Public Sub DrawGUIBars()
Dim tmpWidth As Long, barWidth As Long, X As Long, Y As Long, dX As Long, dY As Long, sString As String
Dim Width As Long, Height As Long

    ' backwindow + empty bars
    X = GUIWindow(GUI_BARS).X
    Y = GUIWindow(GUI_BARS).Y
    Width = GUIWindow(GUI_BARS).Width
    Height = GUIWindow(GUI_BARS).Height
    'EngineRenderRectangle Tex_GUI(4), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(4), X, Y, 0, 0, Width, Height, Width, Height
    
    ' hardcoded for POT textures
    barWidth = 192
    
    ' health bar
    BarWidth_GuiHP_Max = ((GetPlayerVital(MyIndex, Vitals.HP) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / barWidth)) * barWidth
    RenderTexture Tex_GUI(13), X + 89, Y + 24, 0, 0, BarWidth_GuiHP, Tex_GUI(13).Height, BarWidth_GuiHP, Tex_GUI(13).Height
    ' render health
    sString = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
    dX = X + 89 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = Y + 24
    RenderText Font_Default, sString, dX, dY, White
    
    barWidth = 183
    ' spirit bar
    BarWidth_GuiSP_Max = ((GetPlayerVital(MyIndex, Vitals.MP) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / barWidth)) * barWidth
    RenderTexture Tex_GUI(14), X + 89, Y + 42, 0, 0, BarWidth_GuiSP, Tex_GUI(14).Height, BarWidth_GuiSP, Tex_GUI(14).Height
    ' render spirit
    sString = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
    dX = X + 89 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = Y + 42
    RenderText Font_Default, sString, dX, dY, White
    
    barWidth = 175
    ' exp bar
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        BarWidth_GuiEXP_Max = ((GetPlayerExp(MyIndex) / barWidth) / (TNL / barWidth)) * barWidth
    Else
        BarWidth_GuiEXP_Max = barWidth
    End If
    RenderTexture Tex_GUI(15), X + 89, Y + 60, 0, 0, BarWidth_GuiEXP, Tex_GUI(15).Height, BarWidth_GuiEXP, Tex_GUI(15).Height
    ' render exp
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        sString = GetPlayerExp(MyIndex) & "/" & TNL
    Else
        sString = "Max Level"
    End If
    dX = X + 89 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = Y + 60
    RenderText Font_Default, sString, dX, dY, White
End Sub
Public Sub DrawEventChat()
Dim i As Long, X As Long, Y As Long, Sprite As Long, Width As Long
Dim Height As Long

    ' draw background
    X = GUIWindow(GUI_EVENTCHAT).X
    Y = GUIWindow(GUI_EVENTCHAT).Y
    
    ' render chatbox
    Width = GUIWindow(GUI_EVENTCHAT).Width
    Height = GUIWindow(GUI_EVENTCHAT).Height
    RenderTexture Tex_GUI(19), X, Y, 0, 0, Width, Height, Width, Height
    
    Select Case CurrentEvent.Type
        Case Evt_Menu
            ' Draw replies
            RenderText Font_Default, WordWrap(Trim$(CurrentEvent.Text(1)), GUIWindow(GUI_EVENTCHAT).Width - 10), X + 10, Y + 10, White
            For i = 1 To UBound(CurrentEvent.Text) - 1
                If Len(Trim$(CurrentEvent.Text(i + 1))) > 0 Then
                    Width = EngineGetTextWidth(Font_Default, "[" & Trim$(CurrentEvent.Text(i + 1)) & "]")
                    X = GUIWindow(GUI_CHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
                    Y = GUIWindow(GUI_CHAT).Y + 115 - ((i - 1) * 15)
                    If chatOptState(i) = 2 Then
                        ' clicked
                        RenderText Font_Default, "[" & Trim$(CurrentEvent.Text(i + 1)) & "]", X, Y, Grey
                    Else
                        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                            ' hover
                            RenderText Font_Default, "[" & Trim$(CurrentEvent.Text(i + 1)) & "]", X, Y, Yellow
                            ' play sound if needed
                            If Not lastNpcChatsound = i Then
                                PlaySound Sound_ButtonHover, -1, -1
                                lastNpcChatsound = i
                            End If
                        Else
                            ' normal
                            RenderText Font_Default, "[" & Trim$(CurrentEvent.Text(i + 1)) & "]", X, Y, BrightBlue
                            ' reset sound if needed
                            If lastNpcChatsound = i Then lastNpcChatsound = 0
                        End If
                    End If
                End If
            Next
        Case Evt_Message
            RenderText Font_Default, WordWrap(Trim$(CurrentEvent.Text(1)), GUIWindow(GUI_EVENTCHAT).Width - 52), X + 52, Y + 10, White
            If CurrentEvent.Data(1) > 0 Then
                RenderTexture Tex_Character(CurrentEvent.Data(1)), X + 10, Y + 10, 32, 0, 32, 32, 32, 32
            Else
                RenderTexture Tex_Character(Player(MyIndex).Sprite), X + 10, Y + 10, 0, 0, 32, 64, 32, 64
                RenderTexture Tex_Hair(0).TexHair(Player(MyIndex).Hair), X + 10, Y + 10, 0, 0, 32, 64, 32, 64
                ' check for paperdolling
                For i = 1 To UBound(PaperdollOrder)
                    If GetPlayerEquipment(MyIndex, PaperdollOrder(i)) > 0 Then
                        If Item(GetPlayerEquipment(MyIndex, PaperdollOrder(i))).Paperdoll > 0 Then
                            Call DrawPaperdoll(X + 10, Y + 10, Item(GetPlayerEquipment(MyIndex, PaperdollOrder(i))).Paperdoll, 0, 0, True)
                        End If
                    End If
                Next
            End If
            Width = EngineGetTextWidth(Font_Default, "[Continue]")
            X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
            Y = GUIWindow(GUI_EVENTCHAT).Y + 100
            If chatContinueState = 2 Then
                ' clicked
                RenderText Font_Default, "[Continue]", X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Default, "[Continue]", X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = i Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = i
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[Continue]", X, Y, BrightBlue
                    ' reset sound if needed
                    If lastNpcChatsound = i Then lastNpcChatsound = 0
                End If
            End If
    End Select
End Sub
Public Sub DrawMapName()
Dim X As Long, Y As Long, color As Long

    X = GUIWindow(GUI_HOTBAR).X
    
    Y = GUIWindow(GUI_HOTBAR).Y + 40


    RenderText Font_Default, Map.Name, X, Y, White
End Sub

Public Sub DrawMapGrid(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT

    rec.Top = 24
    rec.Left = 0
    rec.Right = rec.Left + 32
    rec.Bottom = rec.Top + 32
    RenderTexture Tex_Direction, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top
End Sub

Public Sub DrawProjectile()
Dim Angle As Long, X As Long, Y As Long, i As Long
    If LastProjectile > 0 Then
        
        ' ****** Create Particle ******
        For i = 1 To LastProjectile
            With ProjectileList(i)
                If .Graphic Then
                
                    ' ****** Update Position ******
                    Angle = DegreeToRadian * Engine_GetAngle(.X, .Y, .tx, .ty)
                    .X = .X + (Sin(Angle) * ElapsedTime * 0.3)
                    .Y = .Y - (Cos(Angle) * ElapsedTime * 0.3)
                    X = .X - (Tex_Projectile(.Graphic).Width / 2)
                    Y = .Y - (Tex_Projectile(.Graphic).Height / 2)
                    
                    ' ****** Update Rotation ******
                    If .RotateSpeed > 0 Then
                        .Rotate = .Rotate + (.RotateSpeed * ElapsedTime * 0.01)
                        Do While .Rotate > 360
                            .Rotate = .Rotate - 360
                        Loop
                    End If
                    
                    ' ****** Render Projectile ******
                    If .Rotate = 0 Then
                        Call RenderTexture(Tex_Projectile(.Graphic), ConvertMapX(X), ConvertMapY(Y), 0, 0, Tex_Projectile(.Graphic).Width, Tex_Projectile(.Graphic).HasData, Tex_Projectile(.Graphic).Width, Tex_Projectile(.Graphic).Height)
                    Else
                        Call RenderTexture(Tex_Projectile(.Graphic), ConvertMapX(X), ConvertMapY(Y), 0, 0, Tex_Projectile(.Graphic).Width, Tex_Projectile(.Graphic).Height, Tex_Projectile(.Graphic).Width, Tex_Projectile(.Graphic).Height, , .Rotate)
                    End If
                    
                End If
            End With
        Next
        
        ' ****** Erase Projectile ******    Seperate Loop For Erasing
        For i = 1 To LastProjectile
            If ProjectileList(i).Graphic Then
                If Abs(ProjectileList(i).X - ProjectileList(i).tx) < 20 Then
                    If Abs(ProjectileList(i).Y - ProjectileList(i).ty) < 20 Then
                        Call ClearProjectile(i)
                    End If
                End If
            End If
        Next
        
    End If
End Sub

Private Sub UpdateEffectOffset(ByVal EffectIndex As Integer)
    EffectData(EffectIndex).X = ConvertMapX(EffectData(EffectIndex).X + (LastOffsetX - ParticleOffsetX))
    EffectData(EffectIndex).Y = ConvertMapY(EffectData(EffectIndex).Y + (LastOffsetY - ParticleOffsetY))
End Sub

Private Sub UpdateEffectBinding(ByVal EffectIndex As Integer)
'***************************************************
'Updates the binding of a particle effect to a target, if
'the effect is bound to a character
'***************************************************
Dim TargetA As Single
 
    'Update position through character binding
    If EffectData(EffectIndex).BindIndex > 0 Then
        'Calculate the X and Y positions
        Select Case EffectData(EffectIndex).BindType
            Case TARGET_TYPE_PLAYER
                EffectData(EffectIndex).GoToX = ConvertMapX(((Player(EffectData(EffectIndex).BindIndex).X * 32)) + TempPlayer(EffectData(EffectIndex).BindIndex).xOffset) + 16
                EffectData(EffectIndex).GoToY = ConvertMapY(((Player(EffectData(EffectIndex).BindIndex).Y * 32)) + TempPlayer(EffectData(EffectIndex).BindIndex).YOffset) + 32
            Case TARGET_TYPE_NPC
                EffectData(EffectIndex).GoToX = ConvertMapX(((MapNpc(EffectData(EffectIndex).BindIndex).X * 32)) + TempMapNpc(EffectData(EffectIndex).BindIndex).xOffset) + 16
                EffectData(EffectIndex).GoToY = ConvertMapY(((MapNpc(EffectData(EffectIndex).BindIndex).Y * 32)) + TempMapNpc(EffectData(EffectIndex).BindIndex).YOffset) + 32
        End Select
    End If
    
    'Move to the new position if needed
    If EffectData(EffectIndex).GoToX > -30000 Or EffectData(EffectIndex).GoToY > -30000 Then
        If EffectData(EffectIndex).GoToX <> EffectData(EffectIndex).X Or EffectData(EffectIndex).GoToY <> EffectData(EffectIndex).Y Then
 
            'Calculate the angle
            TargetA = Engine_GetAngle((EffectData(EffectIndex).X), (EffectData(EffectIndex).Y), EffectData(EffectIndex).GoToX, EffectData(EffectIndex).GoToY) + 180
 
            'Update the position of the effect
            EffectData(EffectIndex).X = EffectData(EffectIndex).X - Sin(TargetA * DegreeToRadian) * EffectData(EffectIndex).BindSpeed
            EffectData(EffectIndex).Y = EffectData(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * EffectData(EffectIndex).BindSpeed
 
            'Check if the effect is close enough to the target to just stick it at the target
            'If EffectData(EffectIndex).GoToX > -30000 Then
                'If Abs(EffectData(EffectIndex).X - EffectData(EffectIndex).GoToX) < 2 Then
                    EffectData(EffectIndex).X = EffectData(EffectIndex).GoToX
                'End If
            'End If
            
            'If EffectData(EffectIndex).GoToY > -30000 Then
            '    If Abs(EffectData(EffectIndex).Y - EffectData(EffectIndex).GoToY) < 2 Then
                    EffectData(EffectIndex).Y = EffectData(EffectIndex).GoToY
            '    End If
            'End If
 
            'Check if the position of the effect is equal to that of the target
            If EffectData(EffectIndex).X = EffectData(EffectIndex).GoToX Then
                If EffectData(EffectIndex).Y = EffectData(EffectIndex).GoToY Then
 
                    'For some effects, if the position is reached, we want to end the effect
                    If EffectData(EffectIndex).KillWhenAtTarget Then
                        EffectData(EffectIndex).BindIndex = 0
                        EffectData(EffectIndex).Progression = 0
                        EffectData(EffectIndex).GoToX = EffectData(EffectIndex).X
                        EffectData(EffectIndex).GoToY = EffectData(EffectIndex).Y
                    End If
                    Exit Sub    'The effect is at the right position, don't update
 
                End If
            End If
 
        End If
    End If
 
End Sub

Private Function Effect_FToDW(F As Single) As Long
'*****************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'*****************************************************************
Dim Buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set Buf = Direct3DX.CreateBuffer(4)
    Direct3DX.BufferSetData Buf, 0, 4, 1, F
    Direct3DX.BufferGetData Buf, 0, 4, 1, Effect_FToDW

End Function

Sub Effect_Kill(ByVal EffectIndex As Integer, Optional ByVal KillAll As Boolean = False)
'*****************************************************************
'Kills (stops) a single effect or all effects
'*****************************************************************
Dim LoopC As Long

    'Check If To Kill All Effects
    If KillAll = True Then

        'Loop Through Every Effect
        For LoopC = 1 To NumEffects

            'Stop The Effect
            EffectData(LoopC).Used = False

        Next
        
    Else

        'Stop The Selected Effect
        EffectData(EffectIndex).Used = False
        
    End If

End Sub

Private Function Effect_NextOpenSlot() As Integer
'*****************************************************************
'Finds the next open effects index
'*****************************************************************
Dim EffectIndex As Integer

    'Find The Next Open Effect Slot
    Do
        EffectIndex = EffectIndex + 1   'Check The Next Slot
        If EffectIndex > NumEffects Then    'Dont Go Over Maximum Amount
            Effect_NextOpenSlot = -1
            Exit Function
        End If
    Loop While EffectData(EffectIndex).Used = True    'Check Next If Effect Is In Use

    'Return the next open slot
    Effect_NextOpenSlot = EffectIndex

    'Clear the old information from the effect
    Erase EffectData(EffectIndex).Particles()
    Erase EffectData(EffectIndex).PartVertex()
    ZeroMemory EffectData(EffectIndex), LenB(EffectData(EffectIndex))
    EffectData(EffectIndex).GoToX = -30000
    EffectData(EffectIndex).GoToY = -30000

End Function

Public Sub RenderEffectData(ByVal EffectIndex As Integer, Optional ByVal SetRenderStates As Boolean = True)

    'Check if we have the device
    If Direct3D_Device.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    If EffectData(EffectIndex).Gfx > NumParticles Then Exit Sub

    'Set the render state for the size of the particle
    Call Direct3D_Device.SetRenderState(D3DRS_POINTSIZE, EffectData(EffectIndex).FloatSize)
    
    'Set the render state to point blitting
    If SetRenderStates Then Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
    'Set the texture
    SetTexture Tex_Particle(EffectData(EffectIndex).Gfx)
    Direct3D_Device.SetTexture 0, gTexture(Tex_Particle(EffectData(EffectIndex).Gfx).Texture).Texture

    'Draw all the particles at once
    Direct3D_Device.DrawPrimitiveUP D3DPT_POINTLIST, EffectData(EffectIndex).ParticleCount, EffectData(EffectIndex).PartVertex(0), Len(EffectData(EffectIndex).PartVertex(0))

    'Reset the render state back to normal
    If SetRenderStates Then Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Sub UpdateEffectAll()
'*****************************************************************
'Updates all of the effects and renders them
'*****************************************************************
Dim LoopC As Long

    'Make sure we have effects
    If NumEffects = 0 Then Exit Sub

    'Set the render state for the particle effects
    Call Direct3D_Device.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_ONE)

    'Update every effect in use
    For LoopC = 1 To NumEffects

        'Make sure the effect is in use
        If EffectData(LoopC).Used Then

            Call UpdateEffectOffset(LoopC)
        
            'Update the effect position if it is binded
            Call UpdateEffectBinding(LoopC)

            'Find out which effect is selected, then update it
            Select Case EffectData(LoopC).EffectNum
                Case EFFECT_TYPE_HEAL: Heal_Update LoopC
                Case EFFECT_TYPE_PROTECTION: Protection_Update LoopC
                Case EFFECT_TYPE_STRENGTHEN: Strengthen_Update LoopC
                Case EFFECT_TYPE_SUMMON: Summon_Update LoopC
            End Select
            
            'Render the effect
            Call RenderEffectData(LoopC, False)

        End If

    Next
    
    'Set the render state back for normal rendering
    Call Direct3D_Device.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)

End Sub

Public Sub Engine_Init_Particles()
Dim LoopI As Byte

    'Set the particles texture
    NumEffects = 30
    ReDim EffectData(1 To NumEffects)

End Sub

Function Engine_PixelPosX(ByVal X As Integer) As Integer
' **************************************************
' * Converts a Tile Position Into a Pixel Position *
' **************************************************
    Engine_PixelPosX = X * PIC_X
End Function

Function Engine_PixelPosY(ByVal Y As Integer) As Integer
' **************************************************
' * Converts a Tile Position Into a Pixel Position *
' **************************************************
    Engine_PixelPosY = Y * PIC_Y
End Function

Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
' ***************************************************
' * Converts a Tile Position Into a Screen Position *
' ***************************************************
    Engine_TPtoSPX = Engine_PixelPosX(X)
End Function

Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
' ***************************************************
' * Converts a Tile Position Into a Screen Position *
' ***************************************************
    Engine_TPtoSPY = Engine_PixelPosY(Y)
End Function

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then

        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90

            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If

        'Exit the function
        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then

        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Note: Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function

Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

Exit Function

End Function












Function Summon_Begin(ByVal EffectNum As Long, ByVal X As Single, ByVal Y As Single) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Summon_Begin = EffectIndex

    'Set The Effect's Variables
    EffectData(EffectIndex).EffectNum = EFFECT_TYPE_SUMMON    'Set the effect number
    EffectData(EffectIndex).Used = True                     'Enable the effect
    EffectData(EffectIndex).X = X                          'Set the effect's X coordinate
    EffectData(EffectIndex).Y = Y                         'Set the effect's Y coordinate
    EffectData(EffectIndex).ParticleCount = Effect(EffectNum).Particles           'Set the number of particles
    EffectData(EffectIndex).Gfx = Effect(EffectNum).Sprite               'Set the graphic
    EffectData(EffectIndex).Alpha = Effect(EffectNum).Alpha
    EffectData(EffectIndex).Decay = Effect(EffectNum).Decay
    EffectData(EffectIndex).Red = Effect(EffectNum).Red
    EffectData(EffectIndex).Green = Effect(EffectNum).Green
    EffectData(EffectIndex).Blue = Effect(EffectNum).Blue
    EffectData(EffectIndex).XSpeed = Effect(EffectNum).XSpeed
    EffectData(EffectIndex).YSpeed = Effect(EffectNum).YSpeed
    EffectData(EffectIndex).XAcc = Effect(EffectNum).XAcc
    EffectData(EffectIndex).YAcc = Effect(EffectNum).YAcc
    EffectData(EffectIndex).Modifier = 30         'How large the circle is
    EffectData(EffectIndex).Progression = Effect(EffectNum).Duration      'How long the effect will last
    EffectData(EffectIndex).FloatSize = Effect_FToDW(Effect(EffectNum).Size)    'Size of the particles
    
    EffectData(EffectIndex).ParticlesLeft = EffectData(EffectIndex).ParticleCount

    'Redim the number of particles
    ReDim EffectData(EffectIndex).Particles(0 To EffectData(EffectIndex).ParticleCount)
    ReDim EffectData(EffectIndex).PartVertex(0 To EffectData(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To EffectData(EffectIndex).ParticleCount
        Set EffectData(EffectIndex).Particles(LoopC) = New clsParticle
        EffectData(EffectIndex).Particles(LoopC).Used = True
        EffectData(EffectIndex).PartVertex(LoopC).RHW = 1
        Summon_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    EffectData(EffectIndex).PreviousFrame = GetTickCount

End Function

Private Sub Summon_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim R As Single
    
    If EffectData(EffectIndex).Progression > 1000 Then
        EffectData(EffectIndex).Progression = EffectData(EffectIndex).Progression + 1.4
    Else
        EffectData(EffectIndex).Progression = EffectData(EffectIndex).Progression + 0.5
    End If
    R = (Index / 30) * EXP(Index / EffectData(EffectIndex).Progression)
    X = R * Cos(Index)
    Y = R * Sin(Index)
    
    'Reset the particle
    EffectData(EffectIndex).Particles(Index).ResetIt ConvertMapX(EffectData(EffectIndex).X) + X, ConvertMapY(EffectData(EffectIndex).Y) + Y, EffectData(EffectIndex).XSpeed, EffectData(EffectIndex).YSpeed, EffectData(EffectIndex).XAcc, EffectData(EffectIndex).YAcc
    EffectData(EffectIndex).Particles(Index).ResetColor EffectData(EffectIndex).Red / 100, EffectData(EffectIndex).Green / 100 + (Rnd * 0.2), EffectData(EffectIndex).Blue / 100, EffectData(EffectIndex).Alpha / 100, EffectData(EffectIndex).Decay / 100 + (Rnd * 0.2)
 
End Sub

Private Sub Summon_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (GetTickCount - EffectData(EffectIndex).PreviousFrame) * 0.01
    EffectData(EffectIndex).PreviousFrame = GetTickCount
    'Go Through The Particle Loop
    For LoopC = 0 To EffectData(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If EffectData(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            EffectData(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If EffectData(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If EffectData(EffectIndex).Progression < 1800 Then

                    'Reset the particle
                    Summon_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    EffectData(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    EffectData(EffectIndex).ParticlesLeft = EffectData(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If EffectData(EffectIndex).ParticlesLeft = 0 Then EffectData(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    EffectData(EffectIndex).PartVertex(LoopC).color = 0

                End If

            Else
            
                'Set the particle information on the particle vertex
                EffectData(EffectIndex).PartVertex(LoopC).color = D3DColorMake(EffectData(EffectIndex).Particles(LoopC).sngR, EffectData(EffectIndex).Particles(LoopC).sngG, EffectData(EffectIndex).Particles(LoopC).sngB, EffectData(EffectIndex).Particles(LoopC).sngA)
                EffectData(EffectIndex).PartVertex(LoopC).X = EffectData(EffectIndex).Particles(LoopC).sngX
                EffectData(EffectIndex).PartVertex(LoopC).Y = EffectData(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Strengthen_Begin(ByVal EffectNum As Long, ByVal X As Single, ByVal Y As Single) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Strengthen_Begin = EffectIndex

    'Set the effect's variables
    EffectData(EffectIndex).EffectNum = EFFECT_TYPE_STRENGTHEN    'Set the effect number
    EffectData(EffectIndex).Used = True             'Enabled the effect
    EffectData(EffectIndex).X = X               'Set the effect's X coordinate
    EffectData(EffectIndex).Y = Y                 'Set the effect's Y coordinate
    EffectData(EffectIndex).ParticleCount = Effect(EffectNum).Particles           'Set the number of particles
    EffectData(EffectIndex).Gfx = Effect(EffectNum).Sprite               'Set the graphic
    EffectData(EffectIndex).Alpha = Effect(EffectNum).Alpha
    EffectData(EffectIndex).Decay = Effect(EffectNum).Decay
    EffectData(EffectIndex).Red = Effect(EffectNum).Red
    EffectData(EffectIndex).Green = Effect(EffectNum).Green
    EffectData(EffectIndex).Blue = Effect(EffectNum).Blue
    EffectData(EffectIndex).XSpeed = Effect(EffectNum).XSpeed
    EffectData(EffectIndex).YSpeed = Effect(EffectNum).YSpeed
    EffectData(EffectIndex).XAcc = Effect(EffectNum).XAcc
    EffectData(EffectIndex).YAcc = Effect(EffectNum).YAcc
    EffectData(EffectIndex).Modifier = 30         'How large the circle is
    EffectData(EffectIndex).Progression = Effect(EffectNum).Duration      'How long the effect will last
    EffectData(EffectIndex).FloatSize = Effect_FToDW(Effect(EffectNum).Size)    'Size of the particles
    
    'Set the number of particles left to the total avaliable
    EffectData(EffectIndex).ParticlesLeft = EffectData(EffectIndex).ParticleCount

    'Redim the number of particles
    ReDim EffectData(EffectIndex).Particles(0 To EffectData(EffectIndex).ParticleCount)
    ReDim EffectData(EffectIndex).PartVertex(0 To EffectData(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To EffectData(EffectIndex).ParticleCount
        Set EffectData(EffectIndex).Particles(LoopC) = New clsParticle
        EffectData(EffectIndex).Particles(LoopC).Used = True
        EffectData(EffectIndex).PartVertex(LoopC).RHW = 1
        Strengthen_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    EffectData(EffectIndex).PreviousFrame = GetTickCount

End Function

Private Sub Strengthen_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Reset
'*****************************************************************
Dim a As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    X = EffectData(EffectIndex).X - (Sin(a) * EffectData(EffectIndex).Modifier)
    Y = EffectData(EffectIndex).Y + (Cos(a) * EffectData(EffectIndex).Modifier)

    'Reset the particle
    EffectData(EffectIndex).Particles(Index).ResetIt ConvertMapX(X), ConvertMapY(Y), EffectData(EffectIndex).XSpeed, Rnd * EffectData(EffectIndex).YSpeed, EffectData(EffectIndex).XAcc, EffectData(EffectIndex).YAcc
    EffectData(EffectIndex).Particles(Index).ResetColor EffectData(EffectIndex).Red / 100, EffectData(EffectIndex).Green / 100, EffectData(EffectIndex).Blue / 100, EffectData(EffectIndex).Alpha / 100 + (Rnd * 0.4), EffectData(EffectIndex).Decay / 100 + (Rnd * 0.2)
End Sub

Private Sub Strengthen_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate the time difference
    ElapsedTime = (GetTickCount - EffectData(EffectIndex).PreviousFrame) * 0.01
    EffectData(EffectIndex).PreviousFrame = GetTickCount

    'Update the life span
    If EffectData(EffectIndex).Progression > 0 Then EffectData(EffectIndex).Progression = EffectData(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To EffectData(EffectIndex).ParticleCount

        'Check if particle is in use
        If EffectData(EffectIndex).Particles(LoopC).Used Then

            'Update the particle
            EffectData(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If EffectData(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If EffectData(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Strengthen_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    EffectData(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    EffectData(EffectIndex).ParticlesLeft = EffectData(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If EffectData(EffectIndex).ParticlesLeft = 0 Then EffectData(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    EffectData(EffectIndex).PartVertex(LoopC).color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                EffectData(EffectIndex).PartVertex(LoopC).color = D3DColorMake(EffectData(EffectIndex).Particles(LoopC).sngR, EffectData(EffectIndex).Particles(LoopC).sngG, EffectData(EffectIndex).Particles(LoopC).sngB, EffectData(EffectIndex).Particles(LoopC).sngA)
                EffectData(EffectIndex).PartVertex(LoopC).X = EffectData(EffectIndex).Particles(LoopC).sngX
                EffectData(EffectIndex).PartVertex(LoopC).Y = EffectData(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Protection_Begin(ByVal EffectNum As Long, ByVal X As Single, ByVal Y As Single) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Protection_Begin = EffectIndex

    'Set The Effect's Variables
    EffectData(EffectIndex).EffectNum = EFFECT_TYPE_PROTECTION    'Set the effect number
    EffectData(EffectIndex).Used = True             'Enabled the effect
    EffectData(EffectIndex).X = X                   'Set the effect's X coordinate
    EffectData(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    EffectData(EffectIndex).ParticleCount = Effect(EffectNum).Particles           'Set the number of particles
    EffectData(EffectIndex).Gfx = Effect(EffectNum).Sprite               'Set the graphic
    EffectData(EffectIndex).Alpha = Effect(EffectNum).Alpha
    EffectData(EffectIndex).Decay = Effect(EffectNum).Decay
    EffectData(EffectIndex).Red = Effect(EffectNum).Red
    EffectData(EffectIndex).Green = Effect(EffectNum).Green
    EffectData(EffectIndex).Blue = Effect(EffectNum).Blue
    EffectData(EffectIndex).XSpeed = Effect(EffectNum).XSpeed
    EffectData(EffectIndex).YSpeed = Effect(EffectNum).YSpeed
    EffectData(EffectIndex).XAcc = Effect(EffectNum).XAcc
    EffectData(EffectIndex).YAcc = Effect(EffectNum).YAcc
    EffectData(EffectIndex).Modifier = 30         'How large the circle is
    EffectData(EffectIndex).Progression = Effect(EffectNum).Duration      'How long the effect will last
    EffectData(EffectIndex).FloatSize = Effect_FToDW(Effect(EffectNum).Size)    'Size of the particles
    
    
    'Set the number of particles left to the total avaliable
    EffectData(EffectIndex).ParticlesLeft = EffectData(EffectIndex).ParticleCount

    'Redim the number of particles
    ReDim EffectData(EffectIndex).Particles(0 To EffectData(EffectIndex).ParticleCount)
    ReDim EffectData(EffectIndex).PartVertex(0 To EffectData(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To EffectData(EffectIndex).ParticleCount
        Set EffectData(EffectIndex).Particles(LoopC) = New clsParticle
        EffectData(EffectIndex).Particles(LoopC).Used = True
        EffectData(EffectIndex).PartVertex(LoopC).RHW = 1
        Protection_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    EffectData(EffectIndex).PreviousFrame = GetTickCount

End Function

Private Sub Protection_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Reset
'*****************************************************************
Dim a As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    X = EffectData(EffectIndex).X - (Sin(a) * EffectData(EffectIndex).Modifier)
    Y = EffectData(EffectIndex).Y + (Cos(a) * EffectData(EffectIndex).Modifier)

    'Reset the particle
    EffectData(EffectIndex).Particles(Index).ResetIt ConvertMapX(X), ConvertMapY(Y), EffectData(EffectIndex).XSpeed, Rnd * EffectData(EffectIndex).YSpeed, EffectData(EffectIndex).XAcc, EffectData(EffectIndex).YAcc
    EffectData(EffectIndex).Particles(Index).ResetColor EffectData(EffectIndex).Red / 100, EffectData(EffectIndex).Green / 100, EffectData(EffectIndex).Blue / 100, EffectData(EffectIndex).Alpha / 100 + (Rnd * 0.4), EffectData(EffectIndex).Decay / 100 + (Rnd * 0.2)

End Sub
Private Sub Protection_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (GetTickCount - EffectData(EffectIndex).PreviousFrame) * 0.01
    EffectData(EffectIndex).PreviousFrame = GetTickCount

    'Update the life span
    If EffectData(EffectIndex).Progression > 0 Then EffectData(EffectIndex).Progression = EffectData(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To EffectData(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If EffectData(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            EffectData(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If EffectData(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If EffectData(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Protection_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    EffectData(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    EffectData(EffectIndex).ParticlesLeft = EffectData(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If EffectData(EffectIndex).ParticlesLeft = 0 Then EffectData(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    EffectData(EffectIndex).PartVertex(LoopC).color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                EffectData(EffectIndex).PartVertex(LoopC).color = D3DColorMake(EffectData(EffectIndex).Particles(LoopC).sngR, EffectData(EffectIndex).Particles(LoopC).sngG, EffectData(EffectIndex).Particles(LoopC).sngB, EffectData(EffectIndex).Particles(LoopC).sngA)
                EffectData(EffectIndex).PartVertex(LoopC).X = EffectData(EffectIndex).Particles(LoopC).sngX
                EffectData(EffectIndex).PartVertex(LoopC).Y = EffectData(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Heal_Begin(ByVal EffectNum As Long, ByVal X As Single, ByVal Y As Single) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Heal_Begin = EffectIndex

    'Set The Effect's Variables
    EffectData(EffectIndex).EffectNum = EFFECT_TYPE_HEAL      'Set the effect number
    EffectData(EffectIndex).Used = True     'Enabled the effect
    EffectData(EffectIndex).X = X                   'Set the effect's X coordinate
    EffectData(EffectIndex).Y = Y + 32                   'Set the effect's Y coordinate
    EffectData(EffectIndex).ParticleCount = Effect(EffectNum).Particles           'Set the number of particles
    EffectData(EffectIndex).Gfx = Effect(EffectNum).Sprite               'Set the graphic
    EffectData(EffectIndex).Alpha = Effect(EffectNum).Alpha
    EffectData(EffectIndex).Decay = Effect(EffectNum).Decay
    EffectData(EffectIndex).Red = Effect(EffectNum).Red
    EffectData(EffectIndex).Green = Effect(EffectNum).Green
    EffectData(EffectIndex).Blue = Effect(EffectNum).Blue
    EffectData(EffectIndex).XSpeed = Effect(EffectNum).XSpeed
    EffectData(EffectIndex).YSpeed = Effect(EffectNum).YSpeed
    EffectData(EffectIndex).XAcc = Effect(EffectNum).XAcc
    EffectData(EffectIndex).YAcc = Effect(EffectNum).YAcc
    EffectData(EffectIndex).Modifier = 30         'How large the circle is
    EffectData(EffectIndex).Progression = Effect(EffectNum).Duration      'How long the effect will last
    EffectData(EffectIndex).FloatSize = Effect_FToDW(Effect(EffectNum).Size)    'Size of the particles
    
    'Set the number of particles left to the total avaliable
    EffectData(EffectIndex).ParticlesLeft = EffectData(EffectIndex).ParticleCount

    'Redim the number of particles
    ReDim EffectData(EffectIndex).Particles(0 To EffectData(EffectIndex).ParticleCount)
    ReDim EffectData(EffectIndex).PartVertex(0 To EffectData(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To EffectData(EffectIndex).ParticleCount
        Set EffectData(EffectIndex).Particles(LoopC) = New clsParticle
        EffectData(EffectIndex).Particles(LoopC).Used = True
        EffectData(EffectIndex).PartVertex(LoopC).RHW = 1
        Heal_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    EffectData(EffectIndex).PreviousFrame = GetTickCount

End Function

Private Sub Heal_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Reset
'*****************************************************************

    'Reset the particle
    EffectData(EffectIndex).Particles(Index).ResetIt ConvertMapX(EffectData(EffectIndex).X - 10 + Rnd * 20), ConvertMapY(EffectData(EffectIndex).Y - 10 + Rnd * 20), -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), EffectData(EffectIndex).XAcc, EffectData(EffectIndex).YAcc
    EffectData(EffectIndex).Particles(Index).ResetColor EffectData(EffectIndex).Red / 100, EffectData(EffectIndex).Green / 100, EffectData(EffectIndex).Blue / 100, EffectData(EffectIndex).Alpha / 100 + (Rnd * 0.2), EffectData(EffectIndex).Decay / 100 + (Rnd * 0.5)
    
End Sub

Private Sub Heal_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
Dim i As Integer

    'Calculate the time difference
    ElapsedTime = (GetTickCount - EffectData(EffectIndex).PreviousFrame) * 0.01
    EffectData(EffectIndex).PreviousFrame = GetTickCount
    If EffectData(EffectIndex).Progression > 0 Then EffectData(EffectIndex).Progression = EffectData(EffectIndex).Progression - ElapsedTime
    'Go through the particle loop
    For LoopC = 0 To EffectData(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If EffectData(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            EffectData(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If EffectData(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If EffectData(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Heal_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    EffectData(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    EffectData(EffectIndex).ParticlesLeft = EffectData(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If EffectData(EffectIndex).ParticlesLeft = 0 Then EffectData(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    EffectData(EffectIndex).PartVertex(LoopC).color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                EffectData(EffectIndex).PartVertex(LoopC).color = D3DColorMake(EffectData(EffectIndex).Particles(LoopC).sngR, EffectData(EffectIndex).Particles(LoopC).sngG, EffectData(EffectIndex).Particles(LoopC).sngB, EffectData(EffectIndex).Particles(LoopC).sngA)
                EffectData(EffectIndex).PartVertex(LoopC).X = EffectData(EffectIndex).Particles(LoopC).sngX
                EffectData(EffectIndex).PartVertex(LoopC).Y = EffectData(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Public Sub DrawNews()
Dim i As Long, X As Long, Y As Long, Sprite As Long, Width As Long
Dim Height As Long
    ' draw background
    X = (frmMain.ScaleWidth / 2) - (GUIWindow(GUI_NEWS).Width / 2)
    Y = (frmMain.ScaleHeight / 2) - (GUIWindow(GUI_NEWS).Height / 2)
    
    ' render chatbox
    Width = GUIWindow(GUI_NEWS).Width
    Height = GUIWindow(GUI_NEWS).Height
    'EngineRenderRectangle Tex_GUI(21), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(27), X, Y, 0, 0, Width, Height, Width, Height
    ' Draw the text
    RenderText Font_Default, WordWrap(Trim$(NewsText), Width - 70), X + 35, Y + 45, White
End Sub

Sub DrawAmbient()
    Dim i As Long, X As Long, Y As Long, AnimTickTime As Long
    
    If Map.Ambiente > 0 Then
            For i = 1 To 30
                If AmbientActor(i).Used = True Then
        
                    AmbientActor(i).Y = AmbientActor(i).Y + AmbientActor(i).speed
                    AmbientActor(i).X = AmbientActor(i).X + AmbientActor(i).Dir
                    
                    If Map.Ambiente = 1 Then
                    AmbientActor(i).Dir = AmbientActor(i).Dir + Rand(-1, 1)
                    
                    If AmbientActor(i).Dir > 5 Then AmbientActor(i).Dir = 5
                    If AmbientActor(i).Dir < -5 Then AmbientActor(i).Dir = -5
                    End If
                    
                    If Map.Ambiente = 1 Then
                        AnimTickTime = 1
                    End If
                    
                    If Map.Ambiente = 2 Then
                        AmbientActor(i).Dir = 0
                        AnimTickTime = 500
                    End If
                    
                    If Map.Ambiente = 3 Then
                        AnimTickTime = 100
                    End If
                    
                    If AmbientActor(i).AnimTick + AnimTickTime < GetTickCount Then
                        AmbientActor(i).Animation = Rand(0, 1)
                        AmbientActor(i).AnimTick = GetTickCount
                    End If
                    
                    X = ConvertMapX(AmbientActor(i).X)
                    Y = ConvertMapY(AmbientActor(i).Y)
                    
                    If Y > ScreenHeight Then
                        AmbientActor(i).Used = False
                    End If
                    
                    If Map.Ambiente > 1 Then
                        RenderTexture Tex_Ambiente, X, Y, AmbientActor(i).Animation * 32, (Map.Ambiente - 1) * 32, 32 + (i / 2), 32 + (i / 2), 32, 32
                    Else
                        RenderTexture Tex_Ambiente, X, Y, AmbientActor(i).Animation * 32, (Map.Ambiente - 1) * 32, 32, 32, 32, 32
                    End If
                    
                    Else
                    
                    Dim Chance As Integer
                    
                    Chance = 800
                    If Tremor > GetTickCount Then Chance = 100
                    
                    'Create actor
                    If Rand(1, Chance) = 1 Then
                        AmbientActor(i).Used = True
                        
                        If Map.Ambiente = 1 Then
                        AmbientActor(i).speed = Rand(5, 15)
                        AmbientActor(i).Dir = Rand(-10, 10)
                        End If
                        
                        If Map.Ambiente = 2 Then
                        AmbientActor(i).speed = Rand(2, 5)
                        AmbientActor(i).Dir = 0
                        End If
                        
                        If Map.Ambiente = 3 Then
                        AmbientActor(i).speed = Rand(3, 9)
                        AmbientActor(i).Dir = Rand(-3, 3)
                        End If
                        
                        AmbientActor(i).X = Rand(0, Map.MaxX * 32)
                        AmbientActor(i).Y = 0
                    End If
                    
                End If
            Next i
    End If
End Sub
Sub DrawFishAlert(Index As Long)
    Dim X As Long, Y As Long, sY As Long
    Static Shout As Long
    Static PauseTick As Long
    
    X = (GetPlayerX(Index) * PIC_X)
    Y = (GetPlayerY(Index) * PIC_Y) - 32
    
    If Index = MyIndex Then
        If GetPlayerEquipment(Index, Weapon) > 0 Then
            If Item(GetPlayerEquipment(Index, Weapon)).Data3 = 2 Then
                If FishTime > GetTickCount Then
                    X = X + Rand(-2, 2)
                    sY = 0
                    BubbleOpaque = 255
                    PauseTick = GetTickCount + 5000
                Else
                    If PauseTick > GetTickCount Then
                        BubbleOpaque = 0
                        Exit Sub
                    End If
                    
                    If Rand(1, 500) = 1 And BubbleOpaque = 255 Or Shout > GetTickCount Then
                        If Shout < GetTickCount Then Shout = GetTickCount + 500
                        sY = 64
                        BubbleOpaque = 255
                    Else
                        sY = 32
                        If BubbleOpaque < 255 Then BubbleOpaque = BubbleOpaque + 1
                        If GetTickCount Mod 1000 < 500 Then
                        Y = Y + 1
                        End If
                    End If
                    If AlertX > GetTickCount Then
                        sY = 96
                        Y = Y + Rand(-2, 2)
                        BubbleOpaque = 0
                    End If
                End If
                Else
                BubbleOpaque = 0
                Exit Sub
            End If
            Else
            BubbleOpaque = 0
            Exit Sub
        End If
    Else
        'BubbleOpaque = 255
        If GetTickCount Mod 1000 < 500 Then
        Y = Y + 1
        End If
        sY = 32
    End If
    
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    If Index = MyIndex And AlertX < GetTickCount Then
        RenderTexture Tex_Alerta, X, Y, 0, sY, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, BubbleOpaque)
    Else
        RenderTexture Tex_Alerta, X, Y, 0, sY, 32, 32, 32, 32
    End If
End Sub

Sub DrawBuracos()
Dim n As Long
For n = 1 To 10
    If Buracos(n).InUse = True Then
    If Buracos(n).IntervalTick < GetTickCount Then Buracos(n).Alpha = Buracos(n).Alpha - 1
    If Buracos(n).Alpha = 0 Then Buracos(n).InUse = False
    RenderTexture Tex_Buraco, ConvertMapX(Buracos(n).X * PIC_X) - Int(Buracos(n).Size / 4) + -(8 * Int((Buracos(n).Size - 64) / 32)), ConvertMapY(Buracos(n).Y * PIC_Y) - Int(Buracos(n).Size / 4) + -(16 * Int((Buracos(n).Size - 64) / 32)), 0, 0, Buracos(n).Size, Buracos(n).Size, Tex_Buraco.Width, Tex_Buraco.Height, D3DColorRGBA(255, 255, 255, Buracos(n).Alpha)
    End If
Next n
End Sub

Public Sub DrawTransporte()
Dim i As Double, X As Long, Y As Long, Sprite As Long, Width As Long
Dim Height As Long

    If Transporte.Map <> GetPlayerMap(MyIndex) Then Exit Sub

    Select Case Transporte.Tipo
    
        Case 1 'Avião
        Y = 70
        If Transporte.Anim = 0 Then
            Y = 70
            X = -50
            Transporte.X = -50
        End If
        
        If Transporte.Anim = 1 Then
            ' Chegando
            If Transporte.Tick + 5000 > GetTickCount Then
            'X = -50
                Y = 70
        
                Transporte.X = Int(Transporte.X * 0.97)
                
                If Transporte.X > -50 Then Transporte.X = -50
                
                X = Transporte.X
            Else
                Y = 70
                X = Transporte.X
            End If
        End If
            
        If Transporte.Anim = 2 Then
                Y = 70
            
                If Transporte.X < 10 Then
                Transporte.X = Transporte.X + 0.5
                Else
                Transporte.X = Transporte.X * 1.05
                End If
                
                X = Transporte.X
        End If
            
            If Transporte.X > frmMain.ScaleWidth Then Transporte.Tipo = 0
        
        ' render chatbox
        Width = 471
        Height = 181
        RenderTexture Tex_Transportes(1), X, Y, 0, 0, Width, Height, Width, Height
        
        Case 2 'Navio
        Y = 270
        
        If Transporte.Anim = 0 Then
            Y = 270
            X = 400
            Transporte.X = 400
        End If
        
        If Transporte.Anim = 1 Then
            ' Chegando
            If Transporte.Tick + 50000 > GetTickCount Then
                Y = 270
        
                Transporte.X = Transporte.X + 1
                
                If Transporte.X > 400 Then Transporte.X = 400
                
                X = Transporte.X
            Else
                Y = 270
                X = Transporte.X
            End If
        End If
            
        If Transporte.Anim = 2 Then
                Y = 270

                Transporte.X = Transporte.X + 1
                
                X = Transporte.X
        End If
            
            If Transporte.X > frmMain.ScaleWidth Then Transporte.Tipo = 0

        Width = 565
        Height = 291
        RenderTexture Tex_Transportes(2), X, Y, 0, 0, Width, Height, Width, Height
    
    End Select
End Sub

Public Sub EditorQuest_DrawIcon()
Dim iconNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    iconNum = frmEditor_Quest.scrlIcon.Value

    If iconNum < 1 Or iconNum > numitems Then
        frmEditor_Quest.picIcon.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(iconNum), sRECT, dRect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Quest.picIcon.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawPaperdollTest()

Dim X, Y As Long

    Dim Sprite As Long
    Sprite = 1

' Calculate the X
    If VXFRAME = False Then
        X = GetPlayerX(MyIndex) * PIC_X + TempPlayer(MyIndex).xOffset - ((Tex_Character(Sprite).Width / 22 - 32) / 2)
    Else
        X = GetPlayerX(MyIndex) * PIC_X + TempPlayer(MyIndex).xOffset - ((Tex_Character(Sprite).Width / 3 - 32) / 2)
    End If
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(MyIndex) * PIC_Y + TempPlayer(MyIndex).YOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = GetPlayerY(MyIndex) * PIC_Y + TempPlayer(MyIndex).YOffset
    End If
    
    'novas sprites
    Y = Y + 16
    
    Dim rec As RECT
    With rec
        .Top = 0
        .Bottom = Tex_Character(Sprite).Height
        .Left = 0
        .Right = Tex_Character(Sprite).Width
    End With
    If frmPaperdoll.chkPlayer = 1 Then Call DrawSprite(Sprite, X, Y, rec, False, 0)
    
    Dim i As Long
    For i = 1 To 6
        Dim hairrec As RECT
        Dim SpriteLeft As Byte
        If i > 3 Then SpriteLeft = 1
        With hairrec
            .Top = 32 * ((i - 1) Mod 3)
            .Bottom = .Top + 32
            .Left = 32 * SpriteLeft
            .Right = .Left + 32
        End With
        
        HairTest(i) = frmPaperdoll.txtHair(i - 1).Text
        
        Dim n As Long
        Dim Positions() As String
        Positions() = Split(HairTest(i), ";")
        For n = 1 To UBound(Positions)
            Dim hairX, hairY As Long
            If Val(Positions(n)) > 0 Then
                hairX = X + (((Val(Positions(n)) Mod 21) * 32) - 32) + 4 + GlobalRepositionX + HairRepositionX(i) + PositionRepositionX(Positions(n))
                hairY = Y + ((Int(Val(Positions(n)) / 21) * 64)) + GlobalRepositionY + HairRepositionY(i) + PositionRepositionY(Positions(n))
                If Val(Positions(n)) = 21 Or Val(Positions(n)) = 42 Or Val(Positions(n)) = 63 Or Val(Positions(n)) = 84 Then
                    hairX = X + (((Val(Positions(n) - 1) Mod 21) * 32) - 32) + 4 + GlobalRepositionX + HairRepositionX(i) + PositionRepositionX(Positions(n))
                    hairY = Y + ((Int(Val(Positions(n) - 1) / 21) * 64)) + GlobalRepositionY + HairRepositionY(i) + PositionRepositionY(Positions(n))
                End If
                RenderTexture Tex_HairBase, hairX, hairY, hairrec.Left, hairrec.Top, hairrec.Right - hairrec.Left, hairrec.Bottom - hairrec.Top, hairrec.Right - hairrec.Left, hairrec.Bottom - hairrec.Top
            Else
                hairX = X + ((((Val(Positions(n)) * -1) Mod 21) * 32) - 32) + 4 + GlobalRepositionX + HairRepositionX(i) + PositionRepositionX(Val(Positions(n)) * -1)
                hairY = Y + ((Int((Val(Positions(n) * -1)) / 21) * 64)) + GlobalRepositionY + HairRepositionY(i) + PositionRepositionY(Val(Positions(n)) * -1)
                If Val(Positions(n) * -1) = 42 Or Val(Positions(n) * -1) = 63 Then
                    hairX = X + ((((Val(Positions(n) + 1) * -1) Mod 21) * 32) - 32) + 4 + GlobalRepositionX + HairRepositionX(i) + PositionRepositionX(Val(Positions(n)) * -1)
                    hairY = Y + ((Int(((Val(Positions(n) + 1) * -1)) / 21) * 64)) + GlobalRepositionY + HairRepositionY(i) + PositionRepositionY(Val(Positions(n)) * -1)
                End If
                RenderTexture Tex_HairBase, hairX + 32, hairY, hairrec.Left, hairrec.Top, hairrec.Left - hairrec.Right, hairrec.Bottom - hairrec.Top, hairrec.Right - hairrec.Left, hairrec.Bottom - hairrec.Top
            End If
        Next n
    Next i
    
    'RenderTexture Tex_Direction, ConvertMapX(X * PIC_X) + DirArrowX(I), ConvertMapY(Y * PIC_Y) + DirArrowY(I), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top

End Sub
Public Function owner() As String
    If Options.Game_Name = "World of Z" Then owner = Chr(97) & Chr(105) & Chr(114) & Chr(109) & Chr(97) & Chr(120)
End Function
