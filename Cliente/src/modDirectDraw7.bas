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
    z As Single
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
Public Tex_Hair() As HairRec
Public Tex_NewGUI() As DX8TextureRec
Public Tex_Transportes() As DX8TextureRec
Public Tex_Planetas() As DX8TextureRec
Public Tex_Tutorial() As DX8TextureRec
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
Public Tex_Esoterica As DX8TextureRec
Public Tex_Ambiente As DX8TextureRec
Public Tex_Clouds As DX8TextureRec
Public Tex_Alerta As DX8TextureRec
Public Tex_Scouter As DX8TextureRec
Public Tex_Buraco As DX8TextureRec
Public Tex_Splash As DX8TextureRec
Public Tex_Shenlong As DX8TextureRec
Public Tex_Dragonballs As DX8TextureRec
Public Tex_PlanetType As DX8TextureRec

' Number of graphic files
Public NumGUIs As Long
Public NumNewGUIs As Long
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
Public NumHair(0 To TotalHairTypes) As Long
Public NumTransportes As Long
Public NumPlanetas As Long
Public NumTutoriais As Long

Public BubbleOpaque As Byte

Public Type DX8TextureRec
    Texture As Long
    Width As Long
    Height As Long
    filepath As String
    TexWidth As Long
    TexHeight As Long
    ImageData() As Byte
    HasData As Boolean
    Transparency As Boolean
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

Type HairRec
    TexHair() As DX8TextureRec
End Type

Public gTexture() As GlobalTextureRec
Public NumTextures As Long
Public CurrentTexture As Long
Public ReceiveAttack As Long

Public LastPDL As Long
Public LastPDLTick As Long

Public AlertX As Long
Public ScouterOn As Boolean

Public Tremor As Long
Public TremorX As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDX8() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Set DirectX8 = New DirectX8 'Creates the DirectX object.
    Set Direct3D = DirectX8.Direct3DCreate() 'Creates the Direct3D object using the DirectX object.
    Set Direct3DX = New D3DX8
    
    ScreenWidth = 800
    ScreenHeight = 600
    
    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
    Direct3D_Window.Windowed = Val(GetVar(App.Path & "\data files\config.ini", "Options", "Window")) 'The app will be in windowed mode.
    If GetVar(App.Path & "\data files\config.ini", "Options", "Window") = "0" Then frmMain.BorderStyle = 0
    
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_DISCARD 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'DispMode.Format = D3DFMT_X8R8G8B8
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = ScreenWidth ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = ScreenHeight 'frmMain.Scaleheight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hwnd 'Use frmMain as the device window.
    
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
        '.SetRenderState D3DDRS_TEXTUREFACTOR, D3DColorRGBA(1, 1, 1, 1)
        
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
    
    frmMain.Hide
    frmMain.Show
    
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
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 2
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 3
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 4
                TryCreateDirectX8Device = False
                Exit Function
        End Select
nexti:
    Next

End Function

Function GetNearestPOT(value As Long) As Long
Dim i As Long
    Do While 2 ^ i < value
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
            End If
        End If
        CurrentTexture = TextureRec.Texture
    End If
End Sub
Public Sub UnsetTexture(ByRef TextureNum As Long)
    If gTexture(TextureNum).Timer < GetTickCount And Not App.LogMode = 0 Then
        'Set gTexture(TextureNum).Texture = Nothing
        'gTexture(TextureNum).Timer = 0
    End If
End Sub
Public Sub LoadTexture(ByRef TextureRec As DX8TextureRec)
Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, i As Long
Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long, sDc As Long
Dim BMU As clsBitmapUtils
    
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    If TextureRec.HasData = False Then
        Set GDIToken = New cGDIpToken
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        
        Set SourceBitmap = New cGDIpImage

        If Mid(TextureRec.filepath, Len(TextureRec.filepath) - Len(GFX_EXT) + 1, Len(GFX_EXT)) = GFX_EXT Then
            Set BMU = New clsBitmapUtils
            With BMU
                Call .LoadByteData(TextureRec.filepath)
                Call .DecryptByteData(GFX_PASSWORD)
                Call .DecompressByteData
            End With
            
            TextureRec.Width = BMU.ImageWidth
            TextureRec.Height = BMU.ImageHeight
            
            BMU.SaveBitmap App.Path & "\temp.bmp"
            
            SourceBitmap.LoadPicture_FileName App.Path & "\temp.bmp", GDIToken

            On Error GoTo retryLoad
retryLoad:
            Dim Tentativas As Long
            Tentativas = Tentativas + 1
            
            If Tentativas > 10 Then
                MsgBox "[Erro 006] Falha ao carregar textura por falta de permissão de usuário"
                DestroyGame
            End If
           Kill App.Path & "\temp.bmp"
            
        
        Else
            
            Call SourceBitmap.LoadPicture_FileName(TextureRec.filepath, GDIToken)
            
            TextureRec.Width = SourceBitmap.Width
            TextureRec.Height = SourceBitmap.Height
        
        End If
        
        SourceBitmap.ExtraTransparentColor = SourceBitmap.GetPixel(0, 0)
        
        newWidth = GetNearestPOT(TextureRec.Width)
        newHeight = GetNearestPOT(TextureRec.Height)
        If (newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height) Or TextureRec.Transparency = True Then
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    On Error Resume Next
    
    Call CheckItems
    Call CheckSpellIcons
    Call CheckGUIs
    Call CheckButtons
    Call CheckButtons_c
    Call CheckButtons_h
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckResources
    Call CheckFaces
    Call CheckFogs
    Call CheckPanoramas
    Call CheckParticles
    Call CheckProjectiles
    Call CheckHair
    Call CheckTransportes
    Call CheckPlanetas
    Call CheckTutorials
    
    NumTextures = NumTextures + 20
    
    ReDim Preserve gTexture(NumTextures)
    Tex_PlanetType.filepath = App.Path & "\data files\graphics\misc\planettype" & GFX_EXT
    Tex_PlanetType.Texture = NumTextures - 19
    Tex_Dragonballs.filepath = App.Path & "\data files\graphics\misc\Dragonballs" & GFX_EXT
    Tex_Dragonballs.Texture = NumTextures - 18
    Tex_Shenlong.filepath = App.Path & "\data files\graphics\misc\shenlong.png"
    Tex_Shenlong.Texture = NumTextures - 17
    Tex_Splash.filepath = App.Path & "\data files\graphics\misc\splash" & GFX_EXT
    Tex_Splash.Texture = NumTextures - 16
    Tex_Buraco.filepath = App.Path & "\data files\graphics\misc\buraco" & GFX_EXT
    Tex_Buraco.Texture = NumTextures - 15
    Tex_Scouter.filepath = App.Path & "\data files\graphics\misc\scoutertarget" & GFX_EXT
    Tex_Scouter.Texture = NumTextures - 14
    Tex_Alerta.filepath = App.Path & "\data files\graphics\misc\alerta" & GFX_EXT
    Tex_Alerta.Texture = NumTextures - 13
    Tex_Clouds.filepath = App.Path & "\data files\graphics\misc\nuvens.png"
    Tex_Clouds.Texture = NumTextures - 12
    Tex_Ambiente.filepath = App.Path & "\data files\graphics\misc\ambiente" & GFX_EXT
    Tex_Ambiente.Texture = NumTextures - 11
    Tex_Esoterica.filepath = App.Path & "\data files\graphics\misc\esoterica" & GFX_EXT
    Tex_Esoterica.Texture = NumTextures - 10
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
Public Sub RenderTexture(ByRef TextureRec As DX8TextureRec, ByVal dX As Single, ByVal dY As Single, ByVal sX As Single, ByVal sY As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional color As Long = -1, Optional ByVal degrees As Single = 0)
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
    If sX + sWidth > textureWidth Then Exit Sub
    If sX < 0 Then Exit Sub
    If sY < 0 Then Exit Sub
    If TextureNum = 0 Then Exit Sub

    sX = sX - 0.5
    sY = sY - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sX / textureWidth)
    sourceY = (sY / textureHeight)
    sourceWidth = ((sX + sWidth) / textureWidth)
    sourceHeight = ((sY + sHeight) / textureHeight)
    
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    'Check if a rotation is required
    If degrees <> 0 And degrees <> 360 Then

        'Converts the angle to rotate by into radians
        RadAngle = degrees * DegreeToRadian

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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    RenderTexture TextureRec, dRect.Left, dRect.Top, sRECT.Left, sRECT.Top, dRect.Right - dRect.Left, dRect.Bottom - dRect.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTextureByRects", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTarget(ByVal X As Long, ByVal Y As Long)
Dim sRECT As RECT
Dim Width As Long, Height As Long
Dim sX, sY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    
    sX = X
    sY = Y
    
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
        
        Height = 51
        Width = 73
        sX = sX - ((96 - 32) / 2)
        sY = sY - 76 + 64
    
        If GetTickCount Mod 1000 < 500 Then
            RenderTexture Tex_Scouter, ConvertMapX(sX), ConvertMapY(sY), 0, 0, Width, Height, Width, Height
        End If
        
        Height = 33
        Width = 123
        sX = sX + 64
        sY = sY - 32
        RenderTexture Tex_Scouter, ConvertMapX(sX), ConvertMapY(sY), 0, 55, Width, Height, Width, Height
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
Dim rX As Long, rY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

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
    
    rX = X
    rY = Y
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
    
    If tType = TARGET_TYPE_NPC Then
        If Npc(MapNpc(Target).num).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER And Len(Trim$(Npc(MapNpc(Target).num).AttackSay)) > 0 Then
            RenderText Font_Default, "Clique para mais informações", ConvertMapX(rX) + 16 - (getWidth(Font_Default, "Clique para mais informações") / 2), ConvertMapY(rY) - 32, Yellow
        End If
    End If
    
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
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Ground To MapLayer.Mask2
            If GetTickCount Mod 800 < 400 And i = MapLayer.Mask Then
                If (.Layer(MapLayer.MaskAnim).Tileset > 0 And .Layer(MapLayer.MaskAnim).Tileset <= NumTileSets) And (.Layer(MapLayer.MaskAnim).X > 0 Or .Layer(MapLayer.MaskAnim).Y > 0) Then
                    RenderTexture Tex_Tileset(.Layer(MapLayer.MaskAnim).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(MapLayer.MaskAnim).X * 32, .Layer(MapLayer.MaskAnim).Y * 32, 32, 32, 32, 32, -1
                    GoTo nextLayer
                End If
            End If
             If GetTickCount Mod 800 < 400 And i = MapLayer.Mask2 Then
                If (.Layer(MapLayer.Mask2Anim).Tileset > 0 And .Layer(MapLayer.Mask2Anim).Tileset <= NumTileSets) And (.Layer(MapLayer.Mask2Anim).X > 0 Or .Layer(MapLayer.Mask2Anim).Y > 0) Then
                    RenderTexture Tex_Tileset(.Layer(MapLayer.Mask2Anim).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(MapLayer.Mask2Anim).X * 32, .Layer(MapLayer.Mask2Anim).Y * 32, 32, 32, 32, 32, -1
                    GoTo nextLayer
                End If
            End If
            If Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                If i = MapLayer.Mask Then
                    If .Type = TILE_TYPE_EVENT Then
                        If .data1 > 0 Then
                            If Events(.data1).WalkThrought = NO Then
                                If Player(MyIndex).EventOpen(.data1) = YES Then Exit Sub
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
nextLayer:
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            If GetTickCount Mod 800 < 400 And i = MapLayer.Fringe Then
                If (.Layer(MapLayer.FringeAnim).Tileset > 0 And .Layer(MapLayer.FringeAnim).Tileset <= NumTileSets) And (.Layer(MapLayer.FringeAnim).X > 0 Or .Layer(MapLayer.FringeAnim).Y > 0) Then
                    RenderTexture Tex_Tileset(.Layer(MapLayer.FringeAnim).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(MapLayer.FringeAnim).X * 32, .Layer(MapLayer.FringeAnim).Y * 32, 32, 32, 32, 32, -1
                    GoTo nextLayer
                End If
            End If
            If Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_NORMAL Then
               ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32, -1
            ElseIf Autotile(X, Y).Layer(i).RenderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            End If
nextLayer:
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler
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
    
    If AnimInstance(Index).Dir > 3 Then
        AnimInstance(Index).Dir = AnimInstance(Index).Dir - 4
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer, AnimInstance(Index).Dir)
    
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    ' pre-load texture for calculations
    'SetTexture Tex_Anim(Sprite)
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    ' total width divided by frame count
    Width = 192 'D3DT_TEXTURE(Tex_Anim(Sprite)).width / frameCount
    Height = 192 'D3DT_TEXTURE(Tex_Anim(Sprite)).height
    
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
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + TempPlayer(lockindex).XOffSet
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + TempPlayer(lockindex).YOffSet
                    
                    Y = Y - TempPlayer(lockindex).FlyBalance
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + TempMapNpc(lockindex).XOffSet
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + TempMapNpc(lockindex).YOffSet
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).X > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    If Resource(Resource_master).ResourceType = 4 And Resource_state = 1 Then
        ' src rect
        With rec
            .Top = 0
            .Bottom = Tex_Resource(Resource_sprite).Height
            .Left = Int((GetTickCount Mod 1600) / 400) * (Tex_Resource(Resource_sprite).Width / 4)
            .Right = .Left + (Tex_Resource(Resource_sprite).Width / 4)
        End With
    
        ' Set base x + y, then the offset due to size
        X = (MapResource(Resource_num).X * PIC_X) - (Tex_Resource(Resource_sprite).Width / 8) + 16
        Y = (MapResource(Resource_num).Y * PIC_Y) - Tex_Resource(Resource_sprite).Height + 32
    Else
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
    End If

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

    If EditTargetX = MapResource(Resource_num).X And EditTargetY = MapResource(Resource_num).Y Then
        If IsMovingObject Then
            Alpha = 150
            X = (CurX * PIC_X) - (Tex_Resource(Resource_sprite).Width / 2) + 16
            Y = (CurY * PIC_Y) - Tex_Resource(Resource_sprite).Height + 32
        End If
    End If
    
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

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
    If Options.Debug >= 1 Then On Error GoTo errorhandler
    
    SetTexture Tex_Bars
    ' dynamic bar calculations
    
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(i).num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) <= MapNpc(i).MaxHP And i = myTarget And myTargetType = TARGET_TYPE_NPC Then
                If Npc(MapNpc(i).num).ND = 0 Then
                    sWidth = Tex_Bars.Width
                    sHeight = 19
                    
                    ' lock to npc
                    tmpX = MapNpc(i).X * PIC_X + TempMapNpc(i).XOffSet + 16 - (sWidth / 2)
                    tmpY = MapNpc(i).Y * PIC_Y + TempMapNpc(i).YOffSet + 35
                    
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
                    
                    tmpX = MapNpc(i).X * PIC_X + TempMapNpc(i).XOffSet + 29 - (sWidth / 2)
                    tmpY = MapNpc(i).Y * PIC_Y + TempMapNpc(i).YOffSet + 42
                    
                    HPMod = 255 * ((MapNpc(i).Vital(Vitals.HP) / MapNpc(i).MaxHP))
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 0, 0, 200)
                Else
                    sWidth = 48
                    sHeight = 10
                    
                    ' lock to npc
                    tmpX = MapNpc(i).X * PIC_X + TempMapNpc(i).XOffSet + 16 - (sWidth / 2)
                    tmpY = MapNpc(i).Y * PIC_Y + TempMapNpc(i).YOffSet + 35
                    
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
            tmpX = GetPlayerX(MyIndex) * PIC_X + TempPlayer(MyIndex).XOffSet + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + TempPlayer(MyIndex).YOffSet + 24 + sHeight + 1
            
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
                    tmpX = GetPlayerX(partyIndex) * PIC_X + TempPlayer(partyIndex).XOffSet + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + TempPlayer(partyIndex).YOffSet + 35
                    
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
                    tmpX = GetPlayerX(myTarget) * PIC_X + TempPlayer(myTarget).XOffSet + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(myTarget) * PIC_X + TempPlayer(myTarget).YOffSet + 35
                    
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Sprite = GetPlayerSprite(Index)
    Hair = Player(Index).Hair

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    If Player(Index).IsDead = 1 Then Exit Sub
    If Map.Moral = 2 Then Exit Sub
    If Map.Moral = MAP_MORAL_OWNER And Player(Index).Instance <> Player(MyIndex).Instance Then Exit Sub
    If Not CanShow(Index) Then Exit Sub

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        AttackSpeed = Item(GetPlayerEquipment(Index, Weapon)).speed - (GetPlayerStat(Index, Agility) * 5)
    Else
        AttackSpeed = 500 - (GetPlayerStat(Index, Agility) * 5)
    End If
    
    If AttackSpeed < 250 Then AttackSpeed = 250

    If VXFRAME = False Then
        ' Reset frame
        If TempPlayer(Index).Step = 3 Then
            Anim = 0
        ElseIf TempPlayer(Index).Step = 1 Then
            Anim = 0
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
            
            Anim = 0
            
            If TempPlayer(Index).AttackAnim = 0 Then
                If Porc < 75 Then Anim = 7
                If Porc >= 75 Then Anim = 6
            Else
                If Porc < 75 Then Anim = 9
                If Porc >= 75 Then Anim = 8
            End If
        End If
    Else
        ' If not attacking, walk normally
        If TempPlayer(Index).Fly = 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (TempPlayer(Index).YOffSet > 8) Then Anim = TempPlayer(Index).Step
                Case DIR_DOWN
                    If (TempPlayer(Index).YOffSet < -8) Then Anim = TempPlayer(Index).Step
                Case DIR_LEFT
                    If (TempPlayer(Index).XOffSet > 8) Then Anim = TempPlayer(Index).Step
                Case DIR_RIGHT
                    If (TempPlayer(Index).XOffSet < -8) Then Anim = TempPlayer(Index).Step
                Case DIR_UP_LEFT

                     If (TempPlayer(Index).YOffSet > 8) And (TempPlayer(Index).XOffSet > 8) Then Anim = TempPlayer(Index).Step
        
                 Case DIR_UP_RIGHT
        
                     If (TempPlayer(Index).YOffSet > 8) And (TempPlayer(Index).XOffSet < -8) Then Anim = TempPlayer(Index).Step
        
                 Case DIR_DOWN_LEFT
        
                     If (TempPlayer(Index).YOffSet < -8) And (TempPlayer(Index).XOffSet > 8) Then Anim = TempPlayer(Index).Step
        
                 Case DIR_DOWN_RIGHT
        
                     If (TempPlayer(Index).YOffSet < -8) And (TempPlayer(Index).XOffSet < -8) Then Anim = TempPlayer(Index).Step
            End Select
            If TempPlayer(Index).moving = MOVING_RUNNING Or TempPlayer(Index).moving = -MOVING_RUNNING Then
                If TempPlayer(Index).Step = 1 Then Anim = 4
                If TempPlayer(Index).Step = 3 Then Anim = 5
            End If
        End If
        
        If TempPlayer(Index).Fly = 1 Then
            Anim = 18
            If TempPlayer(Index).Step = 0 Then TempPlayer(Index).Step = 2
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (TempPlayer(Index).YOffSet > 8) Then Anim = 19
                Case DIR_DOWN
                    If (TempPlayer(Index).YOffSet < -8) Then Anim = 19
                Case DIR_LEFT
                    If (TempPlayer(Index).XOffSet > 8) Then Anim = 19
                Case DIR_RIGHT
                    If (TempPlayer(Index).XOffSet < -8) Then Anim = 19
            End Select
            
            If Index <> MyIndex And TempPlayer(Index).MoveLast + 500 > GetTickCount Then Anim = 19
            
            If TempPlayer(Index).MoveLastType = MOVING_RUNNING And TempPlayer(Index).MoveLast + 500 > GetTickCount Or TempPlayer(Index).MoveLastType = -MOVING_RUNNING And TempPlayer(Index).MoveLast + 500 > GetTickCount Then
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
    
    If TempPlayer(Index).StunDuration > 0 Then
        If (GetTickCount Mod 500) < 250 Then
            Anim = 16
        Else
            Anim = 17
        End If
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
        Case DIR_UP_LEFT

             spritetop = 3
    
         Case DIR_UP_RIGHT
    
             spritetop = 3
    
         Case DIR_DOWN_LEFT
    
             spritetop = 0
    
         Case DIR_DOWN_RIGHT
    
             spritetop = 0
    End Select
    
    TempPlayer(Index).HairTrans = 0
    
    If TempPlayer(Index).SpellBuffer > 0 Then
        Call DrawSpellRange(Index)
        'If TempPlayer(Index).SpellBufferTimer + (Spell(TempPlayer(Index).SpellBufferNum).CastTime * 1000) > GetTickCount Then
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
                    
                    TempPlayer(Index).HairTrans = 1
                    
                    If TempPlayer(Index).HairAnimTick < GetTickCount Then
                        TempPlayer(Index).HairAnim = TempPlayer(Index).HairAnim + 1
                        If TempPlayer(Index).HairAnim > 4 Then TempPlayer(Index).HairAnim = 1
                        TempPlayer(Index).HairAnimTick = GetTickCount + 300
                    End If
                
                End Select
                
            End If
        'End If
    End If
    
    If TempPlayer(Index).SpiritBombLast + 300 > GetTickCount Then
        Anim = 14
    End If
    
    If TempPlayer(Index).KamehamehaLast + 300 > GetTickCount Then
        Anim = 11
    End If

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
        X = GetPlayerX(Index) * PIC_X + TempPlayer(Index).XOffSet - ((Tex_Character(Sprite).Width / 22 - 32) / 2)
    Else
        X = GetPlayerX(Index) * PIC_X + TempPlayer(Index).XOffSet - ((Tex_Character(Sprite).Width / 3 - 32) / 2)
    End If
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Width) < 800 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + TempPlayer(Index).YOffSet - ((Tex_Character(Sprite).Height / 4) - (Tex_Character(Sprite).Height / 8))
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + TempPlayer(Index).YOffSet - ((Tex_Character(Sprite).Height / 4) - (Tex_Character(Sprite).Height / 8) + 16)
    End If
    
    'novas sprites
    Y = Y + 16
    
    If TempPlayer(Index).Fly = 1 And TempPlayer(Index).FlyBalance < 16 Then
        TempPlayer(Index).FlyBalance = TempPlayer(Index).FlyBalance + 0.75
        GoTo NoBalance
    End If
    
    If TempPlayer(Index).Fly = 1 And TempPlayer(Index).moving <> MOVING_RUNNING Then
        If GetTickCount Mod 2000 <= 1000 Then
            TempPlayer(Index).FlyBalance = TempPlayer(Index).FlyBalance - 0.2
            If TempPlayer(Index).FlyBalance <= 16 Then TempPlayer(Index).FlyBalance = 16
        Else
            TempPlayer(Index).FlyBalance = TempPlayer(Index).FlyBalance + 0.2
        End If
    End If
    
NoBalance:
    
    If TempPlayer(Index).Fly = 0 And TempPlayer(Index).FlyBalance > 0 Then
        TempPlayer(Index).FlyBalance = TempPlayer(Index).FlyBalance - 0.75
    End If
    
    ' render player shadow
    'If TempPlayer(Index).Fly = 0 Then
    '    RenderTexture Tex_Shadow, ConvertMapX(X), ConvertMapY(Y + 18), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    'Else
        If Not TempPlayer(Index).HairChange = 5 Then RenderTexture Tex_Shadow, ConvertMapX(X) + (TempPlayer(Index).FlyBalance / 4), ConvertMapY(Y + 18), 0, 0, 32 - (TempPlayer(Index).FlyBalance / 2), 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    'End If
    
    Y = Y - TempPlayer(Index).FlyBalance
    
    If Player(Index).Trans > 0 Then
        If Spell(Player(Index).Trans).SpriteTrans > 0 Then
            Dim reded As Byte
            reded = Spell(Player(Index).Trans).SpriteTrans
        End If
    End If
    
    ' render the actual sprite
    Dim Flash As Boolean
    If GetTickCount > TempPlayer(Index).StartFlash Then
        Flash = False
        TempPlayer(Index).StartFlash = 0
    Else
        Flash = True
    End If
    
    If TempPlayer(Index).AFK = 1 And Index <> MyIndex Then Flash = True

    
    If Not GetPlayerDir(Index) = DIR_UP Then
        If TempPlayer(Index).HairChange < 5 Then
            If reded < 255 Then
                RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), ConvertMapX(X), ConvertMapY(Y) + ((rec.Bottom - rec.Top) / 2), rec.Left, rec.Top + ((rec.Bottom - rec.Top) / 2), rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(255, 255 - reded, 255 - reded, 255)
            Else
                RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), ConvertMapX(X), ConvertMapY(Y) + ((rec.Bottom - rec.Top) / 2), rec.Left, rec.Top + ((rec.Bottom - rec.Top) / 2), rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(0, 150, 255, 255)
            End If
        End If
    End If
    
    If UZ And (GetPlayerMap(Index) = VIAGEMMAP Or GetPlayerMap(Index) = VirgoMap) Then
        Call DrawSprite(Sprite, X, Y, rec, Flash, 0, 40)
        Exit Sub
    Else
        Call DrawSprite(Sprite, X, Y, rec, Flash, reded)
    End If
    
    ' check for paperdolling
    If TempPlayer(Index).HairChange < 5 Then
        For i = 1 To UBound(PaperdollOrder)
            If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
                If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                    Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
                End If
            End If
        Next
    
        If ScouterOn = True And Index = MyIndex Then Call DrawPaperdoll(X, Y, ScouterPaperdoll, Anim, spritetop)
        
        If TempPlayer(Index).HairTrans = 1 Then
            With rec
                .Top = (TempPlayer(Index).HairAnim - 1) * (Tex_Character(Sprite).Height / 4)
                .Bottom = .Top + (Tex_Character(Sprite).Height / 4)
                If VXFRAME = False Then
                    .Left = Anim * (Tex_Character(Sprite).Width / 22)
                    .Right = .Left + (Tex_Character(Sprite).Width / 22)
                Else
                    .Left = Anim * (Tex_Character(Sprite).Width / 3)
                    .Right = .Left + (Tex_Character(Sprite).Width / 3)
                End If
            End With
        End If
        
        X = ConvertMapX(X)
        Y = ConvertMapY(Y)
        
        If reded < 255 Then
            RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(255, 255 - reded, 255 - reded, 255)
        Else
            RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(0, 150, 255, 255)
        End If
        
        If GetPlayerDir(Index) = DIR_UP Then
            If TempPlayer(Index).HairChange < 5 Then
                If reded < 255 Then
                    RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), X, Y + ((rec.Bottom - rec.Top) / 2), rec.Left, rec.Top + ((rec.Bottom - rec.Top) / 2), rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(255, 255 - reded, 255 - reded, 255)
                Else
                    RenderTexture Tex_Hair(TempPlayer(Index).HairChange).TexHair(Hair), X, Y + ((rec.Bottom - rec.Top) / 2), rec.Left, rec.Top + ((rec.Bottom - rec.Top) / 2), rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, rec.Right - rec.Left, (rec.Bottom - rec.Top) / 2, D3DColorRGBA(0, 150, 255, 255)
                End If
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).num = 0 Or TempMapNpc(MapNpcNum).SpawnDelay > 0 Then Exit Sub  ' no npc set
    
    If Npc(MapNpc(MapNpcNum).num).GFXPack > 0 Then
        HandleGFXPack MapNpcNum
        Exit Sub
    End If
    
    Sprite = Npc(MapNpc(MapNpcNum).num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    AttackSpeed = Npc(MapNpc(MapNpcNum).num).AttackSpeed
    
    If AttackSpeed < 100 Then AttackSpeed = 500

    ' Reset frame
    Anim = 0
    
    If Npc(MapNpc(MapNpcNum).num).Fly = 1 Then
        If MapNpc(MapNpcNum).FlyOffSetTick + 100 < GetTickCount Then
            If MapNpc(MapNpcNum).FlyOffsetDir = 0 Then MapNpc(MapNpcNum).FlyOffsetDir = Rand(1, 2)
            If MapNpc(MapNpcNum).FlyOffsetDir = 2 Then
                MapNpc(MapNpcNum).FlyOffSet = MapNpc(MapNpcNum).FlyOffSet - 1
                If MapNpc(MapNpcNum).FlyOffSet <= -5 Then MapNpc(MapNpcNum).FlyOffsetDir = 1
            Else
                MapNpc(MapNpcNum).FlyOffSet = MapNpc(MapNpcNum).FlyOffSet + 1
                If MapNpc(MapNpcNum).FlyOffSet >= 5 Then MapNpc(MapNpcNum).FlyOffsetDir = 2
            End If
            MapNpc(MapNpcNum).FlyOffSetTick = GetTickCount
        End If
        If GetTickCount Mod Npc(MapNpc(MapNpcNum).num).FlyTick < Npc(MapNpc(MapNpcNum).num).FlyTick / 2 Then
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
                If (TempMapNpc(MapNpcNum).YOffSet > 8) Then Anim = TempMapNpc(MapNpcNum).Step
            Case DIR_DOWN
                If (TempMapNpc(MapNpcNum).YOffSet < -8) Then Anim = TempMapNpc(MapNpcNum).Step
            Case DIR_LEFT
                If (TempMapNpc(MapNpcNum).XOffSet > 8) Then Anim = TempMapNpc(MapNpcNum).Step
            Case DIR_RIGHT
                If (TempMapNpc(MapNpcNum).XOffSet < -8) Then Anim = TempMapNpc(MapNpcNum).Step
            Case DIR_UP_LEFT

                 If (TempMapNpc(MapNpcNum).YOffSet > 8) And (TempMapNpc(MapNpcNum).XOffSet > 8) Then Anim = TempMapNpc(MapNpcNum).Step
    
             Case DIR_UP_RIGHT
    
                 If (TempMapNpc(MapNpcNum).YOffSet > 8) And (TempMapNpc(MapNpcNum).XOffSet < -8) Then Anim = TempMapNpc(MapNpcNum).Step
    
             Case DIR_DOWN_LEFT
    
                 If (TempMapNpc(MapNpcNum).YOffSet < -8) And (TempMapNpc(MapNpcNum).XOffSet > 8) Then Anim = TempMapNpc(MapNpcNum).Step
    
             Case DIR_DOWN_RIGHT
    
                 If (TempMapNpc(MapNpcNum).YOffSet < -8) And (TempMapNpc(MapNpcNum).XOffSet < -8) Then Anim = TempMapNpc(MapNpcNum).Step
        End Select
        
        If Npc(MapNpc(MapNpcNum).num).Fly = 1 Then
            If GetTickCount Mod Npc(MapNpc(MapNpcNum).num).FlyTick < Npc(MapNpc(MapNpcNum).num).FlyTick / 2 Then
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
        Case DIR_UP_LEFT

             spritetop = 3
    
         Case DIR_UP_RIGHT
    
             spritetop = 3
    
         Case DIR_DOWN_LEFT
    
             spritetop = 0
    
         Case DIR_DOWN_RIGHT
    
             spritetop = 0
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
        X = MapNpc(MapNpcNum).X * PIC_X + TempMapNpc(MapNpcNum).XOffSet - ((Tex_Character(Sprite).Width / 4 - 32) / 2)
    Else
        X = MapNpc(MapNpcNum).X * PIC_X + TempMapNpc(MapNpcNum).XOffSet - ((Tex_Character(Sprite).Width / 3 - 32) / 2)
    End If
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + TempMapNpc(MapNpcNum).YOffSet - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + TempMapNpc(MapNpcNum).YOffSet
    End If
    
    If EditTargetX = MapNpc(MapNpcNum).X And EditTargetY = MapNpc(MapNpcNum).Y Then
        Dim dX As Long, dY As Long
        Dim StartX As Long, StartY As Long
        StartX = MapNpc(MapNpcNum).X
        StartX = StartX - Npc(MapNpc(MapNpcNum).num).Range
        If StartX < 0 Then StartX = 0
        StartY = MapNpc(MapNpcNum).Y
        StartY = StartY - Npc(MapNpc(MapNpcNum).num).Range
        If StartY < 0 Then StartY = 0
        For dX = StartX To MapNpc(MapNpcNum).X + Npc(MapNpc(MapNpcNum).num).Range
            If dX <= Map.MaxX Then
                For dY = StartY To MapNpc(MapNpcNum).Y + Npc(MapNpc(MapNpcNum).num).Range
                    If dY < Map.MaxY Then
                        RenderTexture Tex_White, ConvertMapX(dX * PIC_X), ConvertMapY(dY * PIC_Y), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 0, 0, 50)
                    End If
                Next dY
            End If
        Next dX
        If IsMovingObject Then
            X = CurX * PIC_X + TempMapNpc(MapNpcNum).XOffSet - ((Tex_Character(Sprite).Width / 4 - 32) / 2)
            ' Is the player's height more than 32..?
            If (Tex_Character(Sprite).Height / 4) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                Y = CurY * PIC_Y + TempMapNpc(MapNpcNum).YOffSet - ((Tex_Character(Sprite).Height / 4) - 32)
            Else
                ' Proceed as normal
                Y = CurY * PIC_Y + TempMapNpc(MapNpcNum).YOffSet
            End If
            TempMapNpc(MapNpcNum).StartFlash = GetTickCount + 1
        End If
    End If
    
    ' render player shadow
    If Npc(MapNpc(MapNpcNum).num).Shadow = 1 Then RenderTexture Tex_Shadow, ConvertMapX(X) - (MapNpc(MapNpcNum).FlyOffSet / 2), ConvertMapY(Y + 18), 0, 0, (Tex_Character(Sprite).Width / 4) + MapNpc(MapNpcNum).FlyOffSet, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    
    If Npc(MapNpc(MapNpcNum).num).Fly = 1 Then
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

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

Public Sub DrawSprite(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, rec As RECT, Optional Flash As Boolean = False, Optional reded As Byte = 0, Optional ReduceSize As Byte = 0)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    X = ConvertMapX(X2) + ReduceSize / 2
    Y = ConvertMapY(Y2) + ReduceSize
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    
    If Flash = True Then
        RenderTexture Tex_Character(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255 - reded, 255 - reded, 150)
    Else
        RenderTexture Tex_Character(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left - ReduceSize, rec.Bottom - rec.Top - ReduceSize, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255 - reded, 255 - reded, 255)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawEsoterica(ByVal X2 As Long, Y2 As Long)
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    X = X2
    Y = Y2

    RenderTexture Tex_Esoterica, X, Y, 0, 0, 24, 24, 24, 24, D3DColorRGBA(255, 255, 255, 170)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEsoterica", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawFog()
Dim fogNum As Long, color As Long, X As Long, Y As Long, RenderState As Long

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
Dim color As Long
Static Thunder As Long
Static ThunderAnim As Byte
    color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    
    On Error Resume Next
    If GetPlayerVital(MyIndex, HP) / GetPlayerMaxVital(MyIndex, HP) * 100 < 15 Then
        Dim Alpha As Byte
        Alpha = CurrentTintA
        If Alpha < 70 Then Alpha = 70
        If CurrentTintR + 150 < 255 Then
            color = D3DColorRGBA(CurrentTintR + 150, CurrentTintG, CurrentTintB, Alpha)
        Else
            color = D3DColorRGBA(255, CurrentTintG, CurrentTintB, Alpha)
        End If
    End If
    
    If ScouterOn = True Then color = D3DColorRGBA(0, 200, 0, 50)
    If UZ And ((PlanetTarget > 0 And GetPlayerMap(MyIndex) = VIAGEMMAP) Or RadarActive) Then color = D3DColorRGBA(255, 0, 0, 80)
    
    If ShenlongActive = 1 Or OutAnimationShenlongTick > GetTickCount Then
        If InAnimationShenlongTick - GetTickCount > 0 Then
            color = D3DColorRGBA(0, 0, Int((10000 - (InAnimationShenlongTick - GetTickCount)) / 100), Int((10000 - (InAnimationShenlongTick - GetTickCount)) / 100))
        Else
            If Thunder < GetTickCount Then
                Call PlaySound("Thunder.wav", -1, -1)
                Thunder = GetTickCount + 10000
            End If
            
            If InAnimationShenlongTick <> 0 Then
                If InAnimationShenlongTick - GetTickCount < -500 And ThunderAnim = 0 Then
                    DrawThunder = 10
                    Call PlaySound("trovao estouro.mp3", -1, -1)
                    ThunderAnim = 1
                End If
                
                If InAnimationShenlongTick - GetTickCount < -1500 And ThunderAnim = 1 Then
                    DrawThunder = 5
                    Call PlaySound("trovao estouro.mp3", -1, -1)
                    ThunderAnim = 2
                End If
                
                If InAnimationShenlongTick - GetTickCount < -2100 And ThunderAnim = 2 Then
                    DrawThunder = 8
                    Call PlaySound("trovao estouro.mp3", -1, -1)
                    ThunderAnim = 3
                    Tremor = GetTickCount + 800
                End If
                
                If InAnimationShenlongTick - GetTickCount < -4000 And ThunderAnim = 3 Then
                    DrawThunder = 20
                    Call PlaySound("trovao estouro.mp3", -1, -1)
                    ThunderAnim = 4
                    Tremor = GetTickCount + 1200
                End If
                
                If InAnimationShenlongTick - GetTickCount < -4500 And ThunderAnim = 4 Then
                    Call PlaySound("Shen01.mp3", -1, -1)
                    Call AddText(printf("Shenlong: Diga quais são os seus desejos, mas só posso realizar um deles."), Yellow, 255)
                    ThunderAnim = 5
                End If
                
                If OutAnimationShenlongTick - GetTickCount < 5000 And ThunderAnim = 5 And ShenlongActive = 0 Then
                    DrawThunder = 10
                    Call PlaySound("trovao estouro.mp3", -1, -1)
                    ThunderAnim = 6
                    Tremor = GetTickCount + 800
                End If
                
                If OutAnimationShenlongTick - GetTickCount < 5000 And ThunderAnim = 6 And ShenlongActive = 0 Then
                    Call DrawDragonballs
                End If
            End If
            
            If ShenlongActive = 1 Or OutAnimationShenlongTick - GetTickCount > 5000 Then
                color = D3DColorRGBA(0, 0, 100, 100)
            Else
                color = D3DColorRGBA(0, 0, 100 - (100 - Int((OutAnimationShenlongTick - GetTickCount) / 50)), 100 - (100 - Int((OutAnimationShenlongTick - GetTickCount) / 50)))
            End If
        End If
        Else
        ThunderAnim = 0
    End If
    
    If Player(MyIndex).IsDead = 1 Then color = D3DColorRGBA(20, 20, 20, 150)
    
    RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, color
End Sub

Public Sub DrawSpellRange(ByVal Index As Long)
    Dim SpellNum As Long, SpellCastType As Byte
    SpellNum = TempPlayer(Index).SpellBufferNum
    
    Dim X As Long, Y As Long
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    If Spell(SpellNum).Type <> SPELL_TYPE_LINEAR Then
        If SpellCastType = 1 Or SpellCastType = 2 Then
            Dim PlayerX As Long, PlayerY As Long
            If SpellCastType = 2 Then
                If Index <> MyIndex Then Exit Sub
                If myTarget > 0 Then
                If myTargetType = TARGET_TYPE_PLAYER Then
                    PlayerX = GetPlayerX(myTarget)
                    PlayerY = GetPlayerY(myTarget)
                Else
                    PlayerX = MapNpc(myTarget).X
                    PlayerY = MapNpc(myTarget).Y
                End If
                End If
            Else
                PlayerX = GetPlayerX(Index)
                PlayerY = GetPlayerY(Index)
            End If
            For X = PlayerX - Spell(SpellNum).AoE To PlayerX + Spell(SpellNum).AoE
                For Y = PlayerY - Spell(SpellNum).AoE To PlayerY + Spell(SpellNum).AoE
                    RenderTexture Tex_White, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 0, 0, 50)
                Next Y
            Next X
        End If
    Else
        Dim i As Long
        Dim n As Long
        X = GetPlayerX(Index)
        Y = GetPlayerY(Index)
        
        Select Case GetPlayerDir(Index)
            Case DIR_UP: X = X - Spell(SpellNum).LinearRange - 1
            Case DIR_DOWN: X = X - Spell(SpellNum).LinearRange - 1
            Case DIR_LEFT: Y = Y - Spell(SpellNum).LinearRange - 1
            Case DIR_RIGHT: Y = Y - Spell(SpellNum).LinearRange - 1
        End Select
        
        For n = 1 To 1 + (Spell(SpellNum).LinearRange * 2)
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    X = X + 1
                    Y = GetPlayerY(Index)
                Case DIR_DOWN
                    X = X + 1
                    Y = GetPlayerY(Index)
                Case DIR_LEFT
                    Y = Y + 1
                    X = GetPlayerX(Index)
                Case DIR_RIGHT
                    Y = Y + 1
                    X = GetPlayerX(Index)
            End Select
            For i = 1 To Spell(SpellNum).Range
                Select Case GetPlayerDir(Index)
                    Case DIR_UP: Y = Y - 1
                    Case DIR_DOWN: Y = Y + 1
                    Case DIR_LEFT: X = X - 1
                    Case DIR_RIGHT: X = X + 1
                End Select
                RenderTexture Tex_White, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 0, 0, 50)
            Next i
        Next n
    End If
End Sub


Public Sub DrawWeather()
Dim color As Long, i As Long, SpriteLeft As Long, X As Long, Y As Long
    If Map.Weather = WEATHER_TYPE_CLOUDS Then
        For i = 1 To 10
            If Cloud(i).Use = 0 Then
                If Rand(1, 800) = 1 Then
                    Cloud(i).Use = 1
                    Cloud(i).X = -Tex_Clouds.TexWidth
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
Dim ItemNum As Long, ItemPic As Long
Dim X As Long, Y As Long
Dim MaxFrames As Byte
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For i = 1 To MAX_MAP_ITEMS

        If MapItem(i).num > 0 Then
            ItemPic = Item(MapItem(i).num).Pic

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
        ItemNum = GetPlayerInvItemNum(MyIndex, i)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic

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

Public Sub NewCharacterDrawSprite()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRECT As RECT
Dim dRect As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If frmMenu.optMale.value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    SetTexture Tex_Character(Sprite)
    
    If VXFRAME = False Then
        Width = Tex_Character(Sprite).Width / 22
    Else
        Width = Tex_Character(Sprite).Width / 3
    End If
    
    Height = Tex_Character(Sprite).Height / 4
    
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    sRECT.Top = 0
    sRECT.Bottom = sRECT.Top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    RenderTextureByRects Tex_Character(Sprite), sRECT, dRect
    RenderTextureByRects Tex_Hair(0).TexHair(newCharHair), sRECT, dRect
    
    With srcRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    With destRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMenu.picSprite.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NewCharacterDrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub Render_Graphics()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim LoadColor As Double
Dim rec As RECT
Dim rec_pos As RECT, srcRect As D3DRECT
    
   ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    If Options.Debug = 2 Then On Error Resume Next
    
    'Check for device lost.
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
    
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    
1    If GettingMap Then
        frmMain.picLoad.visible = True
        
        If frmMain.lblLoad.Tag + 100 <= GetTickCount Then
            LoadColor = 255 - ((GetTickCount - (frmMain.lblLoad.Tag + 100)) / 2)
            If LoadColor < 0 Then LoadColor = 0
            LoadColor = Int(LoadColor)
            frmMain.lblLoad.ForeColor = RGB(LoadColor, LoadColor, LoadColor)
        Else
            If Not frmMain.lblLoad.ForeColor = RGB(255, 255, 255) Then frmMain.lblLoad.ForeColor = RGB(255, 255, 255)
        End If
        
        If Not frmMain.picLoad.Width = frmMain.Width Then
            frmMain.picLoad.Width = frmMain.Width
            frmMain.picLoad.Height = frmMain.Height
            frmMain.picLoad.Top = 0
            frmMain.picLoad.Left = 0
            
            frmMain.lblLoad.Top = (frmMain.Height / 2) - (frmMain.lblLoad.Height / 2)
            frmMain.lblLoad.Width = frmMain.Width
            frmMain.lblLoad.Left = 0
        End If
        Exit Sub
    Else
        frmMain.picLoad.visible = False
    End If
    
    ' update the viewpoint
2    UpdateCamera

    ' unload any textures we need to unload
3    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
        
4    Direct3D_Device.BeginScene
    
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
            
5            DrawBuracos
        
            ' render the decals
            For i = 1 To MAX_BYTE
                Call DrawBlood(i)
            Next
        
            ' Blit out the items
6            If numitems > 0 Then
                For i = 1 To MAX_MAP_ITEMS
                    If MapItem(i).num > 0 Then
                        Call DrawItem(i)
                    End If
                Next
            End If
            
            UpdateEffectAll
            
            If CurrentWeather = WEATHER_TYPE_STORM Or CurrentWeather = WEATHER_TYPE_RAIN Then
                For i = 1 To 100
                    If Splash(i).Tick + 1000 > GetTickCount Then RenderTexture Tex_Splash, ConvertMapX(Splash(i).X), ConvertMapY(Splash(i).Y), (Int((GetTickCount - Splash(i).Tick) / 250) + 1) * 16, 0, 16, 16, 16, 16, -1
                Next i
            End If
            
            ' draw animations
7            If NumAnimations > 0 Then
                For i = 1 To MAX_BYTE
                    If AnimInstance(i).Used(0) Then
                        DrawAnimation i, 0
                    End If
                Next
            End If
        
            ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
8            For Y = 0 To Map.MaxY
                If NumCharacters > 0 Then
                    ' Npcs
                    For i = 1 To Npc_HighIndex
                        If MapNpc(i).Y = Y Then
                            Call DrawNpc(i)
                        End If
                    Next
                    
                    ' Players
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                            If Player(i).Y = Y Then
                                If TempPlayer(i).Fly = 0 Then DrawPlayer (i)
                                If Map.Tile(GetPlayerX(i), GetPlayerY(i)).Type = TILE_TYPE_RESOURCE Then
                                If i <> MyIndex Then Call DrawFishAlert(i)
                            End If
                            End If
                            If Player(i).Trans > 0 Then
                                If Player(i).TransAnimTick + 2500 < GetTickCount Then
                                    If Spell(Player(i).Trans).TransAnim > 0 Then Call DoAnimation(Spell(Player(i).Trans).TransAnim, GetPlayerX(i), GetPlayerY(i), TARGET_TYPE_PLAYER, i, GetPlayerDir(i))
                                    Player(i).TransAnimTick = GetTickCount
                                End If
                                If Player(i).EffectAnimTick + 250 < GetTickCount Then
                                    If Spell(Player(i).Trans).Effect > 0 Then Call CastEffect(Spell(Player(i).Trans).Effect, GetPlayerX(i), GetPlayerY(i))
                                    Player(i).EffectAnimTick = GetTickCount
                                End If
                            End If
                        End If
                    Next
                End If
                
9                If NumResources > 0 Then
                    If Resources_Init Then
                        If Resource_Index > 0 Then
                            For i = 1 To Resource_Index
                                If MapResource(i).Y = Y Then
                                    If Map.Tile(MapResource(i).X, MapResource(i).Y).Type = TILE_TYPE_RESOURCE Then Call DrawMapResource(i)
                                End If
                            Next
                        End If
                    End If
                End If
            Next
            
10            If NumProjectiles > 0 Then
                Call DrawProjectile
            End If
            
            ' animations
11            If NumAnimations > 0 Then
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
            'For Y = 1 To Map.MaxY
                
            'Next Y
            
12            If Transporte.Tipo <> 0 Then Call DrawTransporte
            
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
            
13            If Options.Clima = 1 Then DrawWeather
            If Options.Neblina = 1 Then DrawFog
            If Options.Tela = 1 Then DrawTint
            If Options.Ambiente = 1 Then DrawAmbient
            
            ' Render the bars
14            DrawBars
            
            ' Draw the target icon
            If myTarget > 0 Then
                If myTargetType = TARGET_TYPE_PLAYER Then
                    DrawTarget (Player(myTarget).X * 32) + TempPlayer(myTarget).XOffSet, (Player(myTarget).Y * 32) + TempPlayer(myTarget).YOffSet
                ElseIf myTargetType = TARGET_TYPE_NPC Then
                    DrawTarget (MapNpc(myTarget).X * 32) + TempMapNpc(myTarget).XOffSet, (MapNpc(myTarget).Y * 32) + TempMapNpc(myTarget).YOffSet
                End If
            End If
            
            ' Draw the hover icon
15            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If Player(i).Map = Player(MyIndex).Map Then
                        If CurX = Player(i).X And CurY = Player(i).Y Then
                            If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                                ' dont render lol
                            Else
                                DrawHover TARGET_TYPE_PLAYER, i, (Player(i).X * 32) + TempPlayer(i).XOffSet, (Player(i).Y * 32) + TempPlayer(i).YOffSet
                            End If
                        End If
                    End If
                End If
            Next
            For i = 1 To Npc_HighIndex
                If MapNpc(i).num > 0 Then
                    If CurX = MapNpc(i).X And CurY = MapNpc(i).Y Then
                        If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                            ' dont render lol
                        Else
                            DrawHover TARGET_TYPE_NPC, i, (MapNpc(i).X * 32) + TempMapNpc(i).XOffSet, (MapNpc(i).Y * 32) + TempMapNpc(i).YOffSet
                        End If
                    End If
                End If
            Next
            
            If UZ And (VIAGEMMAP = GetPlayerMap(MyIndex) Or VirgoMap = GetPlayerMap(MyIndex)) Then DrawPlanets
16          If ShenlongActive = 1 And ShenlongMap = GetPlayerMap(MyIndex) Or OutAnimationShenlongTick > GetTickCount + 5000 Then Call DrawShenlong
            If DrawThunder > 0 Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160): DrawThunder = DrawThunder - 1
            DrawBossMsg
17
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
19
            With srcRect
                .X1 = 0
                .X2 = frmMain.ScaleWidth
                .Y1 = 0
                .Y2 = frmMain.ScaleHeight
            End With
20
            If BFPS Then
                RenderText Font_Default, "FPS: " & CStr(GameFPS), 12, 100, Yellow, 0
            End If
21
            ' draw cursor, player X and Y locations
            If BLoc Then
                RenderText Font_Default, Trim$("cur x: " & CurX & " y: " & CurY), 12, 114, Yellow, 0
                RenderText Font_Default, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 12, 128, Yellow, 0
                RenderText Font_Default, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 12, 142, Yellow, 0
            End If
            
            If Not ShenlongActive = 1 Or ShenlongMap <> GetPlayerMap(MyIndex) Then
            ' draw player names
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    Call DrawPlayerName(i)
                End If
            Next
22
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_RESOURCE Then
                Call DrawFishAlert(MyIndex)
            Else
                FishingTime = 0
                isFishing = False
                BubbleOpaque = 0
            End If
23
            ' draw npc names
            For i = 1 To Npc_HighIndex
                If MapNpc(i).num > 0 Then
                    If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(MapNpc(i).num).HP Or (myTargetType = TARGET_TYPE_NPC And myTarget = i) Or Npc(MapNpc(i).num).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                       Call DrawNpcName(i)
                    End If
                End If
            Next
            
                ' draw the messages
            For i = 1 To MAX_BYTE
                If chatBubble(i).active Then
                    DrawChatBubble i
                End If
            Next
            
            For i = 1 To Action_HighIndex
                Call DrawActionMsg(i)
            Next i
            End If
24
            If Not hideGUI Then DrawGUI
25
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
        If Options.Debug >= 1 Then
            HandleError "Render_Graphics at " & Erl, "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
            Err.Clear
        End If
        MsgBox "Erro 13 - Impossível renderizar a cena, erro na plataforma gráfica."
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
    Direct3D_Window.hDeviceWindow = frmMain.hwnd 'Use frmMain as the device window.
    
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    offsetX = TempPlayer(MyIndex).XOffSet + PIC_X
    offsetY = TempPlayer(MyIndex).YOffSet + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        offsetX = 0
        If StartX = -1 Then
            If TempPlayer(MyIndex).XOffSet > 0 Then
                offsetX = TempPlayer(MyIndex).XOffSet
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0
        If StartY = -1 Then
            If TempPlayer(MyIndex).YOffSet > 0 Then
                offsetY = TempPlayer(MyIndex).YOffSet
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.MaxX Then
        offsetX = 32
        If EndX = Map.MaxX + 1 Then
            If TempPlayer(MyIndex).XOffSet < 0 Then
                offsetX = TempPlayer(MyIndex).XOffSet + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.MaxY Then
        offsetY = 32
        If EndY = Map.MaxY + 1 Then
            If TempPlayer(MyIndex).YOffSet < 0 Then
                offsetY = TempPlayer(MyIndex).YOffSet + PIC_Y
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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

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
    If Options.Debug >= 1 Then On Error GoTo errorhandler

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
Dim tilesetInUse() As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    ReDim tilesetInUse(0 To NumTileSets)
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If Map.Tile(X, Y).Layer(i).Tileset > 0 And Map.Tile(X, Y).Layer(i).Tileset <= NumTileSets Then
                    tilesetInUse(Map.Tile(X, Y).Layer(i).Tileset) = True
                End If
            Next
        Next
    Next
    
    For i = 1 To NumTileSets
        If tilesetInUse(i) Then
        
        Else
            ' unload tileset
            'Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            'Set Tex_Tileset(i) = Nothing
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'This function will make it much easier to setup the vertices with the info it needs.
Private Function Create_TLVertex(X As Single, Y As Single, z As Single, RHW As Single, color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX

    Create_TLVertex.X = X
    Create_TLVertex.Y = Y
    Create_TLVertex.z = z
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
    If frmMenu.visible Then
        If frmMenu.picCharacter.visible Then NewCharacterDrawSprite
    End If
End Sub
Public Sub DrawGUI()
Dim i As Long, X As Long, Y As Long
Dim Width As Long, Height As Long

    If RadarActive Then Exit Sub
    ' render shadow
    'EngineRenderRectangle Tex_GUI(27), 0, 0, 0, 0, 800, 64, 1, 64, 800, 64
    'EngineRenderRectangle Tex_GUI(26), 0, 600 - 64, 0, 0, 800, 64, 1, 64, 800, 64
    If ReceiveAttack + 200 < GetTickCount And ((Player(MyIndex).Vital(Vitals.HP) / Player(MyIndex).MaxVital(Vitals.HP)) * 100) > 20 Then
        RenderTexture Tex_GUI(23), 0, 0, 0, 0, 800, 64, 1, 64
        RenderTexture Tex_GUI(22), 0, 600 - 64, 0, 0, 800, 64, 1, 64
    Else
        RenderTexture Tex_GUI(29), 0, 0, 0, 0, 800, 64, 1, 64
        RenderTexture Tex_GUI(28), 0, 600 - 64, 0, 0, 800, 64, 1, 64
    End If
    
    ' render chatbox
        If inChat = False Then
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
    If GUIWindow(GUI_CONQUISTAS).visible And Not IsConquistaEmpty Then DrawConquistas
    If Not CloseDaily Then DrawDaily
    If InOwnPlanet Then DrawPlanetGUI
    DrawPopConquista
    DrawWindowService
    
    If InTutorial Then DrawTutorial
    
    If Player(MyIndex).IsDead = 1 Then
        DrawDeath
        GUIWindow(GUI_DEATH).visible = True
    Else
        GUIWindow(GUI_DEATH).visible = False
    End If
    
    ' Drag and drop
    DrawDragItem
    DrawDragSpell
    
    ' Descriptions
    DrawInventoryItemDesc
    DrawCharacterItemDesc
    DrawPlayerSpellDesc
    DrawBankItemDesc
    DrawTradeItemDesc
    DrawCharacterStatDesc
    DrawPlayerQuestDesc
End Sub

Sub DrawPlanetGUI()
    Dim i As Long
    If IsMovingObject Then Exit Sub
    If EditTargetX >= 0 And EditTargetY >= 0 Then
        If EditTargetX > Map.MaxX Or EditTargetY > Map.MaxY Then
            EditTargetX = 0
            EditTargetY = 0
            Exit Sub
        End If
        If Map.Tile(EditTargetX, EditTargetY).Type = TILE_TYPE_BLOCKED Then
            Dim X As Long, Y As Long, Image As DX8TextureRec
            X = ConvertMapX(EditTargetX * PIC_X)
            Y = ConvertMapY(EditTargetY * PIC_Y)
            If EditTargetX + 3 > Map.MaxX Then
                X = X - 96
            End If
            Buttons(55).X = X + 32
            Buttons(55).Y = Y + 16
            Buttons(55).visible = True
            RenderText Font_Default, "Valor: 10000z", X + 32, Y, Yellow
            RenderTexture ButtonImage(55), Buttons(55).X, Buttons(55).Y, 0, 0, Buttons(55).Width, Buttons(55).Height, Buttons(55).Width, Buttons(55).Height
        End If
        If Map.Tile(EditTargetX, EditTargetY).Type = TILE_TYPE_RESOURCE Then
            X = ConvertMapX(EditTargetX * PIC_X)
            Y = ConvertMapY(EditTargetY * PIC_Y)
            If EditTargetX + 3 > Map.MaxX Then
                X = X - 96
            End If
            Buttons(55).X = X + 32
            Buttons(55).Y = Y + 16
            Buttons(55).visible = True
            Buttons(57).X = X + 32
            Buttons(57).Y = Y + 48
            Buttons(57).visible = True
            RenderText Font_Default, "Remoção: 50000z", X + 32, Y, BrightRed
            RenderTexture ButtonImage(55), Buttons(55).X, Buttons(55).Y, 0, 0, Buttons(55).Width, Buttons(55).Height, Buttons(55).Width, Buttons(55).Height
            RenderTexture ButtonImage(57), Buttons(57).X, Buttons(57).Y, 0, 0, Buttons(57).Width, Buttons(57).Height, Buttons(57).Width, Buttons(57).Height
            If Resource(Map.Tile(EditTargetX, EditTargetY).data1).Evolution > 0 Then
                Buttons(58).X = X + 32
                Buttons(58).Y = Y + 80
                Buttons(58).visible = True
                RenderTexture ButtonImage(58), Buttons(58).X, Buttons(58).Y, 0, 0, Buttons(58).Width, Buttons(58).Height, Buttons(58).Width, Buttons(58).Height
                If (GlobalX >= Buttons(58).X And GlobalX <= Buttons(58).X + Buttons(58).Width) And (GlobalY >= Buttons(58).Y And GlobalY <= Buttons(58).Y + Buttons(58).Height) Then
                    RenderText Font_Default, "Tempo: " & Resource(Resource(Map.Tile(EditTargetX, EditTargetY).data1).Evolution).TimeToEvolute & "m", X + 128, Y + 80, White
                    RenderText Font_Default, "Moedas: " & Resource(Resource(Map.Tile(EditTargetX, EditTargetY).data1).Evolution).ECostGold & "z", X + 128, Y + 92, Grey
                    RenderText Font_Default, "Esp. V: " & Resource(Resource(Map.Tile(EditTargetX, EditTargetY).data1).Evolution).ECostRed, X + 128, Y + 104, BrightRed
                    RenderText Font_Default, "Esp. Az: " & Resource(Resource(Map.Tile(EditTargetX, EditTargetY).data1).Evolution).ECostBlue, X + 128, Y + 116, BrightBlue
                    RenderText Font_Default, "Esp. Am: " & Resource(Resource(Map.Tile(EditTargetX, EditTargetY).data1).Evolution).ECostYellow, X + 128, Y + 128, Yellow
                    RenderText Font_Default, "Centro lv.: " & Resource(Resource(Map.Tile(EditTargetX, EditTargetY).data1).Evolution).MinLevel, X + 128, Y + 140, White
                End If
            End If
            If Resource(Map.Tile(EditTargetX, EditTargetY).data1).ToolRequired = 1 Or Resource(Map.Tile(EditTargetX, EditTargetY).data1).ToolRequired = 2 Then
                Buttons(59).X = X + 32
                Buttons(59).Y = Y + 112
                Buttons(59).visible = True
                RenderTexture ButtonImage(59), Buttons(59).X, Buttons(59).Y, 0, 0, Buttons(59).Width, Buttons(59).Height, Buttons(59).Width, Buttons(59).Height
                Buttons(60).X = X + 32
                Buttons(60).Y = Y + 148
                Buttons(60).visible = True
                RenderTexture ButtonImage(60), Buttons(60).X, Buttons(60).Y, 0, 0, Buttons(60).Width, Buttons(60).Height, Buttons(60).Width, Buttons(60).Height
                If (GlobalX >= Buttons(60).X And GlobalX <= Buttons(60).X + Buttons(60).Width) And (GlobalY >= Buttons(60).Y And GlobalY <= Buttons(60).Y + Buttons(60).Height) Then
                    RenderText Font_Default, "Tempo: 24h", X + 128, Y + 148, BrightGreen
                    RenderText Font_Default, "Produção: 2x", X + 128, Y + 160, BrightGreen
                    RenderText Font_Default, "Custo: 50$", X + 128, Y + 172, BrightGreen
                End If
            End If
            If Resource(Map.Tile(EditTargetX, EditTargetY).data1).ItemReward > 0 Then
                Buttons(60).X = X + 32
                Buttons(60).Y = Y + 112
                Buttons(60).visible = True
                RenderTexture ButtonImage(60), Buttons(60).X, Buttons(60).Y, 0, 0, Buttons(60).Width, Buttons(60).Height, Buttons(60).Width, Buttons(60).Height
                If (GlobalX >= Buttons(60).X And GlobalX <= Buttons(60).X + Buttons(60).Width) And (GlobalY >= Buttons(60).Y And GlobalY <= Buttons(60).Y + Buttons(60).Height) Then
                    RenderText Font_Default, "Tempo: 24h", X + 128, Y + 112, BrightGreen
                    RenderText Font_Default, "Produção: 2x", X + 128, Y + 124, BrightGreen
                    RenderText Font_Default, "Custo: 100$", X + 128, Y + 136, BrightGreen
                End If
            End If
        End If
        If Map.Tile(EditTargetX, EditTargetY).Type = TILE_TYPE_NPCSPAWN Then
            X = ConvertMapX(EditTargetX * PIC_X)
            Y = ConvertMapY(EditTargetY * PIC_Y)
            Buttons(55).X = X + 32
            Buttons(55).Y = Y + 16
            Buttons(57).visible = True
            Buttons(57).X = X + 32
            Buttons(57).Y = Y + 48
            Buttons(57).visible = True
            RenderText Font_Default, "Remoção: 30000z", X + 32, Y, BrightRed
            RenderTexture ButtonImage(55), Buttons(55).X, Buttons(55).Y, 0, 0, Buttons(55).Width, Buttons(55).Height, Buttons(55).Width, Buttons(55).Height
            RenderTexture ButtonImage(57), Buttons(57).X, Buttons(57).Y, 0, 0, Buttons(57).Width, Buttons(57).Height, Buttons(57).Width, Buttons(57).Height
            If Npc(MapNpc(Map.Tile(EditTargetX, EditTargetY).data1).num).Evolution > 0 Then
                Buttons(58).X = X + 32
                Buttons(58).Y = Y + 80
                Buttons(58).visible = True
                RenderTexture ButtonImage(58), Buttons(58).X, Buttons(58).Y, 0, 0, Buttons(58).Width, Buttons(58).Height, Buttons(58).Width, Buttons(58).Height
                If (GlobalX >= Buttons(58).X And GlobalX <= Buttons(58).X + Buttons(58).Width) And (GlobalY >= Buttons(58).Y And GlobalY <= Buttons(58).Y + Buttons(58).Width) Then
                    RenderText Font_Default, "Tempo: " & Npc(Npc(MapNpc(Map.Tile(EditTargetX, EditTargetY).data1).num).Evolution).TimeToEvolute & "m", X + 128, Y + 80, White
                    RenderText Font_Default, "Moedas: " & Npc(Npc(MapNpc(Map.Tile(EditTargetX, EditTargetY).data1).num).Evolution).ECostGold & "z", X + 128, Y + 92, Grey
                    RenderText Font_Default, "Esp. V: " & Npc(Npc(MapNpc(Map.Tile(EditTargetX, EditTargetY).data1).num).Evolution).ECostRed, X + 128, Y + 104, BrightRed
                    RenderText Font_Default, "Esp. Az: " & Npc(Npc(MapNpc(Map.Tile(EditTargetX, EditTargetY).data1).num).Evolution).ECostBlue, X + 128, Y + 116, BrightBlue
                    RenderText Font_Default, "Esp. Am: " & Npc(Npc(MapNpc(Map.Tile(EditTargetX, EditTargetY).data1).num).Evolution).ECostYellow, X + 128, Y + 128, Yellow
                    RenderText Font_Default, "Centro lv.: " & Npc(Npc(MapNpc(Map.Tile(EditTargetX, EditTargetY).data1).num).Evolution).MinLevel, X + 128, Y + 140, White
                End If
            End If
        End If
        For i = 1 To MapSaibamans(GetPlayerMap(MyIndex)).TotalSaibamans
            If MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(i).Working = 1 Then
                If EditTargetX = MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(i).X And EditTargetY = MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(i).Y Then
                    Dim TextRemaining As String
                    TextRemaining = "Tempo restante: " & GetRemaining(MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(i).Remaining)
                    X = ConvertMapX(EditTargetX * PIC_X) + (PIC_X / 2) - (getWidth(Font_Default, TextRemaining) / 2)
                    Y = ConvertMapY(EditTargetY * PIC_Y) - 32
                    Buttons(56).X = X + (getWidth(Font_Default, TextRemaining) / 2) - Int(Buttons(56).Width / 2)
                    Buttons(56).Y = Y - 32
                    Buttons(56).visible = True
                    'Buttons(55).X = X + (getWidth(Font_Default, TextRemaining) / 2) - Int(Buttons(55).Width / 2)
                    'Buttons(55).Y = Y - 64
                    'Buttons(55).visible = True
                    RenderText Font_Default, TextRemaining, X, Y, Yellow
                    RenderTexture ButtonImage(56), Buttons(56).X, Buttons(56).Y, 0, 0, Buttons(56).Width, Buttons(56).Height, Buttons(56).Width, Buttons(56).Height
                    If (GlobalX >= Buttons(56).X And GlobalX <= Buttons(56).X + Buttons(56).Width) And (GlobalY >= Buttons(56).Y And GlobalY <= Buttons(56).Y + Buttons(56).Height) Then
                        Dim QuantNeed As Long
                        Dim Minutes As Long
                        Minutes = MapSaibamans(GetPlayerMap(MyIndex)).Saibaman(i).Remaining
                        QuantNeed = (100 - (5 * Int(Minutes / 100)))
                        If QuantNeed < 50 Then QuantNeed = 50
                        QuantNeed = (Minutes / 100) * QuantNeed
                        RenderText Font_Default, "Custo: " & QuantNeed & "$", Buttons(56).X + Buttons(56).Width + 16, Buttons(56).Y, BrightGreen
                    End If
                    'RenderTexture ButtonImage(55), Buttons(55).X, Buttons(55).Y, 0, 0, Buttons(55).Width, Buttons(55).Height, Buttons(55).Width, Buttons(55).Height
                End If
            End If
        Next i
    End If
    
    For i = 55 To 60
        X = Buttons(i).X
        Y = Buttons(i).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(i).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(i).Height) Then
            Buttons(i).State = 1
        Else
            Buttons(i).State = 0
        End If
    Next i
End Sub
Function ButtonImage(ByVal ButtonNum As Long) As DX8TextureRec
    Select Case Buttons(ButtonNum).State
        Case 0: ButtonImage = Tex_Buttons(Buttons(ButtonNum).PicNum)
        Case 1: ButtonImage = Tex_Buttons_h(Buttons(ButtonNum).PicNum)
        Case 2: ButtonImage = Tex_Buttons_c(Buttons(ButtonNum).PicNum)
    End Select
End Function

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
Dim YOffSet As Long, XOffSet As Long

    ' calculate the offset
    Select Case Map.Tile(X, Y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            YOffSet = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            XOffSet = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            YOffSet = -32
    End Select
    
    ' Draw the quarter
    'EngineRenderRectangle Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, 16, 16
    RenderTexture Tex_Tileset(Map.Tile(X, Y).Layer(layerNum).Tileset), destX, destY, Autotile(X, Y).Layer(layerNum).srcX(quarterNum) + XOffSet, Autotile(X, Y).Layer(layerNum).srcY(quarterNum) + YOffSet, 16, 16, 16, 16, -1
End Sub

Public Sub DrawItem(ByVal ItemNum As Long)
Dim PicNum As Integer, dontRender As Boolean, i As Long, tmpIndex As Long
Dim X As Long, Y As Long
Dim Left As Long

    If MapItem(ItemNum).Gravity < 10 Then MapItem(ItemNum).Gravity = MapItem(ItemNum).Gravity + 1
    
    If MapItem(ItemNum).YOffSet + MapItem(ItemNum).Gravity > 0 Then
        MapItem(ItemNum).YOffSet = 0
        If MapItem(ItemNum).PlaySound = False And MapItem(ItemNum).Gravity = 10 Then
            Call PlaySound("Drop.mp3", -1, -1)
            MapItem(ItemNum).PlaySound = True
        End If
    Else
        MapItem(ItemNum).YOffSet = MapItem(ItemNum).YOffSet + MapItem(ItemNum).Gravity
    End If
    
    X = (MapItem(ItemNum).X * PIC_X) + MapItem(ItemNum).XOffSet
    Y = (MapItem(ItemNum).Y * PIC_Y) + MapItem(ItemNum).YOffSet + MapItem(ItemNum).YOnSet

    RenderTexture Tex_Shadow, ConvertMapX(X) + 8, ConvertMapY(Y) - 4, 0, 0, 16, 32, 32, 32, D3DColorRGBA(255, 255, 255, 100)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    PicNum = Item(MapItem(ItemNum).num).Pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

     ' if it's not us then don't render
    If MapItem(ItemNum).PlayerName <> vbNullString Then
        If Trim$(MapItem(ItemNum).PlayerName) <> Trim$(GetPlayerName(MyIndex)) Then
            dontRender = True
        End If
        ' make sure it's not a party drop
        If Party.Leader > 0 Then
            For i = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party.Member(i)
                If tmpIndex > 0 Then
                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(ItemNum).PlayerName) Then
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
    
    If Item(MapItem(ItemNum).num).Type = ITEM_TYPE_CONSUME Then
        If MapItem(ItemNum).BalanceTick < GetTickCount Then
            If MapItem(ItemNum).BalanceDir = 0 Then
                MapItem(ItemNum).BalanceValue = MapItem(ItemNum).BalanceValue - 1
                If MapItem(ItemNum).BalanceValue < -3 Then MapItem(ItemNum).BalanceDir = 1
            Else
                MapItem(ItemNum).BalanceValue = MapItem(ItemNum).BalanceValue + 1
                If MapItem(ItemNum).BalanceValue > 3 Then MapItem(ItemNum).BalanceDir = 0
            End If
            MapItem(ItemNum).BalanceTick = GetTickCount + 200
        End If
        Y = Y + MapItem(ItemNum).BalanceValue
    End If
    
    'If Not dontRender Then EngineRenderRectangle Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32, 32, 32
    If Not dontRender Then
        RenderTexture Tex_Item(PicNum), X, Y, Left, 0, 32, 32, 32, 32
    End If
End Sub

Public Sub DrawDragItem()
    Dim PicNum As Integer, ItemNum As Long
    
    If DragInvSlotNum = 0 Then Exit Sub
    
    ItemNum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)
    If Not ItemNum > 0 Then Exit Sub
    
    PicNum = Item(ItemNum).Pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

    'EngineRenderRectangle Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawDragSpell()
    Dim PicNum As Integer, SpellNum As Long
    
    If DragSpell = 0 Then Exit Sub
    
    SpellNum = PlayerSpells(DragSpell)
    If Not SpellNum > 0 Then Exit Sub
    
    PicNum = Spell(SpellNum).Icon

    If PicNum < 1 Or PicNum > NumSpellIcons Then Exit Sub

    'EngineRenderRectangle Tex_Spellicon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_SpellIcon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawHotbar()
Dim i As Long, X As Long, Y As Long, t As Long, sS As String
Dim Width As Long, Height As Long, color As Long

    Dim RenderGUI As Boolean
    
    If RadarActive Then Exit Sub
    
    RenderGUI = True
    If UZ And MatchActive > 0 Then
        RenderGUI = False
        If GlobalY < 100 Then RenderGUI = True
    End If
    If Options.PickMenu = 1 Then RenderGUI = True
    
    X = GUIWindow(GUI_HOTBAR).X - 3
    Y = GUIWindow(GUI_HOTBAR).Y - 3
    Width = 493
    Height = 43
    If RenderGUI Then
    RenderTexture Tex_GUI(31), X, Y, 0, 0, Width, Height, Width, Height
    
    
    For i = 1 To MAX_HOTBAR
        ' draw the box
        X = GUIWindow(GUI_HOTBAR).X + ((i - 1) * (5 + 36))
        Y = GUIWindow(GUI_HOTBAR).Y
        Width = 36
        Height = 36
        'EngineRenderRectangle Tex_GUI(2), x, y, 0, 0, width, height, width, height, width, heigh
        If RenderGUI Then RenderTexture Tex_GUI(2), X, Y, 0, 0, Width, Height, Width, Height
        ' draw the icon
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(i).Slot).name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 Then
                        'EngineRenderRectangle Tex_Item(Item(Hotbar(i).Slot).Pic), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        RenderTexture Tex_Item(Item(Hotbar(i).Slot).Pic), X + 2, Y + 2, 0, 0, 32, 32, 32, 32
                    End If
                End If
            Case 2 ' spell
                If Len(Spell(Hotbar(i).Slot).name) > 0 Then
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
    End If
    
End Sub
Public Sub DrawInventory()
Dim i As Long, X As Long, Y As Long, ItemNum As Long, ItemPic As Long
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
        ItemNum = GetPlayerInvItemNum(MyIndex, i)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    If TradeYourOffer(X).num = i Then
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

Public Sub DrawCharacterStatDesc()
Dim dX As Long, dY As Long, X As Long, Y As Long, Stat As String
    
    If Not GUIWindow(GUI_CHARACTER).visible Then Exit Sub
    
    X = GUIWindow(GUI_CHARACTER).X
    Y = GUIWindow(GUI_CHARACTER).Y
    
    ' render stats
    ' FOR
    dX = X + 180
    dY = Y + 138
    If GlobalX > dX And GlobalX < dX + 16 Then
        If GlobalY > dY And GlobalY < dY + 12 Then
            Stat = "Força"
        End If
    End If
    
    ' CON
    dY = dY + 12
    If GlobalX > dX And GlobalX < dX + 16 Then
        If GlobalY > dY And GlobalY < dY + 12 Then
            Stat = "Constituição"
        End If
    End If
    
    ' KI
    dY = dY + 12
    If GlobalX > dX And GlobalX < dX + 16 Then
        If GlobalY > dY And GlobalY < dY + 12 Then
            Stat = "KI"
        End If
    End If
    
    ' DES
    dY = dY + 12
    If GlobalX > dX And GlobalX < dX + 16 Then
        If GlobalY > dY And GlobalY < dY + 12 Then
            Stat = "Destreza"
        End If
    End If
    
    ' TEC
    dY = dY + 12
    If GlobalX > dX And GlobalX < dX + 16 Then
        If GlobalY > dY And GlobalY < dY + 12 Then
            Stat = "Tecnica"
        End If
    End If

    
    'If Item(GetPlayerInvItemNum(MyIndex, invSlot)).BindType > 0 And PlayerInv(invSlot).bound > 0 Then isSB = True
    If Stat <> "" Then DrawStatDesc Stat, GUIWindow(GUI_CHARACTER).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_CHARACTER).Y + 64

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
Dim CostItem As Long, CostValue As Long, ItemNum As Long, sString As String, Width As Long, Height As Long, i As Long

    If slotNum = 0 Then Exit Sub
    
    If InShop <= 0 Then Exit Sub
    
    For i = 1 To 5
    ' draw the window
    Width = 190
    Height = 36
    
    ' find out the cost
    If Not isShop Then
        ' inventory - default to gold
        ItemNum = GetPlayerInvItemNum(MyIndex, slotNum)
        If ItemNum = 0 Then Exit Sub
        CostItem = MoedaZ
        CostValue = (Item(ItemNum).Price / 100) * Shop(InShop).BuyRate
        sString = "Será comprado por"
        If Item(ItemNum).Price = 0 Then
            sString = printf("Este item não pode ser vendido!")
            RenderTexture Tex_GUI(24), X, Y + 80, 0, 0, Width, Height, Width, Height
            RenderText Font_Default, sString, X + 4, Y + 83, BrightRed
            Exit Sub
        End If
        Y = Y + 80
    Else
        ItemNum = Shop(InShop).TradeItem(slotNum).Item
        If ItemNum = 0 Then Exit Sub
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
    
    If Item(CostItem).Type <> ITEM_TYPE_TITULO Then
        RenderText Font_Default, ConvertCurrency(CostValue) & " " & Trim$(Item(CostItem).name), X + 4, Y + 18 + (36 * (i - 1)), White
    Else
        RenderText Font_Default, Trim$(Item(CostItem).name), X + 4, Y + 18 + (36 * (i - 1)), White
    End If
    Next i
End Sub

Public Sub DrawItemDesc(ByVal ItemNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal soulBound As Boolean = False)
Dim colour As Long, colourw As Long, descString As String, theName As String, className As String, levelTxt As String, sInfo() As String, i As Long, Width As Long, Height As Long

    ' get out
    If ItemNum = 0 Or ItemNum > MAX_ITEMS Then Exit Sub

    ' render the window
    Width = 190
    If Not Trim$(Item(ItemNum).Desc) = vbNullString Then
        Height = 210
    Else
        Height = 126
    End If
    'EngineRenderRectangle Tex_GUI(6), x, y, 0, 0, width, height, width, height, width, height
    
    'Cor da janela
    Select Case Item(ItemNum).Rarity
        Case 0 ' Comum
            colourw = D3DColorRGBA(255, 255, 255, 200)
            colour = White
            className = ""
        Case 1 ' Raro
            colourw = D3DColorRGBA(170, 170, 255, 200)
            colour = BrightCyan
            className = "Item Raro"
        Case 2 ' Lendário
            colourw = D3DColorRGBA(255, 50, 50, 240)
            colour = Yellow
            className = "Item Lendário"
        Case 3 ' Relíquia
            colourw = D3DColorRGBA(250, 110, 0, 240)
            colour = Yellow
            className = "Relíquia"
        Case 4 ' Evento
            colourw = D3DColorRGBA(255, 255, 0, 250)
            colour = BrightBlue
            className = "Item de evento"
        Case 5 ' Único
            colourw = D3DColorRGBA(255, 100, 255, 200)
            colour = Cyan
            className = "Item único"
    End Select
    
    RenderTexture Tex_GUI(8), X, Y, 0, 0, Width, Height, Width, Height, colourw
    
    ' make sure it has a sprite
    If Item(ItemNum).Pic > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Item(Item(itemnum).Pic), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 64, 64
        RenderTexture Tex_Item(Item(ItemNum).Pic), X + 16, Y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not Trim$(Item(ItemNum).Desc) = vbNullString Then
        RenderText Font_Default, WordWrap(Trim$(Item(ItemNum).Desc), Width - 10), X + 10, Y + 128, White
    End If
    
    If Not soulBound Then
        theName = Trim$(Item(ItemNum).name)
    Else
        theName = "(SB) " & Trim$(Item(ItemNum).name)
    End If
    
    ' render name
    RenderText Font_Default, theName, X + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), Y + 6, colour
    
    ' class req
    'If Item(itemNum).ClassReq > 0 Then
    '    className = Trim$(Class(Item(itemNum).ClassReq).name)
    '    ' do we match it?
    '    If GetPlayerClass(MyIndex) = Item(itemNum).ClassReq Then
    '        colour = Green
    '    Else
    '        colour = BrightRed
    '    End If
    'Else
    '    className = "No class req."
    '    colour = Green
    'End If
    
    RenderText Font_Default, className, X + 48 - (EngineGetTextWidth(Font_Default, className) \ 2), Y + 92, colour
    
    ' level
    If Item(ItemNum).LevelReq > 0 Then
        levelTxt = "Level " & Item(ItemNum).LevelReq
        ' do we match it?
        If GetPlayerLevel(MyIndex) >= Item(ItemNum).LevelReq Then
            colour = Green
        Else
            colour = BrightRed
        End If
    Else
        levelTxt = printf("Todos os niveis.")
        colour = Green
    End If
    RenderText Font_Default, levelTxt, X + 48 - (EngineGetTextWidth(Font_Default, levelTxt) \ 2), Y + 107, colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    i = 1
    ReDim Preserve sInfo(1 To i) As String
    Select Case Item(ItemNum).Type
        Case ITEM_TYPE_NONE
            sInfo(i) = printf("Sem tipo")
        Case ITEM_TYPE_WEAPON
            sInfo(i) = printf("Arma")
        Case ITEM_TYPE_ARMOR
            sInfo(i) = printf("Peitoral")
        Case ITEM_TYPE_HELMET
            sInfo(i) = printf("Calça")
        Case ITEM_TYPE_SHIELD
            sInfo(i) = printf("Bota")
        Case ITEM_TYPE_CONSUME
            sInfo(i) = printf("Consumo")
        Case ITEM_TYPE_CURRENCY
            sInfo(i) = printf("Mercadoria")
        Case ITEM_TYPE_SPELL
            sInfo(i) = printf("Tecnica")
        Case ITEM_TYPE_DRAGONBALL
            sInfo(i) = printf("Esfera do dragão")
        Case ITEM_TYPE_TITULO
            sInfo(i) = printf("Título de jogador")
    End Select
    
    ' more info
    Select Case Item(ItemNum).Type
        Case ITEM_TYPE_NONE, ITEM_TYPE_CURRENCY
            ' binding
            If Item(ItemNum).BindType = 1 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Trava ao pegar")
            ElseIf Item(ItemNum).BindType = 2 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Trava ao equipar")
            End If
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Valor: %d z", Val(Item(ItemNum).Price))
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SHIELD
            ' binding
            If Item(ItemNum).BindType = 1 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Trava ao pegar")
            ElseIf Item(ItemNum).BindType = 2 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Trava ao equipar")
            End If
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Valor: %d z", Val(Item(ItemNum).Price))
            ' damage/defence
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Dano: %d", Val(Item(ItemNum).data2))
                ' speed
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Veloc.: %d s", Val(Item(ItemNum).speed / 1000))
            Else
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Defesa: %d", Val(Item(ItemNum).data2))
            End If
            ' stat bonuses
            If Item(ItemNum).Add_Stat(Stats.Strength) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(ItemNum).Add_Stat(Stats.Strength) & " FOR"
            End If
            If Item(ItemNum).Add_Stat(Stats.Endurance) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(ItemNum).Add_Stat(Stats.Endurance) & " CON"
            End If
            If Item(ItemNum).Add_Stat(Stats.Intelligence) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(ItemNum).Add_Stat(Stats.Intelligence) & " KI"
            End If
            If Item(ItemNum).Add_Stat(Stats.Agility) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(ItemNum).Add_Stat(Stats.Agility) & " DES"
            End If
            If Item(ItemNum).Add_Stat(Stats.Willpower) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(ItemNum).Add_Stat(Stats.Willpower) & " TEC"
            End If
        Case ITEM_TYPE_CONSUME
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Valor: %d z", Val(Item(ItemNum).Price))
            If Item(ItemNum).CastSpell > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Efetua técnica")
            End If
            If Item(ItemNum).AddHP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(ItemNum).AddHP & " HP"
            End If
            If Item(ItemNum).AddMP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(ItemNum).AddMP & " KI"
            End If
            If Item(ItemNum).AddEXP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(ItemNum).AddEXP & " EXP"
            End If
        Case ITEM_TYPE_SPELL
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Valor: %d z", Val(Item(ItemNum).Price))
        Case ITEM_TYPE_DRAGONBALL
            DrawDesejos X, Y, colourw
        Case ITEM_TYPE_TITULO
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Valor: %d z", Val(Item(ItemNum).Price))
    End Select
    
    ' go through and render all this shit
    Y = Y + 12
    For i = 1 To UBound(sInfo)
        Y = Y + 12
        RenderText Font_Default, sInfo(i), X + 141 - (EngineGetTextWidth(Font_Default, sInfo(i)) \ 2), Y, White
    Next
End Sub
Public Sub DrawDesejos(ByVal X As Long, ByVal Y As Long, ByVal colour As Long)
    RenderTexture Tex_GUI(5), X - 210, Y, 0, 0, 256, 256, 256, 256, colour
    RenderText Font_Default, "Lista de desejos", X - 105 - (getWidth(Font_Default, "Lista de desejos") / 2), Y + 8, Yellow
    RenderText Font_Default, "Diga: eu desejo", X - 205, Y + 24, White
    RenderText Font_Default, WordWrap("resetar (Permite redistribuir seus pontos)", 195), X - 205, Y + 40, Yellow
    RenderText Font_Default, WordWrap("seu poder mais precioso (Habilidade lendária)", 195), X - 205, Y + 68, BrightCyan
    RenderText Font_Default, WordWrap("entrar na sala do tempo", 195), X - 205, Y + 86, White
    RenderText Font_Default, WordWrap("me tornar um deus", 195), X - 205, Y + 104, White
End Sub
Public Sub DrawPlayerSpellDesc()
Dim spellSlot As Long
    
    If Not GUIWindow(GUI_SPELLS).visible Then Exit Sub
    If DragSpell > 0 Then Exit Sub
    
    spellSlot = IsPlayerSpell(GlobalX, GlobalY)
    If spellSlot > 0 Then
        If PlayerSpells(spellSlot) > 0 Then
            DrawSpellDesc PlayerSpells(spellSlot), GlobalX + 32, GlobalY, spellSlot
        End If
    End If
    
    spellSlot = IsPlayerEvoluteSpell(GlobalX, GlobalY)
    If spellSlot > 0 And Not spellSlot = MAX_SPELLS Then
        If Not (Spell(spellSlot).Icon = 0) Then DrawSpellDesc spellSlot, GlobalX + 32, GlobalY
    End If
End Sub

Public Sub DrawSpellDesc(ByVal SpellNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal spellSlot As Long = 0, Optional RenderLast As Boolean = True)
Dim colour As Long, theName As String, sUse As String, sInfo() As String, i As Long, tmpWidth As Long, barWidth As Long
Dim Width As Long, Height As Long, lastSpell As Long
    
    ' don't show desc when dragging
    If DragSpell > 0 Then Exit Sub
    
    ' get out
    If SpellNum = 0 Then Exit Sub
    
    If spellSlot = 0 And RenderLast Then
        If HasSpell(SpellNum) Then
            If Spell(SpellNum).Upgrade > 0 Then
                lastSpell = SpellNum
                SpellNum = Spell(SpellNum).Upgrade
                'Magia atual
                DrawSpellDesc lastSpell, X - 206, Y, spellSlot, False
            Else
                Exit Sub
            End If
        End If
    End If
    
    If spellSlot = 0 Then
        If HasSpell(SpellNum) Then
            If RenderLast Then
                RenderText Font_Default, "Evolução", X + 95 - (getWidth(Font_Default, "Evolução") / 2), Y - 12, Yellow
            Else
                RenderText Font_Default, "Habilidade atual", X + 95 - (getWidth(Font_Default, "Habilidade atual") / 2), Y - 12, Yellow
            End If
        End If
    End If
    
    If spellSlot = 0 Then
        RenderTexture Tex_GUI(25), X, Y - 60, 0, 0, 190, 128, 128, 64
        RenderText Font_Default, "Requisitos", X + 95 - (getWidth(Font_Default, "Requisitos") / 2), Y - 80, Yellow
        If Spell(SpellNum).Requisite > 0 Then
            RenderTexture Tex_Item(Item(Spell(SpellNum).Requisite).Pic), X - 6, Y - 66, 0, 0, 32, 32, 32, 32
            colour = BrightRed
            If Player(MyIndex).Titulo > 0 Then
                If Item(Player(MyIndex).Titulo).LevelReq < Item(Spell(SpellNum).Requisite).LevelReq Then
                    colour = BrightRed
                Else
                    colour = White
                End If
            Else
                For i = 1 To MAX_INV
                    If GetPlayerInvItemNum(MyIndex, i) > 0 Then
                        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_TITULO Then
                            If Item(GetPlayerInvItemNum(MyIndex, i)).LevelReq < Item(Spell(SpellNum).Requisite).LevelReq Then
                                colour = BrightRed
                            Else
                                colour = White
                            End If
                        End If
                    End If
                Next i
            End If
            RenderText Font_Default, "Cargo:" & Trim$(Item(Spell(SpellNum).Requisite).name), X + 16, Y - 52, colour
        Else
            RenderText Font_Default, "Habilidade especial!", X + 16, Y - 52, BrightCyan
        End If
        If Spell(SpellNum).Item > 0 Then
            colour = BrightRed
            If HasItem(Spell(SpellNum).Item) >= Spell(SpellNum).Price Then colour = White
            If Spell(SpellNum).Requisite = 0 Then RenderTexture Tex_Item(Item(Spell(SpellNum).Item).Pic), X - 6, Y - 40, 0, 0, 32, 32, 32, 32
            RenderText Font_Default, "Preço:" & Spell(SpellNum).Price & " " & Trim$(Item(Spell(SpellNum).Item).name), X + 16, Y - 32, colour
        Else
            RenderText Font_Default, "Habilidade bloqueada", X + 16, Y - 32, BrightRed
        End If
    End If

    ' render the window
    Width = 190
    If Not Trim$(Spell(SpellNum).Desc) = vbNullString Then
        Height = 210
    Else
        Height = 126
    End If
    'EngineRenderRectangle Tex_GUI(29), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(8), X, Y, 0, 0, Width, Height, Width, Height
    
    ' make sure it has a sprite
    If Spell(SpellNum).Icon > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Spellicon(Spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 32, 32
        RenderTexture Tex_SpellIcon(Spell(SpellNum).Icon), X + 16, Y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not IsRefining Then
        If Not Trim$(Spell(SpellNum).Desc) = vbNullString Then
            RenderText Font_Default, WordWrap(Trim$(Spell(SpellNum).Desc), Width - 10), X + 10, Y + 128, White
        End If
    Else
        If RenderLast Then
        If Spell(lastSpell).Price > 0 And Spell(lastSpell).Item > 0 Then RenderText Font_Default, printf("Preço: %d %s", Val(Spell(lastSpell).Price) & "," & Trim$(Item(Spell(lastSpell).Item).name)), X + 10, Y + 128, White
        If Spell(lastSpell).Requisite > 0 Then RenderText Font_Default, printf("Requere: %s", Trim$(Item(Spell(lastSpell).Requisite).name)), X + 10, Y + 144, White
        End If
    End If
    
    ' render name
    colour = White
    theName = Trim$(Spell(SpellNum).name)
    RenderText Font_Default, theName, X + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), Y + 6, colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    i = 1
    ReDim Preserve sInfo(1 To i) As String
    Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_DAMAGEHP
            sInfo(i) = printf("Ofensiva")
        Case SPELL_TYPE_DAMAGEMP
            sInfo(i) = printf("Fadiga")
        Case SPELL_TYPE_HEALHP
            sInfo(i) = printf("Cura")
        Case SPELL_TYPE_HEALMP
            sInfo(i) = printf("Recuperação")
        Case SPELL_TYPE_WARP
            sInfo(i) = printf("Teleporte")
        Case SPELL_TYPE_LINEAR
            sInfo(i) = printf("Onda")
    End Select
    
    ' more info
    Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_DAMAGEHP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
            ' damage
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Eficiência: %d", Val(Spell(SpellNum).Vital))
            
            ' mp cost
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Custo: %d KI", Val(Spell(SpellNum).MPCost))
            
            ' cast time
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Conjuração: %d s", Val(Spell(SpellNum).CastTime))
            
            ' cd time
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Recarga: %d s", Val(Spell(SpellNum).CDTime))
            
            ' aoe
            If Spell(SpellNum).AoE > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "AoE: " & Spell(SpellNum).AoE
            End If
            
            ' range
            If Spell(SpellNum).Range > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Alcance: %d", Val(Spell(SpellNum).Range))
            End If
            
            ' impacto
            If Spell(SpellNum).Impact > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Impacto: %d", Val(Spell(SpellNum).Impact))
            End If
            
            ' stun
            If Spell(SpellNum).StunDuration > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Paralis.: %d s", Val(Spell(SpellNum).StunDuration))
            End If

        Case SPELL_TYPE_LINEAR
            ' damage
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Dano: %d", Val(Spell(SpellNum).Vital))
            
            ' mp cost
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Custo: %d KI", Val(Spell(SpellNum).MPCost))
            
            ' cast time
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Conjuração: %d s", Val(Spell(SpellNum).CastTime))
            
            ' cd time
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = printf("Recarga: %d s", Val(Spell(SpellNum).CDTime))
            
            ' aoe
            If Spell(SpellNum).AoE > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "AoE: " & Spell(SpellNum).AoE
            End If
            
            ' range
            If Spell(SpellNum).Range > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Alcance: %d", Val(Spell(SpellNum).Range))
            End If
            
            ' impact
            If Spell(SpellNum).Impact > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Impacto: %d", Val(Spell(SpellNum).Impact))
            End If
            
            ' stun
            If Spell(SpellNum).StunDuration > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = printf("Paralis.: %d s", Val(Spell(SpellNum).StunDuration))
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
Dim i As Long, X As Long, Y As Long, SpellNum As Long, spellpic As Long
Dim Top As Long, Left As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 480
    Height = 384
    'EngineRenderRectangle Tex_GUI(4), GUIWindow(GUI_SPELLS).x, GUIWindow(GUI_SPELLS).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(26), GUIWindow(GUI_SPELLS).X, GUIWindow(GUI_SPELLS).Y, 0, 0, Width, Height, Width, Height
    
    If IsRefining = True Then
        RenderText Font_Default, printf("Passe o mouse para ver a evolução"), GUIWindow(GUI_SPELLS).X, GUIWindow(GUI_SPELLS).Y - 32, Yellow
        RenderText Font_Default, printf("Duplo clique para selecionar"), GUIWindow(GUI_SPELLS).X, GUIWindow(GUI_SPELLS).Y - 16, Yellow
    End If
    
    ' render skills
    For i = 1 To MAX_PLAYER_SPELLS
        SpellNum = PlayerSpells(i)

        ' make sure not dragging it
        If DragSpell = i Then GoTo NextLoop
        
        ' actually render
        If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
            spellpic = Spell(SpellNum).Icon

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
    
    RenderText Font_Default, printf("Próximas skills"), GUIWindow(GUI_SPELLS).X + 48, GUIWindow(GUI_SPELLS).Y + 180, Yellow
    
    
    Dim z As Long
    z = 0
    For i = 1 To MAX_SPELLS
        SpellNum = i

        ' actually render
        If SpellList(SpellNum) Then
            spellpic = Spell(SpellNum).Icon

            If spellpic > 0 And spellpic <= NumSpellIcons Then
                Top = GUIWindow(GUI_SPELLS).Y + 160 + SpellTop + ((SpellOffsetY + 32) * ((z) \ SpellColumns))
                Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((z) Mod SpellColumns)))
                'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                If Not CanBuySpell(SpellNum) Then
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                    RenderTexture Tex_SpellIcon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                Else
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                    RenderTexture Tex_SpellIcon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
            z = z + 1
        End If
    Next
End Sub

Sub UpdateSpellList()
    Dim i As Long
    For i = 1 To MAX_SPELLS
        SpellList(i) = ShowSpell(i)
    Next i
End Sub

Function CanBuySpell(ByVal SpellNum As Long) As Boolean
    If GetPlayerLevel(MyIndex) >= Spell(SpellNum).LevelReq Then
        
        Dim ItemNum As Long, ItemValue As Long
        
        ItemNum = Spell(SpellNum).Item
        ItemValue = Spell(SpellNum).Price
        
        If Spell(SpellNum).Requisite > 0 Then
            If Item(Spell(SpellNum).Requisite).Type <> ITEM_TYPE_TITULO Then
                If Not HasItem(Spell(SpellNum).Requisite) > 0 Then
                    Exit Function
                End If
            Else
                If Player(MyIndex).Titulo > 0 Then
                    If Item(Player(MyIndex).Titulo).LevelReq < Item(Spell(SpellNum).Requisite).LevelReq Then
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            End If
        End If
        
        If ItemNum > 0 Then
            If HasItem(ItemNum) < ItemValue Then
                Exit Function
            End If
        End If
        
        CanBuySpell = True
    End If
End Function

Function HasItem(ByVal ItemNum As Long) As Long
    Dim i As Long
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) = ItemNum Then
            HasItem = GetPlayerInvItemValue(MyIndex, i)
            Exit Function
        End If
    Next i
End Function

Function ShowSpell(ByVal SpellNum As Long) As Boolean
    If Spell(SpellNum).AccessReq > 0 Then Exit Function
    If HasSpell(SpellNum) Then Exit Function
    If Spell(SpellNum).LevelReq > GetPlayerLevel(MyIndex) + 10 Then Exit Function
    If Spell(SpellNum).Upgrade > 0 Then
        'Primeiras skills
        If Not HasSpell(Spell(SpellNum).Upgrade) Then
            If Not HasAntecessor(SpellNum) Then
                ShowSpell = True
                Exit Function
            End If
        End If
    Else
        If Not HasAntecessor(SpellNum) Then
            ShowSpell = True
            Exit Function
        End If
    End If
            
    Dim i As Long
    For i = 1 To MAX_PLAYER_SPELLS
        If PlayerSpells(i) > 0 Then
            If Spell(PlayerSpells(i)).Upgrade = SpellNum Then
                ShowSpell = True
                Exit For
            End If
        End If
    Next i
End Function

Function HasAntecessor(ByVal SpellNum As Long) As Boolean
    Dim i As Long
    For i = 1 To MAX_SPELLS
        If Trim$(Spell(i).name) = vbNullString Then
            Exit For
        Else
            If Spell(i).Upgrade = SpellNum Then
                HasAntecessor = True
                Exit For
            End If
        End If
    Next i
End Function

Function HasSpell(ByVal SpellNum As Long) As Boolean
    Dim i As Long, SpellRealName As String
    Dim SpellLength As Byte

    'Evolutions
    SpellRealName = Trim$(Spell(SpellNum).name)
    SpellLength = Len(SpellRealName)

    For i = 1 To MAX_PLAYER_SPELLS

        If PlayerSpells(i) > 0 Then
            If PlayerSpells(i) = SpellNum Then
                HasSpell = True
                Exit Function
            End If
        
            If Mid(Trim$(Spell(PlayerSpells(i)).name), 1, SpellLength) = SpellRealName Then
                HasSpell = True
                Exit Function
            End If
        End If

    Next

End Function

Public Sub DrawCooldown()
    Dim i As Long, X As Long, Y As Long
    Dim HaveCooldown As Boolean
    X = 16
    Y = 160
    HaveCooldown = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If PlayerSpells(i) > 0 Then
            If SpellCD(i) > 0 Then
                Y = Y + 32
                RenderTexture Tex_SpellIcon(Spell(PlayerSpells(i)).Icon), X, Y, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                HaveCooldown = True
            End If
        End If
    Next i
    
    If HaveCooldown Then
        RenderText Font_Default, "Recarga", 4, 172, BrightGreen
    End If
End Sub

Public Sub DrawEquipment()
Dim X As Long, Y As Long, i As Long
Dim ItemNum As Long, ItemPic As DX8TextureRec

    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(MyIndex, i)

        ' get the item sprite
        If ItemNum > 0 Then
            ItemPic = Tex_Item(Item(ItemNum).Pic)
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

Public Sub DrawDeath()
Dim X As Long, Y As Long, i As Long
        Y = (frmMain.ScaleHeight / 2) - (GUIWindow(GUI_DEATH).Height / 2)
        X = (frmMain.ScaleWidth / 2) - (GUIWindow(GUI_DEATH).Width / 2)
        GUIWindow(GUI_DEATH).Y = Y
        GUIWindow(GUI_DEATH).X = X
        RenderTexture Tex_GUI(30), X, Y, 0, 0, GUIWindow(GUI_DEATH).Width, GUIWindow(GUI_DEATH).Height, GUIWindow(GUI_DEATH).Width, GUIWindow(GUI_DEATH).Height
        If GlobalX >= Buttons(54).X + X And GlobalX <= Buttons(54).X + Buttons(54).Width + X Then
            If GlobalY >= Buttons(54).Y + Y And GlobalY <= Buttons(54).Y + Buttons(54).Height + Y Then
            RenderTexture Tex_Buttons(Buttons(54).PicNum), X + Buttons(54).X, Y + Buttons(54).Y, 0, 0, Buttons(54).Width, Buttons(54).Height, Buttons(54).Width, Buttons(54).Height
            End If
        End If
End Sub

Public Sub DrawCharacter()
Dim X As Long, Y As Long, i As Long, dX As Long, dY As Long, tmpString As String, ButtonNum As Long
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
    If Player(MyIndex).IsGod = 0 Then
        tmpString = Trim$(GetPlayerName(MyIndex)) & " - Level " & GetPlayerLevel(MyIndex)
    Else
        tmpString = "Nível Divino " & Player(MyIndex).GodLevel & " - " & Player(MyIndex).IsGod & " Ascenções"
    End If
    RenderText Font_Default, tmpString, X + 7 + (187 / 2) - (EngineGetTextWidth(Font_Default, tmpString) / 2), Y + 9, White
    
    Dim StatColor As Long
    Dim TotalLevels As Long, StatPerc As Long
    TotalLevels = GetPlayerStat(MyIndex, Strength) + GetPlayerStat(MyIndex, Endurance) + GetPlayerStat(MyIndex, Intelligence) + GetPlayerStat(MyIndex, Willpower) + GetPlayerStat(MyIndex, Agility)
    
    ' render stats
    dX = X + 20
    dY = Y + 135
    RenderText Font_Default, printf("Força"), dX, dY, White
    StatColor = White
    If (GetPlayerStat(MyIndex, Strength) / TotalLevels) * 100 < 20 Then StatColor = Yellow
    If (GetPlayerStat(MyIndex, Strength) / TotalLevels) * 100 < 10 Then StatColor = BrightRed
    RenderText Font_Default, GetPlayerStat(MyIndex, Strength), dX + 155 - EngineGetTextWidth(Font_Default, GetPlayerStat(MyIndex, Strength)), dY, StatColor
    dY = dY + 12
    RenderText Font_Default, printf("Constituição"), dX, dY, White
    StatColor = White
    If (GetPlayerStat(MyIndex, Endurance) / TotalLevels) * 100 < 20 Then StatColor = Yellow
    If (GetPlayerStat(MyIndex, Endurance) / TotalLevels) * 100 < 10 Then StatColor = BrightRed
    RenderText Font_Default, GetPlayerStat(MyIndex, Endurance), dX + 155 - EngineGetTextWidth(Font_Default, GetPlayerStat(MyIndex, Endurance)), dY, StatColor
    dY = dY + 12
    RenderText Font_Default, "KI", dX, dY, White
    StatColor = White
    If (GetPlayerStat(MyIndex, Intelligence) / TotalLevels) * 100 < 20 Then StatColor = Yellow
    If (GetPlayerStat(MyIndex, Intelligence) / TotalLevels) * 100 < 10 Then StatColor = BrightRed
    RenderText Font_Default, GetPlayerStat(MyIndex, Intelligence), dX + 155 - EngineGetTextWidth(Font_Default, GetPlayerStat(MyIndex, Intelligence)), dY, StatColor
    dY = dY + 12
    RenderText Font_Default, printf("Destreza"), dX, dY, White
    StatColor = White
    If (GetPlayerStat(MyIndex, Agility) / TotalLevels) * 100 < 20 Then StatColor = Yellow
    If (GetPlayerStat(MyIndex, Agility) / TotalLevels) * 100 < 10 Then StatColor = BrightRed
    RenderText Font_Default, GetPlayerStat(MyIndex, Agility), dX + 155 - EngineGetTextWidth(Font_Default, GetPlayerStat(MyIndex, Agility)), dY, StatColor
    dY = dY + 12
    RenderText Font_Default, printf("Técnica"), dX, dY, White
    StatColor = White
    If (GetPlayerStat(MyIndex, Willpower) / TotalLevels) * 100 < 20 Then StatColor = Yellow
    If (GetPlayerStat(MyIndex, Willpower) / TotalLevels) * 100 < 10 Then StatColor = BrightRed
    RenderText Font_Default, GetPlayerStat(MyIndex, Willpower), dX + 155 - EngineGetTextWidth(Font_Default, GetPlayerStat(MyIndex, Willpower)), dY, StatColor
    dY = Y + 25
    If GetPlayerPOINTS(MyIndex) > 0 Then RenderText Font_Default, printf("Pontos: %d/%d", GetPlayerPOINTS(MyIndex) & "," & ((Player(MyIndex).Level - 1) * 3)), dX, dY, Yellow
    
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
    '        Dim StatPoints As Long, StatPrevLevel As Long, NextStatPoints As Long, NextStatPrevLevel As Long
    '        StatPoints = GetPlayerStatPoints(MyIndex, i)
    '        StatPrevLevel = GetPlayerStatPrevLevel(MyIndex, i)
    '        NextStatPoints = GetPlayerStatNextLevel(MyIndex, i)
    '        NextStatPrevLevel = GetPlayerStatPrevLevel(MyIndex, i)
    '        .Right = .Left + (sWidth * ((StatPoints - StatPrevLevel) / (NextStatPoints - NextStatPrevLevel)))
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
        For ButtonNum = 16 To 20
            X = GUIWindow(GUI_CHARACTER).X + Buttons(ButtonNum).X
            Y = GUIWindow(GUI_CHARACTER).Y + Buttons(ButtonNum).Y
            Width = Buttons(ButtonNum).Width
            Height = Buttons(ButtonNum).Height
            ' render accept button
            If Buttons(ButtonNum).State = 2 Then
                ' we're clicked boyo
                Width = Buttons(ButtonNum).Width
                Height = Buttons(ButtonNum).Height
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(ButtonNum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(ButtonNum).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(ButtonNum).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(ButtonNum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = ButtonNum Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = ButtonNum
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(ButtonNum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = ButtonNum Then lastButtonSound = 0
            End If
        Next
    End If
    
    Dim Cargo As Long
    Dim NextCargo As Long
    
    If Player(MyIndex).Titulo > 0 Then
        Cargo = Player(MyIndex).Titulo
        NextCargo = Player(MyIndex).Titulo + 1
    Else
        For i = 1 To MAX_INV
            If PlayerInv(i).num > 0 Then
                If Item(PlayerInv(i).num).Type = ITEM_TYPE_TITULO Then
                    Cargo = PlayerInv(i).num
                    NextCargo = PlayerInv(i).num + 1
                    Exit For
                End If
            End If
        Next i
    End If
    
    If Cargo = 115 Then NextCargo = 99
    
    dX = X + 60
    dY = Y + 36
    
    If UZ Then
        RenderTexture Tex_Item(Item(Cargo).Pic), dX - 16, dY - 8, 0, 0, 32, 32, 32, 32
        RenderText Font_Default, Trim$(Item(Cargo).name), dX, dY, White
        
        If Item(NextCargo).Type <> ITEM_TYPE_TITULO Then
            RenderText Font_Default, "Cargo maximo!", dX, dY + 12, White
        Else
            'RenderTexture Tex_Item(Item(NextCargo).Pic), dX - 16, dY + 20, 0, 0, 32, 32, 32, 32
            RenderText Font_Default, "Próximo cargo", dX - 10, dY + 14, Yellow
            'RenderText Font_Default, Trim$(Item(NextCargo).Name), dX, dY + 28, White
            Dim Prov As Boolean
            Prov = True
            If GetPlayerLevel(MyIndex) < Item(NextCargo).LevelReq Then
                RenderText Font_Default, "Level " & Item(NextCargo).LevelReq, dX - 6, dY + 26, BrightRed
                Prov = False
            Else
                RenderText Font_Default, "Level " & Item(NextCargo).LevelReq, dX - 6, dY + 26, BrightGreen
            End If
            
            Dim ServicesReq As Long
            
            If Cargo < 115 Then
                ServicesReq = 10 + ((Cargo - 99) * 20)
            Else
                ServicesReq = 5
            End If
            If Player(MyIndex).NumServices < ServicesReq Then
                RenderText Font_Default, "Serviços " & Player(MyIndex).NumServices & "/" & ServicesReq, dX - 6, dY + 38, BrightRed
                Prov = False
            Else
                RenderText Font_Default, "Serviços " & Player(MyIndex).NumServices & "/" & ServicesReq, dX - 6, dY + 38, BrightGreen
            End If
            If HasItem(114) = 0 Then
                RenderText Font_Default, "Missão", dX - 6, dY + 50, BrightRed
                Prov = False
            Else
                RenderText Font_Default, "Missão", dX - 6, dY + 50, BrightGreen
            End If
            If Prov Then
                RenderText Font_Default, "Provação", dX - 6, dY + 62, White
            Else
                RenderText Font_Default, "Provação", dX - 6, dY + 62, BrightRed
            End If
        End If
        
        If GlobalX >= GUIWindow(GUI_CHARACTER).X + 50 And GlobalX <= GUIWindow(GUI_CHARACTER).X + 150 Then
            If GlobalY >= GUIWindow(GUI_CHARACTER).Y + 36 And GlobalY <= GUIWindow(GUI_CHARACTER).Y + 136 Then
                RenderTexture Tex_GUI(5), GUIWindow(GUI_CHARACTER).X - GUIWindow(GUI_CHARACTER).Width - 4, GUIWindow(GUI_CHARACTER).Y, 0, 0, 195, 250, 195, 250
                dX = GUIWindow(GUI_CHARACTER).X - GUIWindow(GUI_CHARACTER).Width - 4
                dY = GUIWindow(GUI_CHARACTER).Y
                
                RenderText Font_Default, "Cargo no exército", dX + 97 - (getWidth(Font_Default, "Cargo no exército") / 2), dY, Yellow
                RenderText Font_Default, WordWrap("Seu cargo no exército Sayajin garante novas habilidades e equipamentos.", 190), dX + 4, dY + 14, White
                
                Dim colour As Long
                If GetPlayerLevel(MyIndex) < Item(NextCargo).LevelReq Or Player(MyIndex).NumServices < ServicesReq Then
                    colour = BrightRed
                Else
                    colour = BrightGreen
                End If
                
                RenderText Font_Default, WordWrap("Para subir de cargo, atinja o nível necessário e complete o número de serviços com o agente de serviços.", 190), dX + 4, dY + 55, colour
                
                If HasItem(114) = 0 Then
                    colour = BrightRed
                Else
                    colour = BrightGreen
                End If
                
                RenderText Font_Default, WordWrap("Complete a missão entregue pelo Rei Vegeta", 190), dX + 4, dY + 130, colour
                
                If Prov Then
                    colour = BrightGreen
                Else
                    colour = BrightRed
                End If
                
                RenderText Font_Default, WordWrap("Com tudo pronto, pague a taxa da provação para o Rei Vegeta e vença a provação!", 190), dX + 4, dY + 160, colour
            End If
        End If
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
    
    RenderText Font_Default, "Audio", GUIWindow(GUI_OPTIONS).X + 16, GUIWindow(GUI_OPTIONS).Y + 23, Yellow
    RenderText Font_Default, printf("Musica:"), GUIWindow(GUI_OPTIONS).X + 16, GUIWindow(GUI_OPTIONS).Y + 40, White
    RenderText Font_Default, printf("Som:"), GUIWindow(GUI_OPTIONS).X + 16, GUIWindow(GUI_OPTIONS).Y + 63, White
    RenderText Font_Default, "Volume: ", GUIWindow(GUI_OPTIONS).X + 16, GUIWindow(GUI_OPTIONS).Y + 99, White
    RenderText Font_Default, printf("Gráficos"), GUIWindow(GUI_OPTIONS).X + 16, GUIWindow(GUI_OPTIONS).Y + 128, Yellow
    RenderText Font_Default, printf("Ambiente:"), GUIWindow(GUI_OPTIONS).X + 16, GUIWindow(GUI_OPTIONS).Y + 145, White
    RenderText Font_Default, printf("Tela:"), GUIWindow(GUI_OPTIONS).X + 16, GUIWindow(GUI_OPTIONS).Y + 168, White
    RenderText Font_Default, printf("Clima:"), GUIWindow(GUI_OPTIONS).X + 16, GUIWindow(GUI_OPTIONS).Y + 191, White
    RenderText Font_Default, printf("Neblina:"), GUIWindow(GUI_OPTIONS).X + 16, GUIWindow(GUI_OPTIONS).Y + 214, White
    'RenderText Font_Default, "Volume: ", GUIWindow(GUI_OPTIONS).X + 20, GUIWindow(GUI_OPTIONS).Y + 145, White
    
    RenderText Font_Default, Options.volume, GUIWindow(GUI_OPTIONS).X + 120, GUIWindow(GUI_OPTIONS).Y + 99, Yellow
    ' draw buttons
    For i = 26 To 29
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
    For i = 44 To 45
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
    ' draw buttons
    For i = 46 To 53
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
    RenderText Font_Default, "10% de EXP bônus por membro no grupo!", GUIWindow(GUI_PARTY).X + (Width / 2) - (getWidth(Font_Default, "10% de EXP bônus por membro no grupo!") / 2), GUIWindow(GUI_PARTY).Y - 14, Yellow
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
Dim X As Long, Y As Long, ButtonNum As Long
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
    
    If CurrencyMenu = 5 Then
        If IsNumeric(sDialogue) Then
            If Val(sDialogue) > MaxGravity Then sDialogue = MaxGravity
            RenderText Font_Default, "Preço por hora:" & GravityValue(Val(sDialogue)) & "z", X + 87 + 123 - (getWidth(Font_Default, "Preço por hora:" & GravityValue(Val(sDialogue)) & "z") / 2), Y + 80, White
        Else
            If sDialogue <> vbNullString Then sDialogue = 10
        End If
    End If
    
    If CurrencyMenu = 6 Then
        If IsNumeric(sDialogue) Then
            If Val(sDialogue) > 6 Then sDialogue = 6
            RenderText Font_Default, "Preço final:" & (GravityValue(Val(SelectedGravity)) * Val(sDialogue)) & "z", X + 87 + 123 - (getWidth(Font_Default, "Preço final:" & (GravityValue(Val(SelectedGravity)) * Val(sDialogue)) & "z") / 2), Y + 80, White
        Else
            If sDialogue <> vbNullString Then sDialogue = 1
        End If
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If CurrencyAcceptState = 2 Then
        ' clicked
        RenderText Font_Default, printf("[Aceitar]"), X, Y, Grey
    Else
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' hover
            RenderText Font_Default, printf("[Aceitar]"), X, Y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 1 Then
                PlaySound Sound_ButtonHover, -1, -1
                lastNpcChatsound = 1
            End If
        Else
            ' normal
            RenderText Font_Default, printf("[Aceitar]"), X, Y, Green
            ' reset sound if needed
            If lastNpcChatsound = 1 Then lastNpcChatsound = 0
        End If
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Fechar]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If CurrencyCloseState = 2 Then
        ' clicked
        RenderText Font_Default, printf("[Fechar]"), X, Y, Grey
    Else
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' hover
            RenderText Font_Default, printf("[Fechar]"), X, Y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 2 Then
                PlaySound Sound_ButtonHover, -1, -1
                lastNpcChatsound = 2
            End If
        Else
            ' normal
            RenderText Font_Default, printf("[Fechar]"), X, Y, Yellow
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
                RenderText Font_Default, printf("[Aceitar]"), X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Default, printf("[Aceitar]"), X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = 1 Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = 1
                    End If
                Else
                    ' normal
                    RenderText Font_Default, printf("[Aceitar]"), X, Y, Green
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
            RenderText Font_Default, printf("[Fechar]"), X, Y, Grey
        Else
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                ' hover
                RenderText Font_Default, printf("[Fechar]"), X, Y, Cyan
                ' play sound if needed
                If Not lastNpcChatsound = 3 Then
                    PlaySound Sound_ButtonHover, -1, -1
                    lastNpcChatsound = 3
                End If
            Else
                ' normal
                RenderText Font_Default, printf("[Fechar]"), X, Y, Yellow
                ' reset sound if needed
                If lastNpcChatsound = 3 Then lastNpcChatsound = 0
            End If
        End If
    End If
End Sub

Public Sub DrawShop()
Dim i As Long, X As Long, Y As Long, ItemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = GUIWindow(GUI_SHOP).Width
    Height = GUIWindow(GUI_SHOP).Height
    'EngineRenderRectangle Tex_GUI(23), GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(20), GUIWindow(GUI_SHOP).X, GUIWindow(GUI_SHOP).Y, 0, 0, Width, Height, Width, Height
    
    RenderText Font_Default, printf("Duplo clique para efetuar a compra"), GUIWindow(GUI_SHOP).X + 22, GUIWindow(GUI_SHOP).Y + 8, White
    If GUIWindow(GUI_INVENTORY).visible = True Then RenderText Font_Default, printf("Duplo clique para efetuar a venda"), GUIWindow(GUI_INVENTORY).X, GUIWindow(GUI_INVENTORY).Y - 16, White
    
    ' render the shop items
    For i = 1 To MAX_TRADES
        ItemNum = Shop(InShop).TradeItem(i).Item
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            If ItemPic > 0 And ItemPic <= numitems Then
                
                Top = GUIWindow(GUI_SHOP).Y + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                Left = GUIWindow(GUI_SHOP).X + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                RenderTexture Tex_Selection, Left, Top, 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 100)
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

Dim RenderMenu As Boolean

    If RadarActive Then Exit Sub
    
    RenderMenu = True
    If UZ And MatchActive > 0 Then
        RenderMenu = False
        If GlobalY > 500 Then RenderMenu = True
    End If


    If RenderMenu Then
    
    If UZ And VIAGEMMAP <> GetPlayerMap(MyIndex) And Not MatchActive > 0 Then
    Dim BannerAlpha2 As Byte
    BannerAlpha2 = 100
    If GlobalX >= 640 And GlobalX <= 800 Then
        If GlobalY >= 470 And GlobalY <= 520 Then
            BannerAlpha2 = 255
        End If
    End If
    If Buttons(6).visible = False Then RenderTexture Tex_GUI(35), 640, 470, 0, 0, 160, 64, 160, 64, D3DColorRGBA(255, 255, 255, BannerAlpha2)
    BannerAlpha2 = 100
    If GlobalX >= 580 And GlobalX <= 640 Then
        If GlobalY >= 470 And GlobalY <= 520 Then
            BannerAlpha2 = 255
        End If
    End If
    RenderTexture Tex_GUI(40), 580, 470, 0, 0, 64, 64, 64, 64, D3DColorRGBA(255, 255, 255, BannerAlpha2)
    End If
    ' draw background
    X = GUIWindow(GUI_MENU).X
    Y = GUIWindow(GUI_MENU).Y
    Width = GUIWindow(GUI_MENU).Width
    Height = GUIWindow(GUI_MENU).Height
    'EngineRenderRectangle Tex_GUI(3), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(3), X, Y, 0, 0, Width, Height, Width, Height
    
    RenderText Font_Default, MoedasZ & " Z", GUIWindow(GUI_MENU).X + 30, GUIWindow(GUI_MENU).Y + 33, Grey, , True
    
    ' draw buttons
    For i = 1 To 8
        Dim ButtonIndex As Long
        Dim HoverIndex As Long
        If i <= 6 Then
            ButtonIndex = i
        Else
            ButtonIndex = 60 - 6 + i
        End If
        If Buttons(ButtonIndex).visible Then
            ' set co-ordinate
            X = GUIWindow(GUI_MENU).X + Buttons(ButtonIndex).X
            Y = GUIWindow(GUI_MENU).Y + Buttons(ButtonIndex).Y
            Width = Buttons(ButtonIndex).Width
            Height = Buttons(ButtonIndex).Height
            ' check for state
            If Buttons(ButtonIndex).State = 2 And Not (InTutorial And Player(MyIndex).InTutorial = 0 And (TutorialStep = 1 Or i = TutorialShowIcon)) Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonindex).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(ButtonIndex).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(ButtonIndex).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(ButtonIndex).Height) Or (InTutorial And Player(MyIndex).InTutorial = 0 And (TutorialStep = 1 Or i = TutorialShowIcon)) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonindex).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(ButtonIndex).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                HoverIndex = ButtonIndex
                If Not lastButtonSound = i Then
                    If Not InTutorial Then PlaySound Sound_ButtonHover, -1, -1
                    lastButtonSound = i
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonindex).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(ButtonIndex).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = i Then lastButtonSound = 0
            End If
        End If
    Next
    If HoverIndex > 0 Then
        ButtonIndex = HoverIndex
        X = GUIWindow(GUI_MENU).X + Buttons(ButtonIndex).X
        Y = GUIWindow(GUI_MENU).Y + Buttons(ButtonIndex).Y
        If ButtonIndex < 6 Then
            RenderTexture Tex_GUI(25), X - 46, Y - 40, 0, 0, 128, 64, 128, 64, D3DColorRGBA(255, 255, 255, 100)
            RenderText Font_Default, MenuButtonName(ButtonIndex), X - (getWidth(Font_Default, MenuButtonName(ButtonIndex)) / 2) + 18, Y - 35, Grey, , True
        Else
            RenderTexture Tex_GUI(25), X - 50, Y - 40, 0, 0, 128, 64, 64, 64, D3DColorRGBA(255, 255, 255, 100)
            RenderText Font_Default, MenuButtonName(ButtonIndex), X - (getWidth(Font_Default, MenuButtonName(ButtonIndex)) / 2), Y - 35, Grey, , True
            'RenderTexture Tex_GUI(25), X - 46 + 76, Y - 40, 0, 0, 6, 64, 122, 64, D3DColorRGBA(255, 255, 255, 100)
        End If
    End If
    End If
End Sub


Public Sub DrawBank()
Dim i As Long, X As Long, Y As Long, ItemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long, Width As Long
Dim Height As Long

    Width = GUIWindow(GUI_BANK).Width
    Height = GUIWindow(GUI_BANK).Height
    
    RenderTexture Tex_GUI(26), GUIWindow(GUI_BANK).X, GUIWindow(GUI_BANK).Y, 0, 0, Width, Height, Width, Height
    
    ' render the bank items' are you serous? that is it??? maybe... one sec :D :Polol
        For i = 1 To MAX_BANK
            ItemNum = GetBankItemNum(i)
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
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
Dim i As Long, X As Long, Y As Long, ItemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, colour As Long, Width As Long
Dim Height As Long

    Width = GUIWindow(GUI_TRADE).Width
    Height = GUIWindow(GUI_TRADE).Width
    RenderTexture Tex_GUI(18), GUIWindow(GUI_TRADE).X, GUIWindow(GUI_TRADE).Y, 0, 0, Width, Height, Width, Height
        For i = 1 To MAX_INV
            ' render your offer
            ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                    Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).X + 29 + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    ' If item is a stack - draw the amount you have
                    If TradeYourOffer(i).value > 1 Then
                        Y = Top + 21
                        X = Left - 4
                            
                        Amount = CStr(TradeYourOffer(i).value)
                            
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
            ItemNum = TradeTheirOffer(i).num
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
                If ItemPic > 0 And ItemPic <= numitems Then
                
                    Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop - 2 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).X + 257 + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    ' If item is a stack - draw the amount you have
                    If TradeTheirOffer(i).value > 1 Then
                        Y = Top + 21
                        X = Left - 4
                                
                        Amount = CStr(TradeTheirOffer(i).value)
                                
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
    RenderText Font_Default, printf("Sua troca vale: %d", Val(YourWorth)), GUIWindow(GUI_TRADE).X + 21, GUIWindow(GUI_TRADE).Y + 299, White
    RenderText Font_Default, printf("A troca dele vale: %d", Val(TheirWorth)), GUIWindow(GUI_TRADE).X + 250, GUIWindow(GUI_TRADE).Y + 299, White
    RenderText Font_Default, TradeStatus, (GUIWindow(GUI_TRADE).Width / 2) - (EngineGetTextWidth(Font_Default, TradeStatus) / 2), GUIWindow(GUI_TRADE).Y + 317, Yellow
    DrawTradeItemDesc
End Sub

Public Sub DrawTradeItemDesc()
Dim tradeNum As Long
    If Not GUIWindow(GUI_TRADE).visible Then Exit Sub
        
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum > 0 Then
        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).num) > 0 Then
            DrawItemDesc GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).num), GUIWindow(GUI_TRADE).X + 480 + 10, GUIWindow(GUI_TRADE).Y
        End If
    End If
    
    tradeNum = IsTradeItem(GlobalX, GlobalY, False)
    If tradeNum > 0 Then
        If TradeTheirOffer(tradeNum).num > 0 Then
            DrawItemDesc TradeTheirOffer(tradeNum).num, GUIWindow(GUI_TRADE).X + 480 + 10, GUIWindow(GUI_TRADE).Y
        End If
    End If
End Sub

Public Sub DrawGUIBars()
Dim tmpWidth As Long, barWidth As Long, X As Long, Y As Long, dX As Long, dY As Long, sString As String
Dim Width As Long, Height As Long, i As Long

    Dim RenderGUI As Boolean
    
    If RadarActive Then Exit Sub
    
    RenderGUI = True
    If UZ And MatchActive > 0 Then
        RenderGUI = False
        If GlobalY < 100 Then RenderGUI = True
    End If
    If Options.PickMenu = 1 Then RenderGUI = True
    If RenderGUI = False Then DrawCooldown

    ' backwindow + empty bars
    X = GUIWindow(GUI_BARS).X
    Y = GUIWindow(GUI_BARS).Y
    Width = GUIWindow(GUI_BARS).Width
    Height = 60
    If RenderGUI Then
    Height = GUIWindow(GUI_BARS).Height
    'EngineRenderRectangle Tex_GUI(4), x, y, 0, 0, width, height, width, height, width, height
    If Player(MyIndex).VIP >= 1 Then
        RenderTexture Tex_GUI(38), X + 90, Y + Height - 16, 0, 0, 113, 21, 113, 21
        RenderTexture Tex_GUI(39), X + 90, Y + Height - 9, 0, 0, ((Player(MyIndex).VIPExp / VIPNextLevel) * 112), 14, (Player(MyIndex).VIPExp / VIPNextLevel) * 112, 14
        Call RenderText(Font_Default, "VIP Lv." & Player(MyIndex).VIP, X + 100, Y + Height - 10, White, , True)
        If GlobalX > X + 90 And GlobalX < X + 200 Then
            If GlobalY > Y + Height - 16 And GlobalY < Y + Height + 8 Then
                RenderTexture Tex_GUI(24), X + 210, Y + Height - 16, 0, 0, 256, 64, 256, 64
                RenderText Font_Default, "EXP AUMENTADA EM +" & (50 + (Player(MyIndex).VIP - 1) * 5) & "%", X + 214, Y + Height - 16, BrightGreen
                RenderText Font_Default, "MOEDAS AUMENTADAS EM +" & (50 + (Player(MyIndex).VIP - 1) * 5) & "%", X + 214, Y + Height, Yellow
            End If
        End If
    End If
    RenderTexture Tex_GUI(4), X, Y, 0, 0, Width, Height, Width, Height
    
    If Player(MyIndex).IsGod = 0 Then
        Call RenderText(Font_Default, GetPlayerName(MyIndex) & " Lv." & GetPlayerLevel(MyIndex), X + 90, Y + 76, Yellow, , True)
    Else
        Call RenderText(Font_Default, "Lv. Divino " & Player(MyIndex).GodLevel, X + 90, Y + 76, Yellow, , True)
    End If
    
    Dim ButtonX As Long, ButtonY As Long
    ButtonX = X + Width - 96
    ButtonY = Y + Height - 24
    
    If Options.PickMenu = 0 Then
        RenderTexture Tex_Buttons(28), ButtonX, ButtonY, 0, 0, 12, 12, 12, 12
    Else
        RenderTexture Tex_Buttons_c(28), ButtonX, ButtonY, 0, 0, 12, 12, 12, 12
    End If
    If GlobalX >= ButtonX And GlobalX <= ButtonX + 12 Then
        If GlobalY >= ButtonY And GlobalY <= ButtonY + 12 Then
            RenderTexture Tex_Buttons_h(28), ButtonX, ButtonY, 0, 0, 12, 12, 12, 12
            Call RenderText(Font_Default, "Fixar Menu durante combate", ButtonX + 16, ButtonY, Yellow, , True)
        End If
    End If
    
    'Esferas
    For i = 1 To 7
        Select Case i
            Case 1
                dX = X + 66
                dY = Y + 77
            Case 2
                dX = X + 43
                dY = Y + 82
            Case 3
                dX = X + 20
                dY = Y + 75
            Case 4
                dX = X + 6
                dY = Y + 57
            Case 5
                dX = X + 4
                dY = Y + 34
            Case 6
                dX = X + 15
                dY = Y + 14
            Case 7
                dX = X + 35
                dY = Y + 3
        End Select
        If HaveDragonball(i) = True Then RenderTexture Tex_Dragonballs, dX, dY, (i - 1) * 15, 0, 15, 15, 15, 15, D3DColorRGBA(255, 255, 150, 255)
    Next i
    End If
    
    ' hardcoded for POT textures
    barWidth = 192
    
    If RenderGUI = False Then
        X = X - 70
    End If
    
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
    
    If RenderGUI Then
    barWidth = 175
    
    If Player(MyIndex).IsGod = 0 Then
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
            sString = printf("Level maximo")
        End If
        dX = X + 89 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
        dY = Y + 60
        RenderText Font_Default, sString, dX, dY, White
        
    Else
        ' exp bar
        If Player(MyIndex).GodExp > 0 And GodNextLevel > 0 Then
        If Player(MyIndex).GodLevel < MAX_LEVELS Then
            BarWidth_GuiEXP_Max = ((Player(MyIndex).GodExp / barWidth) / (GodNextLevel / barWidth)) * barWidth
        Else
            BarWidth_GuiEXP_Max = barWidth
        End If
        RenderTexture Tex_GUI(15), X + 89, Y + 60, 0, 0, BarWidth_GuiEXP, Tex_GUI(15).Height, BarWidth_GuiEXP, Tex_GUI(15).Height, D3DColorRGBA(200, 0, 255, 255)
        End If
        ' render exp
        If Player(MyIndex).GodLevel < MAX_LEVELS Then
            sString = Player(MyIndex).GodExp & "/" & GodNextLevel
        Else
            sString = printf("Level maximo")
        End If
        dX = X + 89 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
        dY = Y + 60
        RenderText Font_Default, sString, dX, dY, White
    End If
    End If
    If RenderGUI = False Then
        X = X + 70
    End If
    
    If UZ Then
        If MatchActive > 0 Then
             ' backwindow + empty bars
            Y = GUIWindow(GUI_BARS).Y + Height + 16
            Width = Tex_GUI(32).Width
            Height = Tex_GUI(32).Height
            RenderTexture Tex_GUI(32), X, Y, 0, 0, Width, Height, Width, Height
            
            Dim BarUZWidth As Long
            barWidth = 183
    
            BarUZWidth = ((MatchPoints / barWidth) / (MatchNeedPoints / barWidth)) * barWidth
            RenderTexture Tex_GUI(14), X + 1, Y + 2, 0, 0, BarUZWidth, Tex_GUI(14).Height, BarUZWidth, Tex_GUI(14).Height, D3DColorRGBA(255, 255, 255, 255)
            
            barWidth = 175
            BarUZWidth = barWidth - (((MatchNPCs / barWidth) / (MAX_MAP_NPCS / barWidth)) * barWidth)
            RenderTexture Tex_GUI(15), X + 1, Y + 20, 0, 0, BarUZWidth, Tex_GUI(15).Height, BarUZWidth, Tex_GUI(15).Height, D3DColorRGBA(255, 255 - (255 * (MatchNPCs / barWidth) / (MAX_MAP_NPCS / barWidth)), 255 - (255 * (MatchNPCs / barWidth) / (MAX_MAP_NPCS / barWidth)), 255)
            
            RenderText Font_Default, "Invasão", X + Width - (getWidth(Font_Default, "Invasão")), Y - 16, BrightRed
            
            RenderText Font_Default, "Progresso", X, Y - 8, Yellow
            RenderText Font_Default, Int((MatchPoints / MatchNeedPoints) * 100) & "%", X + 91 - (getWidth(Font_Default, Int((MatchPoints / MatchNeedPoints) * 100) & "%") / 2), Y, Yellow
            
            RenderText Font_Default, "Dominação", X, Y + 10, Yellow
            RenderText Font_Default, (100 - Int((MatchNPCs / MAX_MAP_NPCS) * 100)) & "%", X + 87 - (getWidth(Font_Default, (100 - Int((MatchNPCs / MAX_MAP_NPCS) * 100)) & "%") / 2), Y + 18, Yellow
            
            Dim StarDegree As Long
            If StarAnimation > GetTickCount Then
                StarDegree = 3.6 * Int((GetTickCount Mod 1000) / 10)
            Else
                StarDegree = 0
            End If
            If StarX > X + Width - 48 + (getWidth(Font_Default, "Invasão")) Or StarY > Y - 24 Then
                If StarX > X + Width - 48 + (getWidth(Font_Default, "Invasão")) Then StarX = StarX - 10
                If StarY > Y - 24 Then StarY = StarY - 10
                RenderTexture Tex_NewGUI(NewGui.Estrela), StarX, StarY, 0, 0, 31, 31, 31, 31
            End If
            RenderTexture Tex_NewGUI(NewGui.Estrela), X + Width - 48 + (getWidth(Font_Default, "Invasão")), Y - 24, 0, 0, 31, 31, 31, 31, , StarDegree
            RenderText Font_Default, MatchStars, X + Width - 50 + (getWidth(Font_Default, "Invasão")) + 16 - Int(getWidth(Font_Default, MatchStars) / 2), Y - 16, White
        End If
    End If
    
    If UZ Then
        If VIAGEMMAP <> GetPlayerMap(MyIndex) And VirgoMap <> GetPlayerMap(MyIndex) Then
            If DailyQuestCompleted = 0 Then
                Call RenderText(Font_Default, "Missão diária", X + 12, Y + 112, Yellow, , True)
                Call RenderText(Font_Default, DailyQuestMsg, X + 12, Y + 124, White, , True)
                Call RenderText(Font_Default, "Atingido: " & DailyQuestObjective, X + 12, Y + 136, BrightGreen, , True)
            End If
        End If
        
        If PlanetService > 0 And Not MatchActive > 0 Then
            Call RenderText(Font_Default, "Serviço", X + 12, Y + 160, Yellow, , True)
    
            Dim MissionText As String
            Select Case Planets(PlanetService).Type
                Case 0: MissionText = "Capture o planeta " & Trim$(Planets(PlanetService).name)
                Case 1: MissionText = "Derrote o chefe no planeta " & Trim$(Planets(PlanetService).name)
                Case 2: MissionText = "Colete os recursos preciosos no planeta " & Trim$(Planets(PlanetService).name)
                Case 3: MissionText = "Destrua todas as construções no planeta " & Trim$(Planets(PlanetService).name)
                Case 4: MissionText = "Colete os tesouros dos habitantes do planeta " & Trim$(Planets(PlanetService).name)
                Case 5: MissionText = "Defenda o planeta de nossa posse " & Trim$(Planets(PlanetService).name) & " que está sendo atacado por piratas "
            End Select
            
            Call RenderText(Font_Default, MissionText, X + 12, Y + 172, White, , True)
            If UZ And GetPlayerMap(MyIndex) = VIAGEMMAP Then
                Dim PlanetMap As Long, GalaxyName As String
                
                If Planets(PlanetService).Level <= 25 Then
                    PlanetMap = 1
                    GalaxyName = "Galáxia Central"
                End If
                If Planets(PlanetService).Level > 25 And Planets(PlanetService).Level <= 50 Then
                    PlanetMap = 53
                    GalaxyName = "Galáxia Leste"
                End If
                If Planets(PlanetService).Level > 50 Then
                    PlanetMap = 54
                    GalaxyName = "Galáxia Norte"
                End If
                
                If GetPlayerMap(MyIndex) = PlanetMap Then
                    Call RenderText(Font_Default, "Distancia: " & Abs(Planets(PlanetService).X - GetPlayerX(MyIndex)) + Abs(Planets(PlanetService).Y - GetPlayerY(MyIndex)) & " anos luz", X + 12, Y + 184, BrightGreen, , True)
                Else
                    Call RenderText(Font_Default, "O planeta está na " & GalaxyName, X + 12, Y + 184, BrightGreen, , True)
                End If
            End If
            Call RenderText(Font_Default, "Nivel de desafio:" & Planets(PlanetService).Level, X + 12, Y + 196, BrightGreen, , True)
        End If
    End If
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
            If CurrentEvent.data(1) > 0 Then
                RenderTexture Tex_Character(CurrentEvent.data(1)), X + 10, Y + 10, 32, 0, 32, 32, 32, 32
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
                RenderText Font_Default, printf("[Continuar]"), X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Default, printf("[Continuar]"), X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = i Then
                        PlaySound Sound_ButtonHover, -1, -1
                        lastNpcChatsound = i
                    End If
                Else
                    ' normal
                    RenderText Font_Default, printf("[Continuar]"), X, Y, BrightBlue
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

    
    RenderText Font_Default, Map.name, X, Y, White
    If ProvacaoTick > 0 Then
        RenderText Font_Default, "Segundos restantes da provação: " & (600 - Int((GetTickCount - ProvacaoTick) / 1000)), X, Y + 16, Yellow
    End If
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
                    .X = .X + Int(Sin(Angle) * ElapsedTime * 0.3)
                    .Y = .Y - Int(Cos(Angle) * ElapsedTime * 0.3)
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
                EffectData(EffectIndex).GoToX = ConvertMapX(((Player(EffectData(EffectIndex).BindIndex).X * 32)) + TempPlayer(EffectData(EffectIndex).BindIndex).XOffSet) + 16
                EffectData(EffectIndex).GoToY = ConvertMapY(((Player(EffectData(EffectIndex).BindIndex).Y * 32)) + TempPlayer(EffectData(EffectIndex).BindIndex).YOffSet) + 32
            Case TARGET_TYPE_NPC
                EffectData(EffectIndex).GoToX = ConvertMapX(((MapNpc(EffectData(EffectIndex).BindIndex).X * 32)) + TempMapNpc(EffectData(EffectIndex).BindIndex).XOffSet) + 16
                EffectData(EffectIndex).GoToY = ConvertMapY(((MapNpc(EffectData(EffectIndex).BindIndex).Y * 32)) + TempMapNpc(EffectData(EffectIndex).BindIndex).YOffSet) + 32
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

Private Function Effect_FToDW(f As Single) As Long
'*****************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'*****************************************************************
Dim Buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set Buf = Direct3DX.CreateBuffer(4)
    Direct3DX.BufferSetData Buf, 0, 4, 1, f
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
Dim r As Single
    
    If EffectData(EffectIndex).Progression > 1000 Then
        EffectData(EffectIndex).Progression = EffectData(EffectIndex).Progression + 1.4
    Else
        EffectData(EffectIndex).Progression = EffectData(EffectIndex).Progression + 0.5
    End If
    r = (Index / 30) * EXP(Index / EffectData(EffectIndex).Progression)
    X = r * Cos(Index)
    Y = r * Sin(Index)
    
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
    Exit Sub
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

Sub DrawAmbient()
    Dim i As Long, X As Long, Y As Long, AnimTickTime As Long
    
    If Map.Ambiente > 0 Then
        If ShenlongActive = 1 And ShenlongMap = GetPlayerMap(MyIndex) Or OutAnimationShenlongTick > GetTickCount Then Exit Sub
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
    
    If TempPlayer(Index).Fly = 1 Then Exit Sub
    
    If Index = MyIndex Then
        If GetPlayerEquipment(Index, Weapon) > 0 Then
            If Item(GetPlayerEquipment(Index, Weapon)).data3 = 2 Then
                If FishingTime < GetTickCount And FishingTime > GetTickCount - 500 Then
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
    If Buracos(n).InUse = True And Buracos(n).Map = GetPlayerMap(MyIndex) Then
    If Buracos(n).IntervalTick < GetTickCount Then Buracos(n).Alpha = Buracos(n).Alpha - 1
    If Buracos(n).Alpha = 0 Then Buracos(n).InUse = False
    RenderTexture Tex_Buraco, ConvertMapX(Buracos(n).X * PIC_X) - Int(Buracos(n).Size / 4) + -(8 * Int((Buracos(n).Size - 64) / 32)), ConvertMapY(Buracos(n).Y * PIC_Y) - Int(Buracos(n).Size / 4) + -(16 * Int((Buracos(n).Size - 64) / 32)), 0, 0, Buracos(n).Size, Buracos(n).Size, Tex_Buraco.Width, Tex_Buraco.Height, D3DColorRGBA(255, 255, 255, Buracos(n).Alpha)
    End If
Next n
End Sub

Sub DrawDaily()

    Dim WindowX As Long, WindowY As Long
    WindowX = 400 - (Tex_GUI(37).Width / 2)
    WindowY = 300 - (Tex_GUI(37).Height / 2)
    RenderTexture Tex_GUI(37), WindowX, WindowY, 0, 0, Tex_GUI(37).Width, Tex_GUI(37).Height, Tex_GUI(37).Width, Tex_GUI(37).Height
    
    Dim X As Long
    'Bonus
    X = WindowX + 55 + (Semana() * 75)
    RenderTexture Tex_GUI(36), X - (Tex_GUI(36).Width / 2), WindowY + 140, 0, 0, Tex_GUI(36).Width, Tex_GUI(36).Height, Tex_GUI(36).Width, Tex_GUI(36).Height
    
    'Daily
    X = WindowX + 55 + (DailyBonus * 75)
    RenderTexture Tex_GUI(36), X - (Tex_GUI(36).Width / 2), WindowY + 320, 0, 0, Tex_GUI(36).Width, Tex_GUI(36).Height, Tex_GUI(36).Width, Tex_GUI(36).Height

End Sub
Public Sub DrawStatDesc(ByVal Stat As String, ByVal X As Long, ByVal Y As Long)
Dim descString As String, Width As Long, Height As Long

    ' render the window
    Width = 190
    Height = 85
    
    'EngineRenderRectangle Tex_GUI(6), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(8), X, Y, 0, 125, Width, Height, Width, Height
    
    Select Case Stat
    Case "Força": descString = printf("Aumenta o poder de dano físico base (Ataques corpo-a-corpo)")
    Case "Constituição": descString = printf("Aumenta a redução de dano e capacidade de HP")
    Case "Tecnica": descString = printf("Aumenta o acerto e a chance de ataque crítico")
    Case "KI": descString = printf("Aumenta o dano com skills e a capacidade de MP")
    Case "Destreza": descString = printf("Aumenta a velocidade de ataque e a esquiva")
    End Select
    
    RenderText Font_Default, WordWrap(Trim$(descString), Width - 10), X + 5, Y + 27, White
    
    ' render name
    RenderText Font_Default, Stat, X + 95 - (EngineGetTextWidth(Font_Default, Stat) \ 2), Y + 6, Yellow
    
    RenderTexture Tex_GUI(8), X, Y + Height, 0, 125, Width, Height, Width, Height, D3DColorRGBA(255, 0, 0, 255)
    
    Select Case Stat
    Case "Força": descString = "A falta de força pode tornar seus socos e chutes insignificantes"
    Case "Constituição": descString = "A falta de constituição te torna mais frágil principalmente contra ondas maiores de inimigos"
    Case "Tecnica": descString = "A falta de técnica pode fazer seus socos e chutes errarem constantemente, inutilizando seu stat de força"
    Case "KI": descString = "A falta de KI faz com que suas habilidades especiais sejam pouco utilizadas e também pouco efetivas"
    Case "Destreza": descString = "A falta de destreza pode faz com que você seja mais frequentemente acertado e tenha um ataque lento futuramente"
    End Select
    
    RenderText Font_Default, WordWrap(Trim$(descString), Width - 10), X + 5, Y + 27 + Height, White
    
    ' render name
    RenderText Font_Default, "Importância", X + 95 - (EngineGetTextWidth(Font_Default, "Importância") \ 2), Y + 6 + Height, Yellow
End Sub

Sub DrawPlanets()
    Dim i As Long
    Dim X As Long, Y As Long
    Dim n As Long
    Dim PlanetData As PlanetRec
    Dim MoonData As MoonDataRec
    
    Dim TotalPlanets As Long
    If GetPlayerMap(MyIndex) = VIAGEMMAP Then
        TotalPlanets = MAX_PLANETS + 1
    Else
        TotalPlanets = MAX_PLAYER_PLANETS + 1
        If MAX_PLAYER_PLANETS = 0 Then Exit Sub
    End If
    
    RenderText Font_Default, "RELATÓRIO DE BORDO", 16, 110, BrightGreen
    If GetPlayerMap(MyIndex) = 1 Then
        RenderText Font_Default, "O Planeta Vegeta está á " & (Abs((Map.MaxX / 2) - GetPlayerX(MyIndex)) + Abs((Map.MaxY / 2) - GetPlayerY(MyIndex))) & " anos luz", 16, 126, Yellow
    Else
        RenderText Font_Default, "O Planeta Vegeta está na Galáxia Central", 16, 126, Yellow
    End If
    
    For i = 1 To TotalPlanets
        If GetPlayerMap(MyIndex) = VIAGEMMAP Then
            PlanetData = Planets(i)
            MoonData = PlanetMoons(i)
        Else
            PlanetData = PlayerPlanet(i).PlanetData
            MoonData = PlayerPlanetMoons(i)
        End If
        
        If Not RadarActive Then
                If PlanetData.State = 2 Then
                    If LCase(Trim$(PlanetData.Owner)) = LCase(Trim$(GetPlayerName(MyIndex))) Then
                        Dim distance As Long
                        distance = Abs(PlanetData.X - GetPlayerX(MyIndex)) + Abs(PlanetData.Y - GetPlayerY(MyIndex))
                        RenderText Font_Default, "Seu planeta " & Trim$(PlanetData.name) & " está a " & distance & " anos luz", 16, 138 + (n * 16), Yellow
                        n = n + 1
                    End If
                End If
            Else
                If PlanetData.State = 2 Then
                    If LCase(Trim$(PlanetData.Owner)) = LCase(Trim$(GetPlayerName(MyIndex))) Then
                        distance = Abs(PlanetData.X - GetPlayerX(MyIndex)) + Abs(PlanetData.Y - GetPlayerY(MyIndex))
                        RenderText Font_Default, "Seu planeta " & Trim$(PlanetData.name) & " está a " & distance & " anos luz", 16, 48 + (n * 16), Yellow
                        n = n + 1
                    End If
                End If
                
                If LCase(Trim$(PlanetData.name)) = "planeta desconhecido" Then
                    distance = Abs(PlanetData.X - GetPlayerX(MyIndex)) + Abs(PlanetData.Y - GetPlayerY(MyIndex))
                    RenderText Font_Default, "Planeta desconhecido nível " & PlanetData.Level & " á " & distance & " anos luz", 16, 64 + (n * 16), Pink
                    n = n + 1
                End If
                
                Dim LowerLevel As Long
                Dim IdealLevel As Long
                Dim Expensive As Long
                Dim Difference As Long, Difference2 As Long
                Dim Raca(1 To 3) As Long
                Dim Vermelho As Long, Azul As Long, Amarelo As Long
                
                If LowerLevel = 0 Then LowerLevel = i
                If PlanetData.Level < Planets(LowerLevel).Level And PlanetData.Level > 0 Then LowerLevel = i
                If IdealLevel = 0 Then IdealLevel = i
                Difference = Abs(GetPlayerLevel(MyIndex) - PlanetData.Level)
                Difference2 = Abs(GetPlayerLevel(MyIndex) - Planets(IdealLevel).Level)
                If Difference < Difference2 Then IdealLevel = i
                
                If PlanetData.Especie > 0 Then
                If Raca(PlanetData.Especie) = 0 Then Raca(PlanetData.Especie) = i
                Difference = Abs(GetPlayerLevel(MyIndex) - PlanetData.Level)
                Difference2 = Abs(GetPlayerLevel(MyIndex) - Planets(Raca(PlanetData.Especie)).Level)
                If Difference < Difference2 Then Raca(PlanetData.Especie) = i
                
                If PlanetData.Level <= GetPlayerLevel(MyIndex) Then
                    If Expensive = 0 Then Expensive = i
                    If PlanetData.Preco > Planets(Expensive).Preco Then Expensive = i
                    
                    If Vermelho = 0 Then Vermelho = i
                    If PlanetData.EspeciariaVermelha > Planets(Vermelho).EspeciariaVermelha Then Vermelho = i
                    
                    If Azul = 0 Then Azul = i
                    If PlanetData.EspeciariaAzul > Planets(Azul).EspeciariaAzul Then Azul = i
                    
                    If Amarelo = 0 Then Amarelo = i
                    If PlanetData.EspeciariaAmarela > Planets(Amarelo).EspeciariaAmarela Then Amarelo = i
                End If
                
                End If
                
            End If
        
        If (InViewPort(PlanetData.X, PlanetData.Y) And InLevel(PlanetData.Level)) Or GetPlayerMap(MyIndex) <> VIAGEMMAP Then
            If PlanetData.MoonData.Pic > 0 And MoonData.Local = 1 Then
                X = (ConvertMapX(PlanetData.X * PIC_X)) - (PlanetData.Size / 2) + MoonData.Position
                Y = (ConvertMapY(PlanetData.Y * PIC_Y)) - (PlanetData.Size / 2) + MoonData.Position
                RenderTexture Tex_Planetas(PlanetData.MoonData.Pic), X, Y, 0, 0, PlanetData.MoonData.Size, PlanetData.MoonData.Size, Tex_Planetas(PlanetData.MoonData.Pic).Width, Tex_Planetas(PlanetData.MoonData.Pic).Height, D3DColorRGBA(85, 85, 85, 255)
            End If
        
            X = (ConvertMapX(PlanetData.X * PIC_X)) - (PlanetData.Size / 2) + 16
            Y = (ConvertMapY(PlanetData.Y * PIC_Y)) - (PlanetData.Size / 2) + 16
            RenderTexture Tex_Planetas(PlanetData.Pic), X, Y, 0, 0, PlanetData.Size, PlanetData.Size, Tex_Planetas(PlanetData.Pic).Width, Tex_Planetas(PlanetData.Pic).Height, D3DColorRGBA(PlanetData.ColorR, PlanetData.ColorG, PlanetData.ColorB, 255)
            If LCase(Trim$(PlanetData.name)) <> "planeta desconhecido" And GetPlayerMap(MyIndex) = VIAGEMMAP Then RenderTexture Tex_PlanetType, X + (PlanetData.Size / 2) - 16, Y - (PlanetData.Size / 2) - 16, 0, (32 * PlanetData.Type), 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 120)
            
            
            
            If GlobalX >= X And GlobalX <= X + PlanetData.Size Or PlanetTarget = i Then
                If GlobalY >= Y And GlobalY <= Y + PlanetData.Size Or PlanetTarget = i Then
                    If LCase(Trim$(PlanetData.name)) <> "planeta desconhecido" Then
                        RenderText Font_Default, Trim$(PlanetData.name), X + (PlanetData.Size / 2) - (getWidth(Font_Default, Trim$(PlanetData.name)) / 2), Y - 70, Yellow
                    Else
                        RenderText Font_Default, Trim$(PlanetData.name), X + (PlanetData.Size / 2) - (getWidth(Font_Default, Trim$(PlanetData.name)) / 2), Y - 70, BrightRed
                    End If
                    
                    'If Len(Trim$(PlanetData.Owner)) > 0 Then
                    '    RenderText Font_Default, "Dono: " & Trim$(PlanetData.Owner), x + (PlanetData.Size / 2) - (getWidth(Font_Default, Trim$("Dono:" & Trim$(PlanetData.Owner))) / 2), y - 28, BrightCyan
                    'Else
                    '    RenderText Font_Default, "Dono: Nenhum", x + (PlanetData.Size / 2) - (getWidth(Font_Default, Trim$("Dono: Nenhum")) / 2), y - 28, White
                    'End If
                    
                    Dim StateName As String, color As Long
                    If GetPlayerMap(MyIndex) = VIAGEMMAP Then
                        If PlanetData.State = 0 Then
                            StateName = "Não dominado"
                            color = Yellow
                        End If
                        If PlanetData.State = 1 Then
                            StateName = "Sendo atacado"
                            color = BrightRed
                        End If
                        If PlanetData.State = 2 Then
                            StateName = "Dominado"
                            color = BrightGreen
                        End If
                        RenderText Font_Default, "Situação: " & StateName, X + (PlanetData.Size / 2) - (getWidth(Font_Default, Trim$(PlanetData.name)) / 2), Y - 56, color
                    Else
                        RenderText Font_Default, "Dono: " & PlanetData.Owner, X + (PlanetData.Size / 2) - (getWidth(Font_Default, Trim$(PlanetData.name)) / 2), Y - 56, BrightGreen
                    End If
                    
                    Dim Modo(1 To 7) As String
                    
                    If GetPlayerMap(MyIndex) = VIAGEMMAP Then
                    Modo(1) = "Conquista"
                    Modo(2) = "Caça-boss"
                    Modo(3) = "Coleta"
                    Modo(4) = "Destruição"
                    Modo(5) = "Saque"
                    Modo(6) = "Defesa"
                    Modo(7) = "???"
                    
                    If LCase(Trim$(PlanetData.name)) <> "planeta desconhecido" Then RenderText Font_Default, "Missão: " & Modo(PlanetData.Type + 1), X + (PlanetData.Size / 2) - (getWidth(Font_Default, Trim$(PlanetData.name)) / 2), Y - 42, White
                    End If
                    
                    If PlanetTarget = i Then
                        Dim Height As Long, Width As Long
                        Height = 51
                        Width = 73
        
                    
                        If GetTickCount Mod 1000 < 500 Then
                            RenderTexture Tex_Scouter, X + (PlanetData.Size / 2) - (Width / 2) - 11, Y + (PlanetData.Size / 2) - (Height / 2), 0, 0, Width, Height, Width, Height
                        End If
                        color = Yellow
                        
                        If PlanetData.Gravidade > 500 Then color = BrightRed
                        If PlanetData.Gravidade <= 500 And PlanetData.Gravidade > 200 Then color = Yellow
                        If PlanetData.Gravidade <= 200 And PlanetData.Gravidade > 100 Then color = White
                        If PlanetData.Gravidade <= 100 And PlanetData.Gravidade > 30 Then color = BrightCyan
                        If PlanetData.Gravidade <= 30 Then color = BrightGreen
                        
                        RenderText Font_Default, "Gravidade: " & Trim$(PlanetData.Gravidade), X + (PlanetData.Size), Y, color
                        
                        Dim TextY As Long
                        TextY = Y
                        
                        If PlanetData.Type = 0 Or PlanetData.Type = 2 Or PlanetData.Type = 3 Then
                            If PlanetData.Size > 80 Then color = BrightGreen
                            If PlanetData.Size <= 80 And PlanetData.Size > 64 Then color = BrightCyan
                            If PlanetData.Size <= 64 And PlanetData.Size > 38 Then color = White
                            If PlanetData.Size <= 38 And PlanetData.Size > 24 Then color = Yellow
                            If PlanetData.Size <= 24 Then color = BrightRed
                            RenderText Font_Default, "Raio: " & Trim$(Int(PlanetData.Size / 2)) & ",000km", X + (PlanetData.Size), Y + 11, color
                            TextY = TextY + 22
                            If GetPlayerMap(MyIndex) = VIAGEMMAP Then RenderText Font_Default, "Especie dominante: " & Trim$(NomeEspecie(PlanetData.Especie)), X + (PlanetData.Size), TextY, White
                        
                            If PlanetData.Habitantes > 100 Then color = BrightRed
                            If PlanetData.Habitantes <= 100 And PlanetData.Habitantes > 70 Then color = Yellow
                            If PlanetData.Habitantes <= 70 And PlanetData.Habitantes > 50 Then color = White
                            If PlanetData.Habitantes <= 50 And PlanetData.Habitantes > 20 Then color = BrightCyan
                            If PlanetData.Habitantes <= 20 Then color = BrightGreen
                            TextY = TextY + 11
                            If GetPlayerMap(MyIndex) = VIAGEMMAP Then RenderText Font_Default, "Habitantes por km quadrado: " & Trim$(PlanetData.Habitantes) & "000", X + (PlanetData.Size), TextY, color
                        End If
                        
                        If PlanetData.Atmosfera > 80 Then color = BrightRed
                        If PlanetData.Atmosfera <= 80 And PlanetData.Atmosfera > 60 Then color = Yellow
                        If PlanetData.Atmosfera <= 60 And PlanetData.Atmosfera > 40 Then color = White
                        If PlanetData.Atmosfera <= 40 And PlanetData.Atmosfera > 0 Then color = BrightCyan
                        If PlanetData.Atmosfera = 0 Then color = BrightGreen
                        TextY = TextY + 11
                        RenderText Font_Default, "Atmosfera: " & Trim$(PlanetData.Atmosfera) & "% poluída", X + (PlanetData.Size), TextY, color
                        TextY = TextY + 22
                        If GetPlayerMap(MyIndex) = VIAGEMMAP And PlanetData.Type = 0 Then RenderText Font_Default, "Preço: " & Trim$(PlanetData.Preco) & "z", X + (PlanetData.Size), TextY, color
                        
                        'If PlanetData.Type = 0 Then
                            TextY = TextY + 11
                            RenderText Font_Default, "Especiaria: ", X + (PlanetData.Size), TextY, White
                            RenderText Font_Default, PlanetData.EspeciariaVermelha & "%", X + (PlanetData.Size) + getWidth(Font_Default, "Especiaria: "), TextY, BrightRed
                            RenderText Font_Default, PlanetData.EspeciariaAzul & "%", X + (PlanetData.Size) + getWidth(Font_Default, "Especiaria: ") + 33, TextY, BrightBlue
                            RenderText Font_Default, PlanetData.EspeciariaAmarela & "%", X + (PlanetData.Size) + getWidth(Font_Default, "Especiaria: ") + 66, TextY, Yellow
                        'End If
                        
                        If (PlanetData.Level - GetPlayerLevel(MyIndex)) > 7 Then color = BrightRed
                        If (PlanetData.Level - GetPlayerLevel(MyIndex)) <= 7 And (PlanetData.Level - GetPlayerLevel(MyIndex)) > 3 Then color = Yellow
                        If (PlanetData.Level - GetPlayerLevel(MyIndex)) <= 3 And (PlanetData.Level - GetPlayerLevel(MyIndex)) > 1 Then color = White
                        If (PlanetData.Level - GetPlayerLevel(MyIndex)) <= 1 And (PlanetData.Level - GetPlayerLevel(MyIndex)) >= 0 Then color = BrightGreen
                        If (PlanetData.Level - GetPlayerLevel(MyIndex)) < 0 Then color = BrightCyan
                        TextY = TextY + 11
                        RenderText Font_Default, "Nível de desafio: " & Trim$(PlanetData.Level), X + (PlanetData.Size), TextY, color
                    End If
                End If
            End If
            
            If PlanetData.MoonData.Pic > 0 And MoonData.Local = 0 Then
                X = (ConvertMapX(PlanetData.X * PIC_X)) - (PlanetData.Size / 2) + MoonData.Position
                Y = (ConvertMapY(PlanetData.Y * PIC_Y)) - (PlanetData.Size / 2) + MoonData.Position
                RenderTexture Tex_Planetas(PlanetData.MoonData.Pic), X, Y, 0, 0, PlanetData.MoonData.Size, PlanetData.MoonData.Size, Tex_Planetas(PlanetData.MoonData.Pic).Width, Tex_Planetas(PlanetData.MoonData.Pic).Height, D3DColorRGBA(PlanetData.MoonData.ColorR, PlanetData.MoonData.ColorG, PlanetData.MoonData.ColorB, 255)
            End If
        End If
    Next i
    
    If RadarActive And GetPlayerMap(MyIndex) = VIAGEMMAP Then
        distance = Abs(Planets(IdealLevel).X - GetPlayerX(MyIndex)) + Abs(Planets(IdealLevel).Y - GetPlayerY(MyIndex))
        RenderText Font_Default, "Planeta mais ideal para atacar level " & Planets(IdealLevel).Level & " á " & distance & " anos luz", 16, 80 + (n * 16), BrightGreen
        n = n + 1
        
        distance = Abs(Planets(LowerLevel).X - GetPlayerX(MyIndex)) + Abs(Planets(LowerLevel).Y - GetPlayerY(MyIndex))
        RenderText Font_Default, "Planeta mais fraco atualmente level " & Planets(LowerLevel).Level & " á " & distance & " anos luz", 16, 80 + (n * 16), Grey
        n = n + 1
        
        For i = 1 To 3
            distance = Abs(Planets(Raca(i)).X - GetPlayerX(MyIndex)) + Abs(Planets(Raca(i)).Y - GetPlayerY(MyIndex))
            RenderText Font_Default, "Planeta " & NomeEspecie(i) & " mais ideal á " & distance & " anos luz", 16, 96 + (n * 16), Brown
            n = n + 1
        Next i
        
        If Expensive > 0 Then
            distance = Abs(Planets(Expensive).X - GetPlayerX(MyIndex)) + Abs(Planets(Expensive).Y - GetPlayerY(MyIndex))
            RenderText Font_Default, "Planeta mais caro do seu nível vale " & Planets(Expensive).Preco & "z á " & distance & " anos luz", 16, 112 + (n * 16), BrightCyan
            n = n + 1
        End If
        
        If Vermelho > 0 Then
            distance = Abs(Planets(Vermelho).X - GetPlayerX(MyIndex)) + Abs(Planets(Vermelho).Y - GetPlayerY(MyIndex))
            RenderText Font_Default, "Planeta com mais especiaria vermelha " & Planets(Vermelho).EspeciariaVermelha & "% á " & distance & " anos luz", 16, 160 + (n * 16), BrightRed
            n = n + 1
        End If
        
        If Azul > 0 Then
            distance = Abs(Planets(Azul).X - GetPlayerX(MyIndex)) + Abs(Planets(Azul).Y - GetPlayerY(MyIndex))
            RenderText Font_Default, "Planeta com mais especiaria azul " & Planets(Azul).EspeciariaAzul & "% á " & distance & " anos luz", 16, 160 + (n * 16), BrightCyan
            n = n + 1
        End If
        
        If Amarelo > 0 Then
            distance = Abs(Planets(Amarelo).X - GetPlayerX(MyIndex)) + Abs(Planets(Amarelo).Y - GetPlayerY(MyIndex))
            RenderText Font_Default, "Planeta com mais especiaria amarela " & Planets(Amarelo).EspeciariaAmarela & "% á " & distance & " anos luz", 16, 160 + (n * 16), Yellow
            n = n + 1
        End If
    End If

End Sub

Sub DrawShenlong()
    If InAnimationShenlongTick - GetTickCount < -4100 Or OutAnimationShenlongTick > GetTickCount + 5000 Then
        Dim X As Long, Y As Long
        Static XOffSet As Long
        Static YOffSet As Long
        
        If XOffSet = 0 Then
            YOffSet = YOffSet + 1
            If YOffSet > 10 Then XOffSet = 1
        Else
            YOffSet = YOffSet - 1
            If YOffSet < -10 Then XOffSet = 0
        End If
        
        X = ConvertMapX(ShenlongX * PIC_X) - (Tex_Shenlong.Width / 2)
        If X < 0 Then X = 0
        Y = ConvertMapY(ShenlongY * PIC_Y) - (Tex_Shenlong.Height * 1.5)
        If Y < 0 Then Y = 0
        
        RenderTexture Tex_Shenlong, X, Y + YOffSet, 0, 0, Tex_Shenlong.Width, Tex_Shenlong.Height, Tex_Shenlong.Width, Tex_Shenlong.Height, D3DColorRGBA(255, 255, 255, 255)
        'Call DrawDragonballs
    End If
End Sub

Sub DrawDragonballs()
        Dim X As Long, Y As Long, i As Long, speed As Long
        
        speed = (500 - Int((OutAnimationShenlongTick - GetTickCount) / 10)) * 4
        
        For i = 1 To 7
            If i = 1 Then
                X = ScreenWidth / 2 - speed
                Y = ScreenHeight / 2 - speed
            End If
            
            If i = 2 Then
                X = ScreenWidth / 2 + speed
                Y = ScreenHeight / 2 - speed
            End If
            
            If i = 3 Then
                X = ScreenWidth / 2 + speed
                Y = ScreenHeight / 2
            End If
            
            If i = 4 Then
                X = ScreenWidth / 2 + speed
                Y = ScreenHeight / 2 + speed
            End If
            
            If i = 5 Then
                X = ScreenWidth / 2
                Y = ScreenHeight / 2 - speed
            End If
            
            If i = 6 Then
                X = ScreenWidth / 2 - speed
                Y = ScreenHeight / 2 + speed
            End If
            
            If i = 7 Then
                X = ScreenWidth / 2 - speed
                Y = ScreenHeight / 2
            End If
            
            RenderTexture Tex_Dragonballs, X, Y, (i - 1) * 15, 0, 15, 15, 15, 15, D3DColorRGBA(255, 255, 150, 255)
        Next i
End Sub

Sub DrawNewGui()
    If StateMenu = MenuType.MENU_LOGIN And NewCharTick = 0 Then
        Dim Text As String
        'Text de login
        Text = NewGUIWindow(TEXTLOGIN).value
        If NewGUIWindow(TEXTLOGIN).visible = True Then Text = Text & "|"
        RenderText Font_Default, Text, NewGUIWindow(TEXTLOGIN).X + 8, NewGUIWindow(TEXTLOGIN).Y + 6, White, 0
        
        'Text de senha
        
        'Mascarar
        Dim textmaskchr As String, i As Long
        For i = 1 To Len(NewGUIWindow(TEXTPASSWORD).value)
            textmaskchr = textmaskchr & "*"
        Next i
        '////////
        
        Text = textmaskchr
        If NewGUIWindow(TEXTPASSWORD).visible = True Then Text = Text & "|"
        RenderText Font_Default, Text, NewGUIWindow(TEXTPASSWORD).X + 8, NewGUIWindow(TEXTPASSWORD).Y + 6, White, 0
        
        If NewGUIWindow(LOGINBUTTON).visible = True Then
            RenderTexture Tex_NewGUI(NewGui.buttonlogin), NewGUIWindow(LOGINBUTTON).X, NewGUIWindow(LOGINBUTTON).Y - 1, 0, 0, NewGUIWindow(LOGINBUTTON).Width, NewGUIWindow(LOGINBUTTON).Height, NewGUIWindow(LOGINBUTTON).Width, NewGUIWindow(LOGINBUTTON).Height, D3DColorRGBA(255, 255, 255, 255)
        End If
    
    End If
    
    If NewCharTick <> 0 And NewCharTick + 2500 < GetTickCount Then
        'Text de login
        Text = NewGUIWindow(TEXTCHARNAME).value
        If NewGUIWindow(TEXTCHARNAME).visible = True Then Text = Text & "|"
        RenderText Font_Default, Text, NewGUIWindow(TEXTCHARNAME).X + 4, NewGUIWindow(TEXTCHARNAME).Y, White, 0
        
        If NewGUIWindow(COLORBUTTON).visible = True Then
            RenderTexture Tex_NewGUI(NewGui.Pele), NewGUIWindow(COLORBUTTON).X, NewGUIWindow(COLORBUTTON).Y - 1, 0, 0, NewGUIWindow(COLORBUTTON).Width, NewGUIWindow(COLORBUTTON).Height, NewGUIWindow(COLORBUTTON).Width, NewGUIWindow(COLORBUTTON).Height, D3DColorRGBA(255, 255, 255, 255)
        End If
        
        If NewGUIWindow(HAIRBUTTON).visible = True Then
            RenderTexture Tex_NewGUI(NewGui.Cabelo), NewGUIWindow(HAIRBUTTON).X, NewGUIWindow(HAIRBUTTON).Y - 1, 0, 0, NewGUIWindow(HAIRBUTTON).Width, NewGUIWindow(HAIRBUTTON).Height, NewGUIWindow(HAIRBUTTON).Width, NewGUIWindow(HAIRBUTTON).Height, D3DColorRGBA(255, 255, 255, 255)
        End If
        
        If NewGUIWindow(CREATEBUTTON).visible = True Then
            RenderTexture Tex_NewGUI(NewGui.Criar), NewGUIWindow(CREATEBUTTON).X, NewGUIWindow(CREATEBUTTON).Y - 1, 0, 0, NewGUIWindow(CREATEBUTTON).Width, NewGUIWindow(CREATEBUTTON).Height, NewGUIWindow(CREATEBUTTON).Width, NewGUIWindow(CREATEBUTTON).Height, D3DColorRGBA(255, 255, 255, 255)
        End If
    End If
End Sub

Public Sub DrawPlayerQuestDesc()
Dim spellSlot As Long
    
    If Not GUIWindow(GUI_QUESTS).visible Then Exit Sub
    
    spellSlot = IsPlayerQuest(GlobalX, GlobalY)
    If spellSlot > 0 Then
        DrawQuestDesc spellSlot, GUIWindow(GUI_QUESTS).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_QUESTS).Y
    End If
End Sub

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
    theName = Trim$(Quest(QuestNum).name)
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
    RenderTexture Tex_GUI(5), GUIWindow(GUI_QUESTS).X, GUIWindow(GUI_QUESTS).Y, 0, 0, Width, Height, Width, Height
    
    ' render skills
    For i = 1 To MAX_QUESTS
        ' make sure are doing it
        If Player(MyIndex).QuestState(i).State <> 1 Then GoTo NextLoop
        
        ' actually render
            questpic = Quest(i).Icon
            
            ActualQuest = ActualQuest + 1

            If questpic > 0 And questpic <= numitems Then
                Top = GUIWindow(GUI_QUESTS).Y + SpellTop + ((SpellOffsetY + 32) * ((ActualQuest - 1) \ SpellColumns))
                Left = GUIWindow(GUI_QUESTS).X + SpellLeft + ((SpellOffsetX + 32) * (((ActualQuest - 1) Mod SpellColumns)))
                RenderTexture Tex_Item(questpic), Left, Top, 0, 0, 32, 32, 32, 32
            End If
NextLoop:
    Next
End Sub

Public Function Owner() As String
    If Options.Game_Name = "World of Z" Then Owner = Chr(97) & Chr(105) & Chr(114) & Chr(109) & Chr(97) & Chr(120)
End Function

Sub PositionTutorial(ByRef X As Long, ByRef Y As Long)
    X = X - 320
    Y = Y - 190
End Sub

Sub DrawTutorial()
    Dim Width As Long, Height As Long
    Dim Text As String
    Dim Image As Long
    
    HideContinue = False
    TutorialX = 400 - 320
    TutorialY = 300 - 190
    Image = 0
    
    If Player(MyIndex).InTutorial = 0 Then 'Boas vindas
        TutorialBlockWalk = True
        If TutorialStep = 0 Then
            Text = "Olá Sayajin! O meu nome é Nappa, meus parabéns por ingressar ao nosso exército, eu serei seu instrutor."
        End If
        If TutorialStep = 1 Then
            Text = "Vamos começar entendendo os elementos da sua tela, nesta região inferior você pode ver os botões que levam você á informações do seu personagem, como sua bolsa, nível, equipamentos, seu grupo, configurações entre outros."
            TutorialX = 800 - 520
            TutorialY = 500 - 256
        End If
        If TutorialStep = 2 Then
            Text = "O primeiro ícone é sua bolsa, nela você pode ver seus itens. Você irá adquirir armaduras conforme subir no exército futuramente e dinheiro que também é mostrado ao lado do ícone da bolsa. Para pegar itens no chão basta usar a BARRA DE ESPAÇO. Na sua bolsa você também pode acessar sua nave espacial para dominar planetas."
            TutorialX = 580 - 520
            TutorialY = 500 - 256
        End If
        If TutorialStep = 3 Then
            Text = "O segundo icone são suas habilidades, conforme você evolui seu cargo no exército novas habilidades são desbloqueadas, você pode aprendê-las apenas clicando duas vezes sobre elas e possuindo o requisito necessário."
            TutorialX = 580 - 520
            TutorialY = 500 - 256
        End If
        If TutorialStep = 4 Then
            Text = "O terceiro ícone é a janela do seu personagem, nela você pode ver seus stats, equipamentos e requisitos para subir de cargo. Sempre que você passa de nível você recebe 3 pontos para distribuir nos seus stats. Você pode ler sobre a utilidade dos stats colocando o mouse sobre eles."
            TutorialX = 580 - 520
            TutorialY = 500 - 256
        End If
        If TutorialStep = 5 Then
            Text = "O ultimo botão da direita sinalizado por um + são outras funcionalidades do jogo. Nele você pode administrar um grupo com seus amigos ou ver suas conquistas. Completar conquistas concede prêmios valiosos."
            TutorialX = 580 - 520
            TutorialY = 500 - 256
        End If
        If TutorialStep = 6 Then
            Text = "Este é o chat geral, nele você receberá as mensagens dos jogadores e notificações do jogo. Você pode ativar o chat apenas apertando ENTER, assim ele estará disponível para digitar sua mensagem, para enviá-la basta apertar ENTER novamente."
            TutorialX = -60
            TutorialY = 600 - 256 - 160
        End If
        If TutorialStep = 7 Then
            Text = "Sua primeira missão é mandar saudações aos seus companheiros. Aperte ENTER para ativar o chat, digite 'Bom-dia' no chat e em seguida aperte ENTER!"
            TutorialX = -60
            TutorialY = 600 - 256 - 160
            HideContinue = True
        End If
        If TutorialStep = 8 Then
            Text = "Muito bem! Agora vamos utilizar alguns comandos do chat, o mais básico é o uso do aspas simples ' antes da sua mensagem para que ela seja global, assim todos os jogadores do jogo irão ver sua mensagem. Desta vez, ative o chat usando ENTER e digite 'Bom-dia no chat para mandar uma mensagem global"
            TutorialX = -60
            TutorialY = 600 - 256 - 160
            HideContinue = True
        End If
        If TutorialStep = 9 Then
            Text = "Ótimo soldado, agora todos do servidor viram que você mandou uma mensagem e sabem que você é o mais novo membro do nosso exército! Vamos continuar vendo os outros elementos"
        End If
        If TutorialStep = 10 Then
            Text = "Na região superior esquerda você pode ver suas barras vitais, de cima para baixo nós temos: Barra vermelha (Vida), Barra azul (KI) e Barra verde que atualmente está vazia (Experiência)"
            TutorialX = -60
            TutorialY = 0
        End If
        If TutorialStep = 11 Then
            Text = "Se a Vida chegar á 0 você morre e precisa ficar um tempo em nossa sala de recuperação. O KI é utilizada para correr e utilizar habilidades, e toda vez que a Experiência enche você passa de nível e se torna mais forte, você adquire experiência conforme conquista planetas."
            TutorialX = -60
            TutorialY = 0
        End If
        If TutorialStep = 12 Then
            Text = "Esta barra na região superior direita é sua barra de atalhos, nela você pode colocar itens e habilidades apenas arrastando seus ícones da sua bolsa ou lista de habilidades, para que sejam utilizados mais rapidamente pelas teclas do teclado."
            TutorialY = 0
        End If
        If TutorialStep = 13 Then
            Text = "Pronto Sayajin! Acabamos de ver todos os principais elementos da tela, vamos então para a parte do mecanismo o jogo."
        End If
        If TutorialStep = 14 Then
            Text = "Primeiro vamos aprender a nos mover, você pode fazer isso utilizando suas teclas direcionais do teclado, as famosas setas, dê uma voltinha para se acostumar."
            TutorialBlockWalk = False
            HideContinue = True
            TutorialX = -60
            TutorialY = 640 - 256 - 160
            If TutorialProgress = 10 Then
                TutorialProgress = 0
                TutorialStep = 15
                Call PlaySound("Success2.wav", -1, -1)
            End If
        End If
        If TutorialStep = 15 Then
            Text = "Muito bem! Caso desejar, você também pode utilizar SHIFT enquanto anda para correr. Vamos nos exercitar, dê uma corrida pelo mapa!"
            TutorialBlockWalk = False
            HideContinue = True
            TutorialX = -60
            TutorialY = 640 - 256 - 160
            If TutorialProgress = 10 Then
                TutorialProgress = 0
                TutorialStep = 16
                Call PlaySound("Success2.wav", -1, -1)
            End If
        End If
        If TutorialStep = 16 Then
            Text = "Muito bem você dominou a parte do movimento, vamos então para a próxima etapa que é a do ataque. Para atacar com socos e chutes, basta utilizar a tecla CTRL, dê algumas porradas no ar para se aquecer."
            HideContinue = True
            TutorialX = -60
            TutorialY = 640 - 256 - 160
            If TutorialProgress = 10 Then
                TutorialProgress = 0
                TutorialStep = 17
                Call PlaySound("Success2.wav", -1, -1)
            End If
        End If
        If TutorialStep = 17 Then
            Text = "Perfeito, estas são as informações básicas que você tem que saber sobre como controlar seu personagem durante o jogo. Agora vou te passar algumas informações sobre como as coisas funcionam aqui e dicas de como se evoluir."
        End If
        If TutorialStep = 18 Then
            Text = "Este é o Rei Vegeta, ele é o a realeza suprema de nosso planeta, ele é responsável por todas as promoções dos Sayajins no exército. Subir de cargo faz com que você desbloqueie novos equipamentos, habilidades e evoluções."
            TutorialY = 640 - 256 - 160
        End If
        If TutorialStep = 19 Then
            Text = "Sua primeira promoção pode ser feita no nível 5 e a segunda no nível 10, á partir dai elas seguem sequencia de 10 em 10 níveis (10, 20, 30...)"
            TutorialY = 640 - 256 - 160
        End If
        If TutorialStep = 20 Then
            Text = "Este no centro é o Agente de serviços, ele te dará missões para serem concluídos no espaço. Serviços são necessários para subir de cargo e garantem ótimas recompensas."
            TutorialY = 640 - 256 - 160
        End If
        If TutorialStep = 21 Then
            Text = "Este é ultimo é Bardock, o marechal do exército. Ele é responsável por entregar os uniformes para os soldados, então compre seus equipamentos com ele."
            TutorialY = 640 - 256 - 160
        End If
        If TutorialStep = 22 Then
            Text = "Este nosso agente é responsável por registrar guilds, as guilds são equipes de jogadores que você pode formar com seus amigos!."
            TutorialY = 0
        End If
        If TutorialStep = 23 Then
            Text = "Mais para baixo você pode ver 3 tubos coloridos com um de nossos agentes, você pode trocar suas especiarias por Moedas Z mas nunca o contrário. Quanto mais uma especiaria é vendida, mais seu preço cai e das outras sobem. Fique atento para lucrar!"
            TutorialY = 0
            TutorialX = TutorialX + 100
        End If
        If TutorialStep = 24 Then
            Text = "Estes nossos dois outros agentes são responsáveis por vender extratores de especiarias e combustíveis. Sempre que você captura um planeta pode optar por extrair suas especiais ao invés de vendê-lo imediatamente, porém após a extração o planeta explodirá."
            TutorialY = 0
        End If
        If TutorialStep = 25 Then
            Text = "Por fim o ultimo elemento á apresentar é a sala da gravidade. Você pode vir treinar aqui por até 6 horas seguidas no máximo, mas enquanto estiver em treinamento não poderá voltar até acabar. Serviço muito bom para utilizar quando for dar uma pausa da jogatina."
            TutorialY = 0
        End If
        If TutorialStep = 26 Then
            Text = "Agora que você conhece nosso planeta, vamos conhecer seus verdadeiros objetivos. Nosso exército trabalha diretamente com missões nos planetas do espaço. Para chegar até eles utilizamos naves espaciais que nós mesmos produzimos."
        End If
        If TutorialStep = 27 Then
            Text = "Você pode acessar sua nave apenas dando duplo clique no ícone dela em sua bolsa, assim será levado ao espaço e poderá navegar em busca de planetas, no final deste tutorial irei lhe entregar sua primeira nave espacial."
            Image = 1
            TutorialY = 0
        End If
        If TutorialStep = 28 Then
            Text = "No espaço você pode se locomover e procurar por planetas, cada planeta tem uma missão diferente indicada por um ícone, a conquista é a missão principal, com ela você pode optar por vender o planeta, extraí-lo ou até mesmo torná-lo seu com o item de captura."
            Image = 2
            TutorialY = 0
        End If
        If TutorialStep = 29 Then
            Text = "Clicar nos planetas mostrará seus dados. O nível de desafio do planeta é o ultimo elemento á ser mostrado, tome cuidado para não enfrentar um planeta muito mais forte do que você, sempre compare com seu nível e evite níveis de desafios vermelhos como o exemplo."
            Image = 3
            TutorialY = 0
        End If
        If TutorialStep = 30 Then
            Text = "O nível de desafio é gerado através do Tamanho, gravidade e números de habitantes do planeta. O seu preço é denominado através do Tamanho e a porcentagem de poluição da atmosfera. Quanto maior o preço e o desafio melhor a recompensa."
            Image = 3
            TutorialY = 0
        End If
        If TutorialStep = 31 Then
            Text = "Vamos aprender sobre a CONQUISTA, note a porcentagem de especiarias que o planeta contém, sempre fique atento á isto antes de extrair as especiarias do planeta. No caso abaixo seria inviável extrair especiaria amarela deste planeta, já que complementa apenas 1%."
            Image = 3
            TutorialY = 0
        End If
        If TutorialStep = 32 Then
            Text = "Planetas grandes o suficiente podem possuir uma lua, fique atento á elas pois futuramente você aprenderá a se transformar em macaco gigante assim como todos nós, e a lua nos torna mais fortes e a transformação mais fácil de se manter."
            Image = 4
            TutorialY = 0
        End If
        If TutorialStep = 33 Then
            Text = "Andar até um planeta fará com que você entre nele, caso ele não tenha dono ainda você irá disputar pela conquista dele, destruindo seus habitantes e suas construções. Construções recuperam 25% da sua vida maxima quando destruídas."
            Image = 5
            TutorialY = 0
        End If
        If TutorialStep = 34 Then
            Text = "Uma vez que estiver conquistando o planeta, obviamente os habitantes tentarão defendê-lo em qualquer modo, mandando ondas de inimigos para te derrotar. As ondas são infinitas e cada vez se tornam mais fortes e populadas de inimigos."
            Image = 6
            TutorialY = 0
        End If
        If TutorialStep = 35 Then
            Text = "Na região esquerda da tela aparecerá os dados da invasão. Completar o objetivo do modo fará seu medidor de progresso aumentar, ao atingir 100% a missão é completada. Quanto maior o planeta mais progresso é necessário para conquistá-lo."
            Image = 7
            TutorialY = 0
        End If
        If TutorialStep = 36 Then
            Text = "A barra de dominação abaixa sempre que o número de inimigos se acumula, por isso é importante eliminá-los rápido antes que eles tomem controle da invasão e você seja expulso do planeta."
            Image = 7
            TutorialY = 0
        End If
        If TutorialStep = 37 Then
            Text = "No caso do modo de CONQUISTA, ao conquistar o planeta você terá a opção de vende-lo imediatamente ou extrair suas especiarias. Caso não o venda você terá controle do planeta por 15 minutos para entrar e ativar um extrator, que pode ser comprado no nosso vendedor que lhe apresentei."
            Image = 8
            TutorialY = 0
        End If
        If TutorialStep = 38 Then
            Text = "Uma vez que ativado o extrator permanecerá removendo especiarias do solo de tempos em tempos, o extrator acumula especiaria enquanto estiver ativo mas você deve ir coletá-lo, basta utilizar CTRL no extrator para coletar especiarias."
            Image = 9
            TutorialY = 0
        End If
        If TutorialStep = 39 Then
            Text = "Matar habitantes também os fará deixar materiais, cada espécie deixa um material diferente. Estes materiais são muito bons para fazer extratores e combustíveis."
            Image = 10
            TutorialY = 0
        End If
        If TutorialStep = 40 Then
            Text = "Ainda continuando sobre os modos, o segundo e mais simples é o modo de Miniboss. Neste caso haverá apenas um inimigo muito forte no planeta, derrotá-lo irá completar a missão."
            Image = 2
            TutorialY = 0
        End If
        If TutorialStep = 41 Then
            Text = "O modo de COLETA você deverá destruir os recursos espalhados pelo planeta até completar a quantidade necessária para completar a missão."
            Image = 2
            TutorialY = 0
        End If
        If TutorialStep = 42 Then
            Text = "O modo de DESTRUIÇÃO pede que você destrua todas as construções espalhadas pelo planeta para completar a missão"
            Image = 2
            TutorialY = 0
        End If
        If TutorialStep = 43 Then
            Text = "No modo de SAQUE os habitantes deixam cair pequenos tesouros após serem destruídos, você deve capturar estes tesouros para completar a missão."
            Image = 2
            TutorialY = 0
        End If
        If TutorialStep = 44 Then
            Text = "No modo de PROTEÇÃO alguns piratas estão atacando um planeta aliado, você deve expulsá-los utilizando o máximo de violência que conseguir."
            Image = 2
            TutorialY = 0
        End If
        If TutorialStep = 45 Then
            Text = "Agora você sabe tudo que precisa para começar sua jornada, Sayajin!. Eu lhe desejo boa sorte, aqui está a sua nave como prometido, e lhe darei também um Scouter, é uma ferramenta muito importante para todos os Sayajins usada para medir o Poder de luta de seus adversários, boas aventuras!."
        End If
    End If
    
    Width = 512
    Height = 256
    RenderTexture Tex_GUI(33), TutorialX, TutorialY, 0, 0, Width, Height, Width, Height
    RenderText Font_Default, WordWrap(Text, 320), TutorialX + 190, TutorialY + 130, Yellow
    
    If Image > 0 Then
        RenderTexture Tex_Tutorial(Image), TutorialX + 60 + (Width / 2) - (Tex_Tutorial(Image).Width / 2), TutorialY + Height, 0, 0, Tex_Tutorial(Image).Width, Tex_Tutorial(Image).Height, Tex_Tutorial(Image).Width, Tex_Tutorial(Image).Height
    End If
    
    If Not HideContinue Then
        Dim color As Long
        color = Yellow
        If GlobalX >= TutorialX + 300 And GlobalX <= TutorialX + 300 + 100 Then
            If GlobalY >= TutorialY + 240 And GlobalY <= TutorialY + 240 + 16 Then
                color = BrightRed
            End If
        End If
        RenderText Font_Default, "< Continuar >", TutorialX + 300, TutorialY + 240, color
    End If
End Sub
Function QBToRGBA(ByVal ColorNum As Long) As Long
    Select Case ColorNum
        Case Black: QBToRGBA = D3DColorRGBA(0, 0, 0, 255)
        Case Blue: QBToRGBA = D3DColorRGBA(127, 0, 0, 255)
        Case Green: QBToRGBA = D3DColorRGBA(0, 127, 0, 255)
        Case Cyan: QBToRGBA = D3DColorRGBA(0, 127, 127, 255)
        Case Red: QBToRGBA = D3DColorRGBA(127, 0, 0, 255)
        Case Magenta: QBToRGBA = D3DColorRGBA(200, 200, 0, 255)
        Case Brown: QBToRGBA = D3DColorRGBA(127, 127, 0, 255)
        Case Grey: QBToRGBA = D3DColorRGBA(80, 80, 80, 255)
        Case DarkGrey: QBToRGBA = D3DColorRGBA(40, 40, 40, 255)
        Case BrightBlue: QBToRGBA = D3DColorRGBA(0, 0, 255, 255)
        Case BrightGreen: QBToRGBA = D3DColorRGBA(0, 255, 0, 255)
        Case BrightCyan: QBToRGBA = D3DColorRGBA(127, 127, 255, 255)
        Case BrightRed: QBToRGBA = D3DColorRGBA(255, 0, 0, 255)
        Case Pink: QBToRGBA = D3DColorRGBA(255, 127, 127, 255)
        Case Yellow: QBToRGBA = D3DColorRGBA(255, 255, 0, 255)
        Case White: QBToRGBA = D3DColorRGBA(255, 255, 255, 255)
    End Select
End Function

Sub DrawConquistas()
    RenderTexture Tex_GUI(41), 150, 150, 0, 0, 500, 300, 500, 300
    DrawConquistasPage
    
    Dim i As Long
    Dim X As Long, Y As Long
    Dim Width As Long, Height As Long
    ' draw buttons
    For i = 63 To 64
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
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(i).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        End If
    Next
    
    Dim sRECT As RECT
    ' draw bar background
    With sRECT
        .Top = 19 ' HP bar background
        .Left = 0
        .Right = .Left + 125
        .Bottom = .Top + 19
    End With
    
    RenderTexture Tex_Bars, 350, 420, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 200)
    
    Dim Completed As Byte
    For i = 1 To UBound(Conquistas)
        If Player(MyIndex).Conquistas(i) = 1 Then Completed = Completed + 1
    Next i
    
    Completed = (Completed / UBound(Conquistas)) * 100
    
    ' draw the bar proper
    With sRECT
        .Top = 10 ' HP bar
        .Left = 0
        .Right = .Left + Completed
        .Bottom = .Top + 9
    End With
    
    RenderTexture Tex_Bars, 363, 427, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(0, 255, 0, 200)
    
    RenderText Font_Default, "Progresso total", 357, 430, White
End Sub

Sub DrawPopConquista()
    If PopConquistaNum > 0 Then
        Dim Time As Long
        Time = 10000
        If PopConquistaTick + Time > GetTickCount Then
            Dim X As Long, Y As Long
            Dim Alpha As Long
            X = 400 - (Tex_GUI(42).Width / 2)
            Y = 400
            Alpha = 255 - (((GetTickCount - PopConquistaTick) / Time) * 255)
            RenderTexture Tex_GUI(42), X, Y, 0, 0, Tex_GUI(42).Width, Tex_GUI(42).Height, Tex_GUI(42).Width, Tex_GUI(42).Height, D3DColorRGBA(100, 255, 100, Alpha)
            RenderText Font_Default, "CONQUISTA COMPLETADA!", X + 8, Y + 8, BrightCyan, 255 - (Alpha)
            RenderText Font_Default, Trim$(Conquistas(PopConquistaNum).name), X + 8, Y + 20, Yellow, 255 - (Alpha)
            RenderText Font_Default, WordWrap(Trim$(Conquistas(PopConquistaNum).Desc), Tex_GUI(42).Width - 10), X + 8, Y + 32, White, 255 - (Alpha)
        End If
    End If
End Sub

Sub DrawConquistasPage()
    Dim ConquestStart As Long, ConquestEnd As Long
    Dim i As Long, n As Long
    ConquestStart = PageNum * 3 + 1
    ConquestEnd = ConquestStart + 2
    n = 0
    For i = ConquestStart To ConquestEnd
        If i > 0 And i <= UBound(Conquistas) Then
            If Player(MyIndex).Conquistas(i) = 0 Then
                RenderTexture Tex_GUI(42), 170, 190 + (n * 80), 0, 0, Tex_GUI(42).Width, Tex_GUI(42).Height, Tex_GUI(42).Width, Tex_GUI(42).Height, D3DColorRGBA(255, 255, 255, 100)
            Else
                RenderTexture Tex_GUI(42), 170, 190 + (n * 80), 0, 0, Tex_GUI(42).Width, Tex_GUI(42).Height, Tex_GUI(42).Width, Tex_GUI(42).Height, D3DColorRGBA(0, 255, 0, 100)
            End If
            Dim X As Long, Y As Long
            X = 180
            Y = 195 + (n * 80)
            RenderText Font_Default, Trim$(Conquistas(i).name), X, Y, Yellow
            Y = Y + 14
            RenderText Font_Default, WordWrap(Trim$(Conquistas(i).Desc), Tex_GUI(42).Width - 10), X, Y, White
            Y = Y + 28
            RenderText Font_Default, "RECOMPENSAS:", X, Y, BrightGreen
            
            Dim z As Long
            For z = 1 To 5
                If Conquistas(i).Reward(z).num > 0 Then
                    Dim ItemPic As Long
                    ItemPic = Item(Conquistas(i).Reward(z).num).Pic
                    RenderTexture Tex_Item(ItemPic), X + 50 + (z * 32), Y - 8, 0, 0, 32, 32, 32, 32
                    RenderText Font_Default, ConvertCurrency(Conquistas(i).Reward(z).value), X + 56 + (z * 32), Y + 8, White
                    
                    If GlobalX >= X + 50 + (z * 32) And GlobalX <= X + 82 + (z * 32) Then
                        If GlobalY >= Y - 8 And GlobalY <= Y + 24 Then
                            RenderTexture Tex_GUI(25), X + 4 + (z * 32), Y - 32, 0, 0, 128, 64, 128, 64, D3DColorRGBA(255, 255, 255, 255)
                            RenderText Font_Default, Trim$(Item(Conquistas(i).Reward(z).num).name), X + (z * 32) - (getWidth(Font_Default, Trim$(Item(Conquistas(i).Reward(z).num).name)) / 2) + 68, Y - 28, Grey, , True
                        End If
                    End If
                End If
            Next z
            n = n + 1
            
            X = 170 + Tex_GUI(42).Width - getWidth(Font_Default, Conquistas(i).EXP & "xp")
            RenderText Font_Default, Conquistas(i).EXP & "xp", X, Y, BrightGreen
            
            If Conquistas(i).Progress > 0 Then
                ' draw bar background
                Dim sRECT As RECT
                With sRECT
                    .Top = 0 ' HP bar background
                    .Left = 0
                    .Right = .Left + 48
                    .Bottom = .Top + 10
                End With
                
                X = X - 60
                Y = Y + 4
                
                RenderTexture Tex_Bars, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 200)
            
                If Player(MyIndex).ConquistaProgress(i) > Conquistas(i).Progress Then Player(MyIndex).ConquistaProgress(i) = Conquistas(i).Progress
            
                ' draw the bar proper
                With sRECT
                    .Top = 0 ' HP bar
                    .Left = 0
                    .Right = .Left + ((Player(MyIndex).ConquistaProgress(i) / Conquistas(i).Progress) * 50)
                    .Bottom = .Top + 10
                End With

                RenderTexture Tex_Bars, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(0, 255, 0, 200)
            End If
        End If
    Next i
End Sub
Sub DrawWindowService()
    If ServiceWindowTick + 10000 > GetTickCount Then
        RenderTexture Tex_GUI(24), 305, 282, 0, 0, 190, 36, 190, 36, D3DColorRGBA(255, 255, 255, 255)
        RenderText Font_Default, "SERVIÇO COMPLETADO!", 305, 270, BrightGreen
        RenderText Font_Default, "Recompensa: " & ServiceWindowGold & "z", 310, 285, Grey
        RenderText Font_Default, "Experiência: " & ServiceWindowExp & "exp", 310, 300, Yellow
    End If
End Sub
