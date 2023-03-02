VERSION 5.00
Begin VB.Form frmEditor_MapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkFly 
      Caption         =   "Proibir voo (No fly)"
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox cmbSound 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0000
         Left            =   960
         List            =   "frmMapProperties.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   55
         Top             =   240
         Width           =   3135
      End
      Begin VB.HScrollBar scrlPanorama 
         Height          =   255
         Left            =   2400
         TabIndex        =   52
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0004
         Left            =   960
         List            =   "frmMapProperties.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "BGS:"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label lblPanorama 
         Caption         =   "Panorama: None"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moral:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Fog"
      Height          =   2535
      Left            =   4440
      TabIndex        =   44
      Top             =   3960
      Width           =   2055
      Begin VB.ComboBox cmbDir 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0045
         Left            =   120
         List            =   "frmMapProperties.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   2040
         Width           =   1815
      End
      Begin VB.HScrollBar scrlFogOpacity 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   49
         Top             =   1620
         Width           =   1815
      End
      Begin VB.HScrollBar ScrlFog 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   47
         Top             =   480
         Width           =   1815
      End
      Begin VB.HScrollBar ScrlFogSpeed 
         Height          =   255
         Left            =   120
         Max             =   2550
         TabIndex        =   45
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblFogOpacity 
         Caption         =   "Fog Opacity: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label lblFog 
         Caption         =   "Fog: None"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblFogSpeed 
         Caption         =   "Fog Speed: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   810
         Width           =   1815
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1455
      Left            =   4440
      TabIndex        =   37
      Top             =   1080
      Width           =   2055
      Begin VB.HScrollBar scrlWeatherIntensity 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   40
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox CmbWeather 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0065
         Left            =   120
         List            =   "frmMapProperties.frx":007E
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblWeatherIntensity 
         Caption         =   "Intensity: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Weather:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Color"
      Height          =   1455
      Left            =   4440
      TabIndex        =   30
      Top             =   2520
      Width           =   2055
      Begin VB.HScrollBar scrlA 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   42
         Top             =   1080
         Width           =   855
      End
      Begin VB.HScrollBar ScrlR 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   33
         Top             =   480
         Width           =   855
      End
      Begin VB.HScrollBar ScrlG 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   32
         Top             =   480
         Width           =   855
      End
      Begin VB.HScrollBar ScrlB 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   31
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblA 
         Caption         =   "A: 0"
         Height          =   255
         Left            =   1080
         TabIndex        =   43
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblR 
         Caption         =   "R: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblG 
         Caption         =   "G: 0"
         Height          =   255
         Left            =   1080
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblB 
         Caption         =   "B: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Music"
      Height          =   3255
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   2055
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ListBox lstMusic 
         Height          =   2205
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmMaxSizes 
      Height          =   975
      Left            =   4440
      TabIndex        =   20
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtMaxX 
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Max X:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Max Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   630
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      Height          =   1455
      Left            =   2280
      TabIndex        =   14
      Top             =   1800
      Width           =   2055
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   18
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   17
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   16
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Current map: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Boot Map:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Boot X:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Boot Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   630
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      Height          =   3735
      Left            =   2280
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
      Begin VB.CheckBox chkDoNotAutoSpawn 
         Caption         =   "Dont spawn"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3360
         Width           =   1815
      End
      Begin VB.ListBox lstNpcs 
         Height          =   2595
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3000
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   6600
      Width           =   975
   End
   Begin VB.Frame frmAmbiente 
      Height          =   1335
      Left            =   120
      TabIndex        =   59
      Top             =   2280
      Width           =   2055
      Begin VB.OptionButton optAmbiente 
         Caption         =   "Pássaros (Forest)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   63
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton optAmbiente 
         Caption         =   "Gaivotas (Birds)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   62
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optAmbiente 
         Caption         =   "Morcegos (Bats)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optAmbiente 
         Caption         =   "Nada (None)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmEditor_MapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HScroll1_Change()
    
End Sub

Private Sub cmbDir_Click()
    Map.FogDir = cmbDir.ListIndex
End Sub

Private Sub Form_Load()
    ScrlFog.max = NumFogs
    scrlPanorama.max = NumPanoramas
End Sub

Private Sub chkDoNotAutoSpawn_Click()
    If lstNpcs.ListIndex > -1 Then
        Map.NpcSpawnType(lstNpcs.ListIndex + 1) = chkDoNotAutoSpawn.Value
    End If
End Sub

Private Sub cmdPlay_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    StopMusic
    PlayMusic lstMusic.List(lstMusic.ListIndex)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdPlay_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdStop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    StopMusic
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdStop_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOk_Click()
    Dim I As Long
    Dim sTemp As Long
    Dim X As Long, X2 As Long
    Dim Y As Long, Y2 As Long
    Dim tempArr() As TileRec
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsNumeric(txtMaxX.Text) Then txtMaxX.Text = Map.MaxX
    If Val(txtMaxX.Text) < MAX_MAPX Then txtMaxX.Text = MAX_MAPX
    If Val(txtMaxX.Text) > MAX_BYTE Then txtMaxX.Text = MAX_BYTE
    If Not IsNumeric(txtMaxY.Text) Then txtMaxY.Text = Map.MaxY
    If Val(txtMaxY.Text) < MAX_MAPY Then txtMaxY.Text = MAX_MAPY
    If Val(txtMaxY.Text) > MAX_BYTE Then txtMaxY.Text = MAX_BYTE

    With Map
        .Name = Trim$(txtName.Text)
        If lstMusic.ListIndex >= 0 Then
            .Music = lstMusic.List(lstMusic.ListIndex)
        Else
            .Music = vbNullString
        End If
        If cmbSound.ListIndex >= 0 Then
            .BGS = cmbSound.List(cmbSound.ListIndex)
        Else
            .BGS = vbNullString
        End If
        .Up = Val(txtUp.Text)
        .Down = Val(txtDown.Text)
        .Left = Val(txtLeft.Text)
        .Right = Val(txtRight.Text)
        .Moral = cmbMoral.ListIndex
        .BootMap = Val(txtBootMap.Text)
        .BootX = Val(txtBootX.Text)
        .BootY = Val(txtBootY.Text)
        
        .Fly = chkFly.Value
        
        .Weather = CmbWeather.ListIndex
        .WeatherIntensity = scrlWeatherIntensity.Value
        
        .Fog = ScrlFog.Value
        .FogSpeed = ScrlFogSpeed.Value
        .FogOpacity = scrlFogOpacity.Value
        
        .Red = ScrlR.Value
        .Green = ScrlG.Value
        .Blue = ScrlB.Value
        .Alpha = scrlA.Value
        
        ' set the data before changing it
        tempArr = Map.Tile
        X2 = Map.MaxX
        Y2 = Map.MaxY
        ' change the data
        .MaxX = Val(txtMaxX.Text)
        .MaxY = Val(txtMaxY.Text)
        ReDim Map.Tile(0 To .MaxX, 0 To .MaxY)

        If X2 > .MaxX Then X2 = .MaxX
        If Y2 > .MaxY Then Y2 = .MaxY

        For X = 0 To X2
            For Y = 0 To Y2
                .Tile(X, Y) = tempArr(X, Y)
            Next
        Next
        .Panorama = scrlPanorama.Value
    End With
    
    initAutotiles
    Unload frmEditor_MapProperties
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdOk_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Unload frmEditor_MapProperties
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstNpcs_Click()
Dim tmpString() As String
Dim npcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    ' set the combo box properly
    tmpString = Split(lstNpcs.List(lstNpcs.ListIndex))
    npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
    cmbNpc.ListIndex = Map.Npc(npcNum)
    chkDoNotAutoSpawn.Value = Map.NpcSpawnType(npcNum)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstNpcs_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbNpc_Click()
Dim tmpString() As String
Dim npcNum As Long
Dim X As Long, tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    ' set the combo box properly
    tmpString = Split(cmbNpc.List(cmbNpc.ListIndex))
    ' make sure it's not a clear
    If Not cmbNpc.List(cmbNpc.ListIndex) = "No NPC" Then
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
    Else
        npcNum = 0
    End If

    For X = 1 To MAX_MAP_NPCS
        If lstNpcs.Selected(X - 1) Then
             Map.Npc(X) = npcNum
        End If
    Next
    
    ' re-load the list
    tmpIndex = lstNpcs.ListIndex
    lstNpcs.Clear
    For X = 1 To MAX_MAP_NPCS
        If Map.Npc(X) > 0 Then
        lstNpcs.AddItem X & ": " & Trim$(Npc(Map.Npc(X)).Name)
        Else
            lstNpcs.AddItem X & ": No NPC"
        End If
    Next
    lstNpcs.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbNpc_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optAmbiente_Click(Index As Integer)
    Map.Ambiente = Index
End Sub

Private Sub scrlA_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblA.Caption = "Alfa: " & scrlA.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlA_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlB_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblB.Caption = "Azul: " & ScrlB.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlB_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlFog_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ScrlFog.Value = 0 Then
        lblFog.Caption = "Fog: None."
    Else
        lblFog.Caption = "Fog: " & ScrlFog.Value
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlFog_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFogOpacity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFogOpacity.Caption = "Fog Opacity: " & scrlFogOpacity.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlFogOpacity_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlFogSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFogSpeed.Caption = "Fog Speed: " & ScrlFogSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlFogSpeed_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlG_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblG.Caption = "Verd: " & ScrlG.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlG_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPanorama_Change()
    If scrlPanorama.Value = 0 Then
        lblPanorama.Caption = "Panorama: None."
    Else
        lblPanorama.Caption = "Panorama: " & scrlPanorama.Value
    End If
End Sub

Private Sub ScrlR_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblR.Caption = "Verm: " & ScrlR.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlR_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlWeatherIntensity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblWeatherIntensity.Caption = "Intensidade: " & scrlWeatherIntensity.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlWeatherIntensity_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
