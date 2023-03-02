VERSION 5.00
Begin VB.Form frmEditor_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17070
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Map.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   627
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1138
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picAttributes 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   9840
      ScaleHeight     =   7095
      ScaleWidth      =   7095
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Frame fraMapWarp 
         Caption         =   "Map Warp"
         Height          =   2775
         Left            =   1800
         TabIndex        =   35
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapWarp 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   42
            Top             =   2160
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapWarpY 
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1680
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarpX 
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarp 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   37
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label lblMapWarpY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblMapWarpX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblMapWarp 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraShop 
         Caption         =   "Shop"
         Height          =   1335
         Left            =   1920
         TabIndex        =   43
         Top             =   2640
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdShop 
            Caption         =   "Accept"
            Height          =   375
            Left            =   960
            TabIndex        =   45
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbShop 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraMapItem 
         Caption         =   "Map Item"
         Height          =   1815
         Left            =   1800
         TabIndex        =   29
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapItem 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1200
            TabIndex        =   34
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapItemValue 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   33
            Top             =   840
            Value           =   1
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMapItem 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   32
            Top             =   480
            Value           =   1
            Width           =   2535
         End
         Begin VB.PictureBox picMapItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   31
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lblMapItem 
            Caption         =   "Item: None x0"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraArena 
         Caption         =   "Arena"
         Height          =   2775
         Left            =   1800
         TabIndex        =   92
         Top             =   1920
         Width           =   3375
         Begin VB.CommandButton cmdArena 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   99
            Top             =   2160
            Width           =   1455
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   98
            Top             =   1680
            Width           =   2895
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   96
            Top             =   1080
            Width           =   2895
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   240
            Max             =   1000
            TabIndex        =   94
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label lblY 
            Caption         =   "Y:"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label lblX 
            Caption         =   "X:"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblMap 
            Caption         =   "Mapa:"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraEvent 
         Caption         =   "Event"
         Height          =   1455
         Left            =   1800
         TabIndex        =   66
         Top             =   2280
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlEvent 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   68
            Top             =   240
            Value           =   1
            Width           =   3135
         End
         Begin VB.CommandButton cmdEvent 
            Caption         =   "Okay"
            Height          =   375
            Left            =   1080
            TabIndex        =   67
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblEvent 
            Caption         =   "Event: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   600
            Width           =   3135
         End
      End
      Begin VB.Frame fraSoundEffect 
         Caption         =   "Sound Effect"
         Height          =   1455
         Left            =   1800
         TabIndex        =   101
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbSoundEffect 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3332
            Left            =   240
            List            =   "frmEditor_Map.frx":3342
            Style           =   2  'Dropdown List
            TabIndex        =   103
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdSoundEffectOk 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   102
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraSlide 
         Caption         =   "Slide"
         Height          =   1455
         Left            =   1800
         TabIndex        =   61
         Top             =   2640
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbSlide 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":335D
            Left            =   240
            List            =   "frmEditor_Map.frx":336D
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdSlide 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   62
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraTrap 
         Caption         =   "Trap"
         Height          =   1575
         Left            =   1800
         TabIndex        =   57
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlTrap 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   59
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton cmdTrap 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   58
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblTrap 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraHeal 
         Caption         =   "Heal"
         Height          =   1815
         Left            =   1800
         TabIndex        =   52
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbHeal 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3388
            Left            =   240
            List            =   "frmEditor_Map.frx":3392
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdHeal 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   54
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlHeal 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   53
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblHeal 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame fraNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   2655
         Left            =   1800
         TabIndex        =   24
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstNpc 
            Height          =   780
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlNpcDir 
            Height          =   255
            Left            =   240
            Max             =   3
            TabIndex        =   26
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton cmdNpcSpawn 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   25
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblNpcDir 
            Caption         =   "Direction: Up"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame fraResource 
         Caption         =   "Object"
         Height          =   1695
         Left            =   1800
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdResourceOk 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   22
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   21
            Top             =   600
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblResource 
            Caption         =   "Object:"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   2535
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   8160
      TabIndex        =   14
      Top             =   5520
      Width           =   1455
      Begin VB.CheckBox chkResources 
         Caption         =   "Resources"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton optEyeDropper 
         Caption         =   "Dropper"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkGrid 
         Caption         =   "32x32 Grid"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optBlock 
         Caption         =   "Block"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton optAttribs 
         Caption         =   "Atts"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optLayers 
         Caption         =   "Layers"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.HScrollBar scrlPictureX 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7320
      Width           =   7695
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7200
      Left            =   120
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   12
      Top             =   120
      Width           =   7680
   End
   Begin VB.VScrollBar scrlPictureY 
      Height          =   7215
      Left            =   7800
      Max             =   255
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame fraTileSet 
      Caption         =   "Tileset: 0"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   7680
      Width           =   7935
      Begin VB.HScrollBar scrlTileSet 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   1
         Top             =   360
         Value           =   1
         Width           =   7695
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   8640
      Width           =   1695
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      Height          =   5415
      Left            =   8160
      TabIndex        =   72
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask2Anim"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   100
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "FringeAnim"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   88
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "MaskAnim"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   87
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   86
         Top             =   4200
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   85
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   84
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Ground"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   82
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   81
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   390
         Left            =   120
         TabIndex        =   80
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Random tile"
         Height          =   1215
         Left            =   120
         TabIndex        =   76
         Top             =   3000
         Width           =   1215
         Begin VB.HScrollBar scrlFrequency 
            Height          =   255
            Left            =   120
            Max             =   100
            Min             =   1
            TabIndex        =   78
            Top             =   480
            Value           =   75
            Width           =   975
         End
         Begin VB.CommandButton cmdRandomTile 
            Caption         =   "Place"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblFrequency 
            Caption         =   "Freq.: 75"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Render type"
         Height          =   855
         Left            =   120
         TabIndex        =   73
         Top             =   2160
         Width           =   1215
         Begin VB.HScrollBar scrlAutotile 
            Height          =   255
            Left            =   120
            Max             =   10
            TabIndex        =   74
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblAutotile 
            Alignment       =   2  'Center
            Caption         =   "Normal"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Atributos"
      Height          =   5415
      Left            =   8160
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optArena 
         Caption         =   "Arena"
         Height          =   270
         Left            =   120
         TabIndex        =   91
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdPreencher 
         Caption         =   "Preencher"
         Height          =   375
         Left            =   120
         TabIndex        =   90
         Top             =   4440
         Width           =   1215
      End
      Begin VB.OptionButton optEvent 
         Caption         =   "Evento"
         Height          =   270
         Left            =   120
         TabIndex        =   65
         Top             =   3240
         Width           =   1215
      End
      Begin VB.OptionButton optSound 
         Caption         =   "Som"
         Height          =   270
         Left            =   120
         TabIndex        =   64
         Top             =   3000
         Width           =   1215
      End
      Begin VB.OptionButton optSlide 
         Caption         =   "Slide"
         Height          =   270
         Left            =   120
         TabIndex        =   51
         Top             =   2760
         Width           =   1215
      End
      Begin VB.OptionButton optTrap 
         Caption         =   "Armadilha"
         Height          =   270
         Left            =   120
         TabIndex        =   50
         Top             =   2520
         Width           =   1215
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Curar"
         Height          =   270
         Left            =   120
         TabIndex        =   49
         Top             =   2280
         Width           =   1215
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Banco"
         Height          =   270
         Left            =   120
         TabIndex        =   48
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton optShop 
         Caption         =   "Shop"
         Height          =   270
         Left            =   120
         TabIndex        =   46
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton optNpcSpawn 
         Caption         =   "Nascer NPC"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optResource 
         Caption         =   "Recurso"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Bloquear"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Teleportar"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Limpar"
         Height          =   390
         Left            =   120
         TabIndex        =   5
         Top             =   4920
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Evitar NPC"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Label lblTilePosition 
      Height          =   255
      Left            =   1200
      TabIndex        =   104
      Top             =   9120
      Width           =   5535
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArena_Click()
    picAttributes.visible = False
    fraArena.visible = False
End Sub

Private Sub cmdEvent_Click()
    MapEditorEventIndex = scrlEvent.Value
    picAttributes.visible = False
    fraEvent.visible = False
End Sub

Private Sub cmdHeal_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MapEditorHealType = cmbHeal.ListIndex + 1
    MapEditorHealAmount = scrlHeal.Value
    picAttributes.visible = False
    fraHeal.visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdMapItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorNum = scrlMapItem.Value
    ItemEditorValue = scrlMapItemValue.Value
    picAttributes.visible = False
    fraMapItem.visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdMapItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdMapWarp_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EditorWarpMap = scrlMapWarp.Value
    EditorWarpX = scrlMapWarpX.Value
    EditorWarpY = scrlMapWarpY.Value
    picAttributes.visible = False
    fraMapWarp.visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdMapWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdNpcSpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SpawnNpcNum = lstNpc.ListIndex + 1
    SpawnNpcDir = scrlNpcDir.Value
    picAttributes.visible = False
    fraNpcSpawn.visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdNpcSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdPreencher_Click()
Dim X, Y As Long, tileType As Byte

If optBlocked.Value = True Then tileType = TILE_TYPE_BLOCKED
If optNpcAvoid.Value = True Then tileType = TILE_TYPE_NPCAVOID

For X = 0 To Map.MaxX
    For Y = 0 To Map.MaxY
        Map.Tile(X, Y).Type = tileType
    Next Y
Next X
End Sub

Private Sub cmdRandomTile_Click()
Dim X As Long
Dim Y As Long
Dim Chance As Long
Dim rate As Long

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Chance = Rand(1, scrlFrequency.Value)
            rate = Rand(1, 100)
            
            If Chance >= rate Then
                Call MapEditorPlaceRandomTile(X, Y)
            End If
            
            DoEvents
        Next
    Next
End Sub

Private Sub cmdResourceOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ResourceEditorNum = scrlResource.Value
    picAttributes.visible = False
    fraResource.visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdResourceOk_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EditorShop = cmbShop.ListIndex
    picAttributes.visible = False
    fraShop.visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdShop_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSlide_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MapEditorSlideDir = cmbSlide.ListIndex
    picAttributes.visible = False
    fraSlide.visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSoundEffectOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MapEditorSound = soundCache(cmbSoundEffect.ListIndex + 1)
    picAttributes.visible = False
    fraSoundEffect.visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSoundEffectOk_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdTrap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MapEditorHealAmount = scrlTrap.Value
    picAttributes.visible = False
    fraTrap.visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Dim i As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' move the entire attributes box on screen
    picAttributes.Left = 8
    picAttributes.Top = 8
    
    PopulateLists
    
    cmbSoundEffect.Clear
    For i = 1 To UBound(soundCache)
        cmbSoundEffect.AddItem (soundCache(i))
    Next
    cmbSoundEffect.ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optDoor_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraMapWarp.visible = True
    
    scrlMapWarp.max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.max = MAX_BYTE
    scrlMapWarpY.max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optDoor_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub fraAttribs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If optBlocked.Value = True Or optNpcAvoid.Value = True Then
        cmdPreencher.Enabled = True
    Else
        cmdPreencher.Enabled = False
    End If
End Sub

Private Sub optArena_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraArena.visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optEvent_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optEvent_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraEvent.visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optEvent_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optHeal_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraHeal.visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optLayers_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If optLayers.Value Then
        fraLayers.visible = True
        fraAttribs.visible = False
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optLayers_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optAttribs_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If optAttribs.Value Then
        fraLayers.visible = False
        fraAttribs.visible = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optAttribs_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optNpcSpawn_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lstNpc.Clear
    
    For n = 1 To MAX_MAP_NPCS
        If Map.Npc(n) > 0 Then
            lstNpc.AddItem n & ": " & Npc(Map.Npc(n)).Name
        Else
            lstNpc.AddItem n & ": No Npc"
        End If
    Next n
    
    scrlNpcDir.Value = 0
    lstNpc.ListIndex = 0
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraNpcSpawn.visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optNpcSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraResource.visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optResource_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraShop.visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optShop_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSlide_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraSlide.visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraSoundEffect.visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSound_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optTrap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraTrap.visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSend_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorSend
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSend_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    Call MapEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdProperties_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Load frmEditor_MapProperties
    MapEditorProperties
    frmEditor_MapProperties.Show vbModal
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdProperties_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optWarp_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraMapWarp.visible = True
    
    scrlMapWarp.max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.max = MAX_BYTE
    scrlMapWarpY.max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraMapItem.visible = True

    scrlMapItem.max = MAX_ITEMS
    scrlMapItem.Value = 1
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdFill_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MapEditorFillLayer
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdFill_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdClear_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorClearLayer
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdClear_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdClear2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorClearAttribs
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdClear2_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    X = X + (frmEditor_Map.scrlPictureX.Value * 32)
    Y = Y + (frmEditor_Map.scrlPictureY.Value * 32)
    Call MapEditorChooseTile(Button, X, Y)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBack_MouseDown", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    X = X + (frmEditor_Map.scrlPictureX.Value * 32)
    Y = Y + (frmEditor_Map.scrlPictureY.Value * 32)
    Call MapEditorDrag(Button, X, Y)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBack_MouseMove", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAutotile_Change()
    Select Case scrlAutotile.Value
        Case 0 ' normal
            lblAutotile.Caption = "Normal"
        Case 1 ' autotile
            lblAutotile.Caption = "Autotile (VX)"
        Case 2 ' fake autotile
            lblAutotile.Caption = "Fake (VX)"
        Case 3 ' animated
            lblAutotile.Caption = "Animated (VX)"
        Case 4 ' cliff
            lblAutotile.Caption = "Cliff (VX)"
        Case 5 ' waterfall
            lblAutotile.Caption = "Waterfall (VX)"
        Case 6 ' autotile
            lblAutotile.Caption = "Autotile (XP)"
        Case 7 ' fake autotile
            lblAutotile.Caption = "Fake (XP)"
        Case 8 ' animated
            lblAutotile.Caption = "Animated (XP)"
        Case 9 ' cliff
            lblAutotile.Caption = "Cliff (XP)"
        Case 10 ' waterfall
            lblAutotile.Caption = "Waterfall (XP)"
    End Select
End Sub

Private Sub scrlEvent_Change()
    If Trim$(Events(scrlEvent.Value).Name) = vbNullString Then
        lblEvent.Caption = "Event: " & scrlEvent.Value
    Else
        lblEvent.Caption = "Event: " & scrlEvent.Value & " - " & Trim$(Events(scrlEvent.Value).Name)
    End If
End Sub

Private Sub scrlFrequency_Change()
    lblFrequency.Caption = "Freq.: " & scrlFrequency.Value
End Sub

Private Sub scrlHeal_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblHeal.Caption = "Amount: " & scrlHeal.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlHeal_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMap_Change()
    MapEditorArenaMap = scrlMap.Value
    lblMap.Caption = "Map: " & scrlMap.Value
End Sub

Private Sub scrlTrap_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblTrap.Caption = "Amount: " & scrlTrap.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTrap_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    If Item(scrlMapItem.Value).Type = ITEM_TYPE_CURRENCY Or Item(scrlMapItem.Value).Stackable > 0 Then
        scrlMapItemValue.Enabled = True
    Else
        scrlMapItemValue.Value = 1
        scrlMapItemValue.Enabled = False
    End If

    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapItem_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItem_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlMapItem_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapItem_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItemValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapItemValue_change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItemValue_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlMapItemValue_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapItemValue_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblMapWarp.Caption = "Map: " & scrlMapWarp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarp_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarp_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlMapWarp_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarp_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblMapWarpX.Caption = "X: " & scrlMapWarpX.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpX_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlMapWarpX_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblMapWarpY.Caption = "Y: " & scrlMapWarpY.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpY_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlMapWarpY_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNpcDir_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case scrlNpcDir.Value
        Case DIR_DOWN
            lblNpcDir = "Direction: Down"
        Case DIR_UP
            lblNpcDir = "Direction: Up"
        Case DIR_LEFT
            lblNpcDir = "Direction: Left"
        Case DIR_RIGHT
            lblNpcDir = "Direction: Right"
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNpcDir_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNpcDir_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlNpcDir_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNpcDir_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlResource_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblResource.Caption = "Resource: " & Resource(scrlResource.Value).Name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlResource_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlResource_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlResource_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlResource_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorTileScroll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPictureX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorTileScroll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPictureY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureX_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPictureY_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPictureX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureY_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPictureY_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPictureY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTileSet_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    fraTileSet.Caption = "Tileset: " & scrlTileSet.Value
    
    frmEditor_Map.scrlPictureY.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    
    MapEditorTileScroll
    
    EditorTileX = 0
    EditorTileY = 0
    EditorTileWidth = 1
    EditorTileHeight = 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTileSet_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTileSet_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlTileSet_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTileSet_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlX_Change()
    MapEditorArenaX = scrlX.Value
    lblX.Caption = "X: " & scrlX.Value
End Sub

Private Sub scrlY_Change()
    MapEditorArenaY = scrlY.Value
    lblY.Caption = "Y: " & scrlY.Value
End Sub
