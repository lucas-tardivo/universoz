VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
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
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame frmPlanetable 
      Caption         =   "Planetas próprios"
      Height          =   3855
      Left            =   3480
      TabIndex        =   128
      Top             =   3960
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtCentroLevel 
         Height          =   270
         Left            =   3720
         TabIndex        =   145
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtMinutes 
         Height          =   270
         Left            =   3480
         TabIndex        =   143
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Fechar"
         Height          =   255
         Left            =   3120
         TabIndex        =   141
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Frame Frame8 
         Caption         =   "Custo da evolução"
         Height          =   1695
         Left            =   120
         TabIndex        =   132
         Top             =   1200
         Width           =   4575
         Begin VB.TextBox txtYellow 
            Height          =   270
            Left            =   3360
            TabIndex        =   140
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtBlue 
            Height          =   270
            Left            =   3360
            TabIndex        =   138
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtRed 
            Height          =   270
            Left            =   3360
            TabIndex        =   136
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtMoedas 
            Height          =   270
            Left            =   3360
            TabIndex        =   134
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "Especiarias amarelas necessárias:"
            Height          =   255
            Left            =   120
            TabIndex        =   139
            Top             =   1320
            Width           =   4215
         End
         Begin VB.Label Label17 
            Caption         =   "Especiarias azuis necessárias:"
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   960
            Width           =   4215
         End
         Begin VB.Label Label14 
            Caption         =   "Especiarias vermelhas necessárias:"
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label lblGold 
            Caption         =   "Moedas necessárias:"
            Height          =   255
            Left            =   120
            TabIndex        =   133
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.HScrollBar scrlEvolution 
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   840
         Width           =   4575
      End
      Begin VB.CheckBox chkPlanetable 
         Caption         =   "É de planetas próprios"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label20 
         Caption         =   "Centro level:"
         Height          =   255
         Left            =   2640
         TabIndex        =   144
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Tempo em minutos para evoluir:"
         Height          =   255
         Left            =   120
         TabIndex        =   142
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label lblEvolution 
         Caption         =   "Evolui para:"
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   600
         Width           =   4575
      End
   End
   Begin VB.Frame fraIA 
      Caption         =   "Special powers"
      Height          =   3855
      Left            =   3480
      TabIndex        =   83
      Top             =   3960
      Visible         =   0   'False
      Width           =   4815
      Begin VB.ListBox lstIA 
         Height          =   1140
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   120
         List            =   "frmEditor_NPC.frx":3342
         TabIndex        =   93
         Top             =   240
         Width           =   4575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Close"
         Height          =   255
         Left            =   3000
         TabIndex        =   87
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Frame fraStun 
         Caption         =   "Stun"
         Height          =   1935
         Left            =   120
         TabIndex        =   107
         Top             =   1440
         Visible         =   0   'False
         Width           =   4575
         Begin VB.HScrollBar scrlAAnim 
            Height          =   255
            Left            =   2160
            TabIndex        =   116
            Top             =   1560
            Width           =   2295
         End
         Begin VB.HScrollBar scrlAImpact 
            Height          =   255
            Left            =   1320
            Max             =   5
            TabIndex        =   114
            Top             =   1200
            Width           =   3135
         End
         Begin VB.HScrollBar scrlADur 
            Height          =   255
            Left            =   1320
            Max             =   5
            Min             =   1
            TabIndex        =   112
            Top             =   840
            Value           =   1
            Width           =   3135
         End
         Begin VB.HScrollBar scrlAChance 
            Height          =   255
            Left            =   1320
            Max             =   100
            Min             =   1
            TabIndex        =   110
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.CheckBox chkStun 
            Caption         =   "Have skill"
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label lblAAnim 
            Caption         =   "Animação:"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label lblAImpact 
            Caption         =   "Impacto:"
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lblADur 
            Caption         =   "Duração:"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblAChance 
            Caption         =   "Chance:"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame fraSpawn 
         Caption         =   "Spawn Horde"
         Height          =   1935
         Left            =   120
         TabIndex        =   84
         Top             =   1440
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CheckBox chkSActive 
            Caption         =   "Have skill"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   4335
         End
         Begin VB.HScrollBar scrlSAnim 
            Height          =   255
            Left            =   2280
            Max             =   100
            TabIndex        =   91
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ComboBox cmbSEvent 
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   960
            Width           =   3615
         End
         Begin VB.HScrollBar scrlChance 
            Height          =   255
            Left            =   1200
            Max             =   100
            TabIndex        =   86
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label lblSAnim 
            Caption         =   "Animation:"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label lblSEvent 
            Caption         =   "Event:"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblSChance 
            Caption         =   "Chance:"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame fraStorm 
         Caption         =   "Storm"
         Height          =   1935
         Left            =   120
         TabIndex        =   119
         Top             =   1440
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CheckBox chkStorm 
            Caption         =   "Tem essa habilidade"
            Height          =   180
            Left            =   2280
            TabIndex        =   126
            Top             =   240
            Width           =   2175
         End
         Begin VB.HScrollBar scrlCooldown 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   125
            Top             =   1440
            Width           =   4335
         End
         Begin VB.HScrollBar scrlStormChance 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   123
            Top             =   960
            Width           =   4335
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   121
            Top             =   480
            Width           =   4335
         End
         Begin VB.Label lblCooldown 
            Caption         =   "Cooldown:"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   1200
            Width           =   4335
         End
         Begin VB.Label lblChance 
            Caption         =   "Chance:"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   720
            Width           =   4335
         End
         Begin VB.Label lblSpell 
            Caption         =   "Spell:"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame fraTeleport 
         Caption         =   "Teleport"
         Height          =   1935
         Left            =   120
         TabIndex        =   99
         Top             =   1440
         Visible         =   0   'False
         Width           =   4575
         Begin VB.HScrollBar sctlTAnim 
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   1440
            Width           =   4335
         End
         Begin VB.HScrollBar scrlTStun 
            Height          =   255
            Left            =   1440
            Max             =   5
            TabIndex        =   104
            Top             =   840
            Width           =   3015
         End
         Begin VB.HScrollBar scrlTChance 
            Height          =   255
            Left            =   1440
            Max             =   100
            TabIndex        =   101
            Top             =   480
            Width           =   3015
         End
         Begin VB.CheckBox chkTeleport 
            Caption         =   "Have skill"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lblTAnim 
            Caption         =   "Animação:"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lblTStun 
            Caption         =   "Paralisia:"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblTChance 
            Caption         =   "Chance:"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   480
            Width           =   1455
         End
      End
   End
   Begin VB.Frame fraGFX 
      Height          =   2295
      Left            =   3480
      TabIndex        =   95
      Top             =   5520
      Visible         =   0   'False
      Width           =   4815
      Begin VB.ComboBox cmbGFX 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":336A
         Left            =   1320
         List            =   "frmEditor_NPC.frx":3374
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Close"
         Height          =   255
         Left            =   3000
         TabIndex        =   96
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Special graphics are pre-programmed by GoPlay Games"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   117
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label11 
         Caption         =   "Graphic:"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ranged"
      Height          =   2295
      Left            =   3480
      TabIndex        =   73
      Top             =   5520
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlArrowAnim 
         Height          =   255
         Left            =   120
         Max             =   500
         TabIndex        =   80
         Top             =   1920
         Width           =   4575
      End
      Begin VB.CheckBox chkRanged 
         Caption         =   "Have skill"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   2415
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   135
         Left            =   120
         Max             =   1000
         TabIndex        =   77
         Top             =   1440
         Width           =   4575
      End
      Begin VB.HScrollBar scrlProjectile 
         Height          =   255
         Left            =   120
         Max             =   50
         TabIndex        =   75
         Top             =   840
         Width           =   4575
      End
      Begin VB.PictureBox picProjectile 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   74
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label10 
         Caption         =   "Animação:"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Dano base:"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label lblProjectile 
         Caption         =   "Image:"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   39
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   38
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   37
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data"
      Height          =   7815
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command12 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   146
         Top             =   240
         Width           =   255
      End
      Begin VB.Frame Frame4 
         Height          =   2055
         Left            =   2520
         TabIndex        =   64
         Top             =   2040
         Visible         =   0   'False
         Width           =   2295
         Begin VB.CommandButton Command10 
            Caption         =   "Planetas próprios"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Special graphic"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Special powers"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Ranged"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   480
            Width           =   2055
         End
         Begin VB.CheckBox chkShadow 
            Caption         =   "Shadow"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Close"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Flying"
            CausesValidation=   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Menu"
         Height          =   255
         Left            =   4320
         TabIndex        =   63
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         TabIndex        =   59
         Text            =   "0"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<"
         Height          =   255
         Left            =   2280
         TabIndex        =   56
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtLevelToPDL 
         Height          =   285
         Left            =   3360
         TabIndex        =   55
         Top             =   2760
         Width           =   855
      End
      Begin VB.HScrollBar scrlEffect 
         Height          =   255
         Left            =   2640
         Max             =   5
         TabIndex        =   52
         Top             =   3720
         Width           =   2175
      End
      Begin VB.HScrollBar scrlEvent 
         Height          =   255
         Left            =   2640
         TabIndex        =   50
         Top             =   4200
         Width           =   2175
      End
      Begin VB.HScrollBar scrlMoveSpeed 
         Height          =   255
         Left            =   2640
         Max             =   10
         TabIndex        =   49
         Top             =   3960
         Value           =   1
         Width           =   2175
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   600
         TabIndex        =   43
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   4080
         TabIndex        =   42
         Top             =   2400
         Width           =   855
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2640
         TabIndex        =   41
         Top             =   3480
         Width           =   2175
      End
      Begin VB.PictureBox picSprite 
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
         Height          =   480
         Left            =   4440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   29
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   28
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3388
         Left            =   1320
         List            =   "frmEditor_NPC.frx":33A1
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1680
         Width           =   3615
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   25
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   600
         Width           =   3975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stats"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   4440
         Width           =   4815
         Begin VB.HScrollBar scrlAttackSpeed 
            Height          =   135
            Left            =   3240
            Max             =   2000
            Min             =   100
            TabIndex        =   57
            Top             =   600
            Value           =   200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   135
            Index           =   1
            Left            =   120
            Max             =   32000
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   135
            Index           =   2
            Left            =   1680
            Max             =   32000
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   135
            Index           =   3
            Left            =   3240
            Max             =   32000
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   135
            Index           =   4
            Left            =   120
            Max             =   32000
            TabIndex        =   7
            Top             =   600
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   135
            Index           =   5
            Left            =   1680
            Max             =   32000
            TabIndex        =   6
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblAttackSpeed 
            AutoSize        =   -1  'True
            Caption         =   "Speed: 1000 mil"
            Height          =   180
            Left            =   3240
            TabIndex        =   58
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Str: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "End: 0"
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Int: 0"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   13
            Top             =   360
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Agi: 0"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   465
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Will: 0"
            Height          =   180
            Index           =   5
            Left            =   1680
            TabIndex        =   11
            Top             =   720
            Width           =   480
         End
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Caption         =   "Flying"
         Height          =   2295
         Left            =   120
         TabIndex        =   66
         Top             =   5400
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlFlyTick 
            Height          =   255
            Left            =   120
            Max             =   1000
            Min             =   100
            TabIndex        =   70
            Top             =   840
            Value           =   100
            Width           =   4575
         End
         Begin VB.CheckBox chkFly 
            Caption         =   "Flying NPC"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label8 
            Caption         =   "Velocidade que bate as asas:"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   600
            Width           =   4575
         End
      End
      Begin VB.Frame fraDrop 
         Caption         =   "Drop"
         Height          =   2295
         Left            =   120
         TabIndex        =   16
         Top             =   5400
         Width           =   4815
         Begin VB.TextBox txtItem 
            Height          =   270
            Left            =   3840
            TabIndex        =   118
            Top             =   1080
            Width           =   855
         End
         Begin VB.HScrollBar scrlItemDrop 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   62
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   19
            Top             =   1800
            Width           =   3495
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   18
            Top             =   1440
            Width           =   3495
         End
         Begin VB.TextBox txtChance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   17
            Text            =   "0"
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label7 
            Height          =   255
            Left            =   3360
            TabIndex        =   61
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   1800
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "Num: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   21
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Chance"
            Height          =   180
            Left            =   120
            TabIndex        =   20
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   570
         End
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn (in seconds)"
         Height          =   180
         Left            =   120
         TabIndex        =   60
         Top             =   3120
         UseMnemonic     =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   2880
         TabIndex        =   54
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label lblEffect 
         AutoSize        =   -1  'True
         Caption         =   "Effect: None"
         Height          =   180
         Left            =   120
         TabIndex        =   53
         Top             =   3720
         Width           =   930
      End
      Begin VB.Label lblEvent 
         Caption         =   "Event: None"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label lblMoveSpeed 
         Caption         =   "Movement Speed: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   3960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Som:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dmg:"
         Height          =   180
         Left            =   3360
         TabIndex        =   45
         Top             =   2400
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PDL:"
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Speak:"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   525
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   1680
         TabIndex        =   31
         Top             =   2400
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "HP:"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   285
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lista"
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7440
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFly_Click()
    Npc(EditorIndex).Fly = chkFly.Value
End Sub

Private Sub chkPlanetable_Click()
    Npc(EditorIndex).IsPlanetable = chkPlanetable.Value
End Sub

Private Sub chkRanged_Click()
    scrlProjectile.Enabled = chkRanged.Value
    scrlDamage.Enabled = chkRanged.Value
    Npc(EditorIndex).Ranged = chkRanged.Value
End Sub

Private Sub chkSActive_Click()
    Npc(EditorIndex).IA(NPCIA.Spawn).Data(1) = chkSActive.Value
End Sub

Private Sub chkShadow_Click()
    Npc(EditorIndex).Shadow = chkShadow.Value
End Sub

Private Sub chkStorm_Click()
    Npc(EditorIndex).IA(NPCIA.Storm).Data(1) = chkStorm.Value
End Sub

Private Sub chkStun_Click()
    Npc(EditorIndex).IA(NPCIA.Stun).Data(1) = chkStun.Value
End Sub

Private Sub chkTeleport_Click()
    Npc(EditorIndex).IA(NPCIA.Shunppo).Data(1) = chkTeleport.Value
End Sub

Private Sub cmbBehaviour_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Npc(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbGFX_Click()
    Npc(EditorIndex).GFXPack = cmbGFX.ListIndex
End Sub

Private Sub cmbSEvent_Click()
    Npc(EditorIndex).IA(NPCIA.Spawn).Data(3) = cmbSEvent.ListIndex + 1
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Command1_Click()
    txtLevel.Text = 10 + (Int(Val(GetVar(App.Path & "\data files\exp.ini", "EXPERIENCE", "Exp" & (Val(txtLevelToPDL.Text) - 2))) * 0.02))
End Sub

Private Sub Command10_Click()
    UpdateConstruction
End Sub

Private Sub UpdateConstruction()
    frmPlanetable.visible = True
    chkPlanetable.Value = Npc(EditorIndex).IsPlanetable
    scrlEvolution.max = MAX_NPCS
    scrlEvolution.Value = Npc(EditorIndex).Evolution
    txtMoedas.Text = Val(Npc(EditorIndex).ECostGold)
    txtRed.Text = Val(Npc(EditorIndex).ECostRed)
    txtBlue.Text = Val(Npc(EditorIndex).ECostBlue)
    txtYellow.Text = Val(Npc(EditorIndex).ECostYellow)
    txtMinutes.Text = Val(Npc(EditorIndex).TimeToEvolute)
    txtCentroLevel.Text = Val(Npc(EditorIndex).MinLevel)
End Sub

Private Sub Command11_Click()
    frmPlanetable.visible = False
End Sub

Private Sub Command12_Click()
    Dim n As Long
    n = Val(InputBox("Digite o numero do NPC:"))
    Npc(EditorIndex) = Npc(n)
    NpcEditorInit
    
End Sub

Private Sub Command2_Click()
    Frame4.visible = Not Frame4.visible
    chkShadow.Value = Npc(EditorIndex).Shadow
End Sub

Private Sub Command3_Click()
    Frame5.visible = Not Frame5.visible
    chkFly.Value = Npc(EditorIndex).Fly
    If Npc(EditorIndex).FlyTick < 100 Then Npc(EditorIndex).FlyTick = 100
    scrlFlyTick.Value = Npc(EditorIndex).FlyTick
End Sub

Private Sub Command4_Click()
Frame4.visible = False
End Sub

Private Sub Command5_Click()
    Frame6.visible = Not Frame6.visible
End Sub

Private Sub Command6_Click()
    fraIA.visible = True
End Sub

Private Sub Command7_Click()
    fraIA.visible = False
End Sub

Private Sub Command8_Click()
    fraGFX.visible = True
    cmbGFX.ListIndex = Npc(EditorIndex).GFXPack
End Sub

Private Sub Command9_Click()
    fraGFX.visible = False
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite.max = NumCharacters
    scrlAnimation.max = MAX_ANIMATIONS
    scrlEffect.max = MAX_EFFECTS
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub lstIA_Click()
    Dim i As Long
    fraSpawn.visible = False
    fraTeleport.visible = False
    fraStun.visible = False
    
    If lstIA.ListIndex + 1 = NPCIA.Spawn Then
        fraSpawn.visible = True
        
        cmbSEvent.Clear
        For i = 1 To MAX_EVENTS
            cmbSEvent.AddItem i & ": " & Trim$(Events(i).Name)
        Next i
        
        scrlSAnim.max = MAX_ANIMATIONS
        
        chkSActive.Value = Npc(EditorIndex).IA(NPCIA.Spawn).Data(1)
        scrlChance.Value = Npc(EditorIndex).IA(NPCIA.Spawn).Data(2)
        cmbSEvent.ListIndex = Npc(EditorIndex).IA(NPCIA.Spawn).Data(3) - 1
        scrlSAnim.Value = Npc(EditorIndex).IA(NPCIA.Spawn).Data(4)
    End If
    
    If lstIA.ListIndex + 1 = NPCIA.Shunppo Then
        fraTeleport.visible = True
        
        chkTeleport.Value = Npc(EditorIndex).IA(NPCIA.Shunppo).Data(1)
        scrlTChance.Value = Npc(EditorIndex).IA(NPCIA.Shunppo).Data(2)
        scrlTStun.Value = Npc(EditorIndex).IA(NPCIA.Shunppo).Data(3)
        sctlTAnim.Value = Npc(EditorIndex).IA(NPCIA.Shunppo).Data(4)
        sctlTAnim.max = MAX_ANIMATIONS
    End If
    
    If lstIA.ListIndex + 1 = NPCIA.Stun Then
        fraStun.visible = True
        
        chkStun.Value = Npc(EditorIndex).IA(NPCIA.Stun).Data(1)
        If Npc(EditorIndex).IA(NPCIA.Stun).Data(2) = 0 Then Npc(EditorIndex).IA(NPCIA.Stun).Data(2) = 1
        scrlAChance.Value = Npc(EditorIndex).IA(NPCIA.Stun).Data(2)
        If Npc(EditorIndex).IA(NPCIA.Stun).Data(3) = 0 Then Npc(EditorIndex).IA(NPCIA.Stun).Data(3) = 1
        scrlADur.Value = Npc(EditorIndex).IA(NPCIA.Stun).Data(3)
        scrlAImpact.Value = Npc(EditorIndex).IA(NPCIA.Stun).Data(4)
        scrlAAnim.Value = Npc(EditorIndex).IA(NPCIA.Stun).Data(5)
        scrlAAnim.max = MAX_ANIMATIONS
    End If
    
    If lstIA.ListIndex + 1 = NPCIA.Storm Then
        fraStorm.visible = True
        
        chkStorm.Value = Npc(EditorIndex).IA(NPCIA.Storm).Data(1)
        scrlSpell.max = MAX_SPELLS
        scrlSpell.Value = Npc(EditorIndex).IA(NPCIA.Storm).Data(3)
        scrlStormChance.Value = Npc(EditorIndex).IA(NPCIA.Storm).Data(2)
        scrlCooldown.Value = Npc(EditorIndex).IA(NPCIA.Storm).Data(4)
    End If
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NpcEditorInit
    If Npc(EditorIndex).IsPlanetable = 1 Then UpdateConstruction
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAAnim_Change()
    Npc(EditorIndex).IA(NPCIA.Stun).Data(5) = scrlAAnim.Value
    If scrlAAnim.Value > 0 Then
        lblAAnim.Caption = "Animation: " & Trim$(Animation(scrlAAnim.Value).Name)
    Else
        lblAAnim.Caption = "Animation: None"
    End If
End Sub

Private Sub scrlAChance_Change()
    Npc(EditorIndex).IA(NPCIA.Stun).Data(2) = scrlAChance.Value
    lblAChance.Caption = "Chance: " & scrlAChance.Value & "%"
End Sub

Private Sub scrlADur_Change()
    Npc(EditorIndex).IA(NPCIA.Stun).Data(3) = scrlADur.Value
    lblADur.Caption = "Duration: " & scrlADur.Value & "s"
End Sub

Private Sub scrlAImpact_Change()
    Npc(EditorIndex).IA(NPCIA.Stun).Data(4) = scrlAImpact.Value
    lblAImpact.Caption = "Impact: " & scrlAImpact.Value
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnimation.Caption = "Anim: " & sString
    Npc(EditorIndex).Animation = scrlAnimation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlArrowAnim_Change()
    If scrlArrowAnim.Value <> 0 Then
        Label10.Caption = "Animation: " & Trim$(Animation(scrlArrowAnim.Value).Name)
    Else
        Label10.Caption = "Animation: None"
    End If
    Npc(EditorIndex).ArrowAnimation = scrlArrowAnim.Value
End Sub

Private Sub scrlAttackSpeed_Change()
    Npc(EditorIndex).AttackSpeed = scrlAttackSpeed.Value
    lblAttackSpeed.Caption = "Veloc: " & scrlAttackSpeed.Value & " mil"
End Sub

Private Sub scrlChance_Change()
    Npc(EditorIndex).IA(NPCIA.Spawn).Data(2) = scrlChance.Value
    
    lblSChance.Caption = "Chance: " & scrlChance.Value & "%"
End Sub

Private Sub scrlCooldown_Change()
    lblCooldown.Caption = "Cooldown: " & scrlCooldown.Value & "s"
    Npc(EditorIndex).IA(NPCIA.Storm).Data(4) = scrlCooldown.Value
End Sub

Private Sub scrlDamage_Change()
    Label9.Caption = "Dmg Base: " & scrlDamage.Value & "%"
    Npc(EditorIndex).ArrowDamage = scrlDamage.Value
End Sub

Private Sub scrlEffect_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlEffect.Value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Effect(scrlEffect.Value).Name)
    End If
    lblEffect.Caption = "Effect: " & sString
    Npc(EditorIndex).Effect = scrlEffect.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEffect_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlEvent_Change()
    If scrlEvent.Value > 0 Then
        lblEvent.Caption = "Event: " & Trim$(Events(scrlEvent.Value).Name)
    Else
        lblEvent.Caption = "Event: None"
    End If
    Npc(EditorIndex).Event = scrlEvent.Value
End Sub

Private Sub scrlEvolution_Change()
    If scrlEvolution.Value > 0 Then
        lblEvolution.Caption = "Evolui para: " & scrlEvolution.Value & " " & Trim$(Npc(scrlEvolution.Value).Name)
    Else
        lblEvolution.Caption = "Evolui para: <Nenhum>"
    End If
    Npc(EditorIndex).Evolution = scrlEvolution.Value
End Sub

Private Sub scrlFlyTick_Change()
Label8.Caption = "Wing speed: " & scrlFlyTick.Value & "ms"
Npc(EditorIndex).FlyTick = scrlFlyTick.Value
End Sub

Private Sub scrlItemDrop_Change()
    txtChance.Text = Npc(EditorIndex).Drop(scrlItemDrop.Value).Chance
    scrlNum.Value = Npc(EditorIndex).Drop(scrlItemDrop.Value).Num
    scrlValue.Value = Npc(EditorIndex).Drop(scrlItemDrop.Value).Value
End Sub

Private Sub scrlProjectile_Change()
lblProjectile.Caption = "Image: " & scrlProjectile.Value
Npc(EditorIndex).ArrowAnim = scrlProjectile.Value
End Sub

Private Sub scrlSAnim_Change()
    Npc(EditorIndex).IA(NPCIA.Spawn).Data(4) = scrlSAnim.Value
    lblSAnim.Caption = "Animation: " & Trim$(Animation(scrlSAnim.Value).Name)
End Sub

Private Sub scrlSpell_Change()
    Npc(EditorIndex).IA(NPCIA.Storm).Data(3) = scrlSpell.Value
    If scrlSpell.Value > 0 Then
        lblSpell.Caption = "Spell: " & Trim$(Spell(scrlSpell.Value).Name)
    Else
        lblSpell.Caption = "Spell: <Nenhuma>"
    End If
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Npc(EditorIndex).Sprite = scrlSprite.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRange.Caption = "Range: " & scrlRange.Value
    Npc(EditorIndex).Range = scrlRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNum.Caption = "Num: " & scrlNum.Value
    txtItem.Text = scrlNum.Value

    If scrlNum.Value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.Value).Name)
    End If
    
    Npc(EditorIndex).Drop(scrlItemDrop.Value).Num = scrlNum.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStat_Change(Index As Integer)
Dim prefix As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            prefix = "For: "
        Case 2
            prefix = "Con: "
        Case 3
            prefix = "Ki: "
        Case 4
            prefix = "Acerto: "
        Case 5
            prefix = "Esquiva: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).Value
    Npc(EditorIndex).Stat(Index) = scrlStat(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStormChance_Change()
    lblChance.Caption = "Chance: " & scrlStormChance.Value & "%"
    Npc(EditorIndex).IA(NPCIA.Storm).Data(2) = scrlStormChance.Value
End Sub

Private Sub scrlTChance_Change()
    Npc(EditorIndex).IA(NPCIA.Shunppo).Data(2) = scrlTChance.Value
    lblTChance.Caption = "Chance: " & scrlTChance.Value & "%"
End Sub

Private Sub scrlTStun_Change()
    Npc(EditorIndex).IA(NPCIA.Shunppo).Data(3) = scrlTStun.Value
    lblTStun.Caption = "Stun: " & scrlTStun.Value
End Sub

Private Sub scrlValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblValue.Caption = "Value: " & scrlValue.Value
    Npc(EditorIndex).Drop(scrlItemDrop.Value).Value = scrlValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub sctlTAnim_Change()
    Npc(EditorIndex).IA(NPCIA.Shunppo).Data(4) = sctlTAnim.Value
    If sctlTAnim.Value > 0 Then
        lblTAnim.Caption = "Animation: " & Trim$(Animation(sctlTAnim.Value).Name)
    Else
        lblTAnim.Caption = "Animation: None"
    End If
End Sub

Private Sub txtAttackSay_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Npc(EditorIndex).AttackSay = txtAttackSay.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtBlue_Change()
    If IsNumeric(txtBlue) Then
        Npc(EditorIndex).ECostBlue = Val(txtBlue)
    End If
End Sub

Private Sub txtCentroLevel_Change()
    Npc(EditorIndex).MinLevel = Val(txtCentroLevel)
End Sub

Private Sub txtChance_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtChance.Text) > 0 Then Exit Sub
    If IsNumeric(txtChance.Text) Then Npc(EditorIndex).Drop(scrlItemDrop.Value).Chance = Val(txtChance.Text)
    
    On Error Resume Next
    Label7.Caption = ""
    Label7.Caption = (1 / Val(txtChance.Text)) * 100 & "%"
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtChance_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtDamage.Text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.Text) Then Npc(EditorIndex).Damage = Val(txtDamage.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtEXP.Text) > 0 Then Exit Sub
    If IsNumeric(txtEXP.Text) Then Npc(EditorIndex).EXP = Val(txtEXP.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtHP.Text) > 0 Then Exit Sub
    If IsNumeric(txtHP.Text) Then Npc(EditorIndex).HP = Val(txtHP.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtItem_Change()
    If IsNumeric(txtItem) Then
        If Val(txtItem) > 0 And Val(txtItem) <= scrlNum.max Then
            scrlNum.Value = Val(txtItem)
        End If
    End If
End Sub

Private Sub txtLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtLevel.Text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.Text) Then Npc(EditorIndex).Level = Val(txtLevel.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtlevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevelToPDL_Change()
    If IsNumeric(txtLevelToPDL) Then
        Npc(EditorIndex).ND = Val(txtLevelToPDL)
    End If
End Sub

Private Sub txtMinutes_Change()
    Npc(EditorIndex).TimeToEvolute = Val(txtMinutes)
End Sub

Private Sub txtMoedas_Change()
    If IsNumeric(txtMoedas) Then
        Npc(EditorIndex).ECostGold = Val(txtMoedas)
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Npc(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtRed_Change()
    If IsNumeric(txtRed) Then
        Npc(EditorIndex).ECostRed = Val(txtRed)
    End If
End Sub

Private Sub txtSpawnSecs_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtSpawnSecs.Text) > 0 Then Exit Sub
    Npc(EditorIndex).SpawnSecs = Val(txtSpawnSecs.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Npc(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Npc(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMoveSpeed_Change()
    lblMoveSpeed.Caption = "Movement: " & scrlMoveSpeed.Value
    Npc(EditorIndex).speed = scrlMoveSpeed.Value
End Sub

Private Sub txtYellow_Change()
    If IsNumeric(txtYellow) Then
        Npc(EditorIndex).ECostYellow = Val(txtYellow)
    End If
End Sub
