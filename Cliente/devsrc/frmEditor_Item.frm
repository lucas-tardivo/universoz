VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10170
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   678
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame frmPlanetChange 
      Caption         =   "Terraplanagem"
      Height          =   2655
      Left            =   3360
      TabIndex        =   137
      Top             =   3480
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame frmSaibaman 
         Caption         =   "Saibaman"
         Height          =   1935
         Left            =   120
         TabIndex        =   160
         Top             =   600
         Width           =   6375
         Begin VB.TextBox txtIndex 
            Height          =   270
            Left            =   1320
            TabIndex        =   162
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Saibaman n:"
            Height          =   255
            Left            =   240
            TabIndex        =   161
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame frmClima 
         Caption         =   "Clima"
         Height          =   1935
         Left            =   120
         TabIndex        =   147
         Top             =   600
         Visible         =   0   'False
         Width           =   6375
         Begin VB.ComboBox cmbClima 
            Height          =   300
            ItemData        =   "frmEditor_Item.frx":3332
            Left            =   720
            List            =   "frmEditor_Item.frx":334B
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label Label17 
            Caption         =   "Clima:"
            Height          =   255
            Left            =   120
            TabIndex        =   148
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame frmAmbiente 
         Caption         =   "Ambiente"
         Height          =   1935
         Left            =   120
         TabIndex        =   144
         Top             =   600
         Visible         =   0   'False
         Width           =   6375
         Begin VB.ComboBox cmbAmbiente 
            Height          =   300
            ItemData        =   "frmEditor_Item.frx":3395
            Left            =   1080
            List            =   "frmEditor_Item.frx":33A5
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label16 
            Caption         =   "Ambiente:"
            Height          =   255
            Left            =   120
            TabIndex        =   145
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame frmColor 
         Caption         =   "Colorir planeta"
         Height          =   1935
         Left            =   120
         TabIndex        =   139
         Top             =   600
         Visible         =   0   'False
         Width           =   6375
         Begin VB.HScrollBar scrlTom 
            Height          =   255
            Left            =   1080
            Max             =   255
            Min             =   -255
            TabIndex        =   143
            Top             =   600
            Width           =   5175
         End
         Begin VB.ComboBox cmbCanal 
            Height          =   300
            ItemData        =   "frmEditor_Item.frx":33CD
            Left            =   720
            List            =   "frmEditor_Item.frx":33DD
            Style           =   2  'Dropdown List
            TabIndex        =   141
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label lblTom 
            Caption         =   "Tom:"
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Canal:"
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.ComboBox cmbPlanetChange 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":33FF
         Left            =   120
         List            =   "frmEditor_Item.frx":3424
         Style           =   2  'Dropdown List
         TabIndex        =   138
         Top             =   240
         Width           =   6375
      End
      Begin VB.Frame frmResource 
         Caption         =   "Construção"
         Height          =   1935
         Left            =   120
         TabIndex        =   155
         Top             =   600
         Visible         =   0   'False
         Width           =   6375
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   120
            TabIndex        =   157
            Top             =   480
            Width           =   6135
         End
         Begin VB.TextBox txtResourceLimit 
            Height          =   270
            Left            =   1680
            TabIndex        =   156
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblResource 
            Caption         =   "Recurso:"
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   240
            Width           =   6015
         End
         Begin VB.Label Label19 
            Caption         =   "Limite: 1 para cada                      níveis do núcleo "
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   840
            Width           =   4335
         End
      End
      Begin VB.Frame frmNPC 
         Caption         =   "NPC"
         Height          =   1935
         Left            =   120
         TabIndex        =   150
         Top             =   600
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox txtLimit 
            Height          =   270
            Left            =   1680
            TabIndex        =   154
            Top             =   840
            Width           =   855
         End
         Begin VB.HScrollBar scrlNPC 
            Height          =   255
            Left            =   120
            TabIndex        =   152
            Top             =   480
            Width           =   6135
         End
         Begin VB.Label Label18 
            Caption         =   "Limite: 1 para cada                      níveis do núcleo "
            Height          =   255
            Left            =   120
            TabIndex        =   153
            Top             =   840
            Width           =   4335
         End
         Begin VB.Label lblNPC 
            Caption         =   "NPC:"
            Height          =   255
            Left            =   120
            TabIndex        =   151
            Top             =   240
            Width           =   6015
         End
      End
   End
   Begin VB.Frame frmVIP 
      Caption         =   "VIP"
      Height          =   1815
      Left            =   3360
      TabIndex        =   133
      Top             =   3480
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtDiasVIP 
         Height          =   270
         Left            =   600
         TabIndex        =   135
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Dias:"
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frmBau 
      Caption         =   "Bau"
      Height          =   4215
      Left            =   3360
      TabIndex        =   122
      Top             =   3480
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkPacote 
         Caption         =   "Pacote"
         Height          =   255
         Left            =   5160
         TabIndex        =   136
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtItemNum 
         Height          =   270
         Left            =   3960
         TabIndex        =   132
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox lstItens 
         Height          =   3120
         Left            =   120
         TabIndex        =   130
         Top             =   960
         Width           =   6375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salvar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5280
         TabIndex        =   129
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtChance 
         Height          =   270
         Left            =   2880
         TabIndex        =   128
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtQuant 
         Height          =   270
         Left            =   1080
         TabIndex        =   126
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmbItem 
         Height          =   300
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblChanceTotal 
         Caption         =   "Total:"
         Height          =   255
         Left            =   4080
         TabIndex        =   131
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Chance:"
         Height          =   255
         Left            =   2160
         TabIndex        =   127
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Quantidade:"
         Height          =   375
         Left            =   120
         TabIndex        =   125
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Item:"
         Height          =   375
         Left            =   120
         TabIndex        =   123
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   3375
      Left            =   3360
      TabIndex        =   16
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox chkDrop 
         Caption         =   "Undroppable"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtPrice 
         Height          =   270
         Left            =   4440
         TabIndex        =   92
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlEffect 
         Height          =   255
         Left            =   5400
         Max             =   5
         TabIndex        =   66
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkStackable 
         Caption         =   "Stack"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2640
         Width           =   1335
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4440
         Max             =   99
         TabIndex        =   63
         Top             =   3000
         Width           =   2055
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4440
         Max             =   5
         TabIndex        =   61
         Top             =   2640
         Width           =   2055
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   2280
         Width           =   2055
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtDesc 
         Height          =   1095
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Top             =   1440
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4440
         Max             =   5
         TabIndex        =   23
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":34BF
         Left            =   4440
         List            =   "frmEditor_Item.frx":34CF
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   600
         Width           =   2055
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5400
         Max             =   5
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3507
         Left            =   120
         List            =   "frmEditor_Item.frx":3547
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
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
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   17
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblEffect 
         AutoSize        =   -1  'True
         Caption         =   "Effect: None"
         Height          =   180
         Left            =   3000
         TabIndex        =   67
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   64
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   62
         Top             =   2640
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Race:"
         Height          =   180
         Left            =   3000
         TabIndex        =   60
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3000
         TabIndex        =   57
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Raridade: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   29
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type (?):"
         Height          =   180
         Left            =   3000
         TabIndex        =   28
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   3000
         TabIndex        =   26
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "List"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Magia"
      Height          =   3135
      Left            =   3360
      TabIndex        =   44
      Top             =   4680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlSpellQuant 
         Height          =   255
         Left            =   240
         Max             =   100
         TabIndex        =   121
         Top             =   1440
         Width           =   5775
      End
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   45
         Top             =   840
         Value           =   1
         Width           =   5775
      End
      Begin VB.Label lblQuant 
         Caption         =   "Quantidade:"
         Height          =   255
         Left            =   240
         TabIndex        =   120
         Top             =   1200
         Width           =   5775
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   46
         Top             =   600
         Width           =   675
      End
   End
   Begin VB.Frame frmExtrator 
      Caption         =   "Coletor"
      Height          =   975
      Left            =   3360
      TabIndex        =   109
      Top             =   3600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.HScrollBar scrlExtratorNum 
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   480
         Width           =   6375
      End
      Begin VB.Label lblExtratorNum 
         Caption         =   "Resource:"
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame frmTitulo 
      Caption         =   "Title"
      Height          =   975
      Left            =   3360
      TabIndex        =   104
      Top             =   3600
      Width           =   6615
      Begin VB.CheckBox chkIcon 
         Caption         =   "Icon"
         Height          =   255
         Left            =   1800
         TabIndex        =   108
         Top             =   600
         Width           =   1695
      End
      Begin VB.PictureBox picCor 
         Height          =   255
         Left            =   600
         ScaleHeight     =   195
         ScaleWidth      =   1035
         TabIndex        =   107
         Top             =   240
         Width           =   1095
      End
      Begin VB.HScrollBar scrlTituloCor 
         Height          =   255
         Left            =   120
         Max             =   15
         TabIndex        =   106
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frmDragonball 
      Caption         =   "Dragonball"
      Height          =   975
      Left            =   3360
      TabIndex        =   101
      Top             =   3600
      Width           =   2655
      Begin VB.HScrollBar scrlDragonball 
         Height          =   255
         Left            =   120
         Max             =   7
         Min             =   1
         TabIndex        =   103
         Top             =   480
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblNumber 
         Caption         =   "Numero:"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame frmEsoterica 
      Caption         =   "Esoterica"
      Height          =   975
      Left            =   3360
      TabIndex        =   93
      Top             =   3600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   255
         Left            =   5400
         TabIndex        =   99
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtBonus 
         Height          =   270
         Left            =   1200
         TabIndex        =   98
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtSecs 
         Height          =   270
         Left            =   960
         TabIndex        =   96
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "0h"
         Height          =   255
         Left            =   2160
         TabIndex        =   97
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "Bonus %:"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Minutes"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   975
      Left            =   3360
      TabIndex        =   5
      Top             =   3600
      Width           =   6615
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   720
         Max             =   32000
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2880
         Max             =   32000
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   32000
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   720
         Max             =   32000
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2880
         Max             =   32000
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
   Begin VB.Frame FrmCombustivel 
      Caption         =   "Combustivel"
      Height          =   975
      Left            =   3360
      TabIndex        =   117
      Top             =   3600
      Width           =   6615
      Begin VB.HScrollBar scrlBonus 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   119
         Top             =   480
         Width           =   6375
      End
      Begin VB.Label lblBonus 
         Caption         =   "Bonus:"
         Height          =   375
         Left            =   120
         TabIndex        =   118
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frmNave 
      Caption         =   "Nave"
      Height          =   975
      Left            =   3360
      TabIndex        =   112
      Top             =   3600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.HScrollBar scrlNaveSpeed 
         Height          =   255
         Left            =   2640
         Max             =   10
         TabIndex        =   116
         Top             =   480
         Width           =   3855
      End
      Begin VB.HScrollBar scrlNaveSprite 
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblNaveSpeed 
         Caption         =   "Velocidade de movimento:"
         Height          =   255
         Left            =   2640
         TabIndex        =   115
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblNaveSprite 
         Caption         =   "Sprite:"
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equip"
      Height          =   3135
      Left            =   3360
      TabIndex        =   30
      Top             =   4680
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H008080FF&
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
         Height          =   840
         Left            =   1200
         ScaleHeight     =   56
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   87
         Top             =   240
         Width           =   480
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame fraWeapon 
         Caption         =   "Weapon"
         Height          =   2055
         Left            =   120
         TabIndex        =   68
         Top             =   960
         Visible         =   0   'False
         Width           =   6375
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
            Left            =   2880
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   85
            Top             =   1320
            Width           =   480
         End
         Begin VB.Frame fraProjectile 
            Caption         =   "Projectile"
            Height          =   1695
            Left            =   3480
            TabIndex        =   76
            Top             =   240
            Width           =   2775
            Begin VB.HScrollBar scrlProjectileNum 
               Height          =   255
               Left            =   1320
               TabIndex        =   80
               Top             =   240
               Width           =   1335
            End
            Begin VB.HScrollBar scrlProjectileRange 
               Height          =   255
               Left            =   1320
               Max             =   255
               TabIndex        =   79
               Top             =   600
               Width           =   1335
            End
            Begin VB.HScrollBar scrlProjectileRotation 
               Height          =   255
               Left            =   1320
               Max             =   100
               TabIndex        =   78
               Top             =   960
               Width           =   1335
            End
            Begin VB.HScrollBar scrlProjectileAmmo 
               Height          =   255
               Left            =   1320
               TabIndex        =   77
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label lblProjectileNum 
               Caption         =   "Num: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblProjectileRange 
               Caption         =   "Range: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lblProjectileRotation 
               Caption         =   "Rotate: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lblProjectileAmmo 
               Caption         =   "Ammo: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   1320
               Width           =   1215
            End
         End
         Begin VB.CheckBox chk2Handed 
            Caption         =   "2 mãos"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   1320
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cmbTool 
            Height          =   300
            ItemData        =   "frmEditor_Item.frx":361B
            Left            =   1320
            List            =   "frmEditor_Item.frx":362B
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   600
            Width           =   2055
         End
         Begin VB.HScrollBar scrlSpeed 
            Height          =   255
            LargeChange     =   100
            Left            =   1320
            Max             =   3000
            Min             =   100
            SmallChange     =   100
            TabIndex        =   70
            Top             =   960
            Value           =   100
            Width           =   2055
         End
         Begin VB.TextBox txtDamage 
            Height          =   270
            Left            =   1320
            TabIndex        =   69
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tool:"
            Height          =   180
            Left            =   120
            TabIndex        =   74
            Top             =   600
            Width           =   390
         End
         Begin VB.Label lblDamage 
            AutoSize        =   -1  'True
            Caption         =   "Damage:"
            Height          =   180
            Left            =   120
            TabIndex        =   73
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   675
         End
         Begin VB.Label lblSpeed 
            AutoSize        =   -1  'True
            Caption         =   "Speed: 0.1"
            Height          =   180
            Left            =   120
            TabIndex        =   72
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   825
         End
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   4680
         Max             =   32000
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   2760
         Max             =   32000
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5640
         Max             =   32000
         TabIndex        =   33
         Top             =   480
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3720
         Max             =   32000
         TabIndex        =   32
         Top             =   480
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   1800
         Max             =   32000
         TabIndex        =   31
         Top             =   480
         Width           =   855
      End
      Begin VB.Frame fraArmor 
         Caption         =   "Armor"
         Height          =   2055
         Left            =   120
         TabIndex        =   89
         Top             =   960
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox txtDefence 
            Height          =   270
            Left            =   1320
            TabIndex        =   90
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Defence:"
            Height          =   180
            Left            =   120
            TabIndex        =   91
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   675
         End
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   88
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   4680
         TabIndex        =   40
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   2760
         TabIndex        =   39
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   5640
         TabIndex        =   38
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   3720
         TabIndex        =   37
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   1800
         TabIndex        =   36
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Consumo"
      Height          =   3135
      Left            =   3360
      TabIndex        =   41
      Top             =   4680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkInstant 
         Caption         =   "Instantaneo?"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   54
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   52
         Top             =   1800
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   50
         Top             =   1200
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   42
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         Caption         =   "Soltar magia: None"
         Height          =   180
         Left            =   120
         TabIndex        =   55
         Top             =   2160
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   53
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub chkDrop_Click()
    Item(EditorIndex).CantDrop = chkDrop.Value
End Sub

Private Sub chkIcon_Click()
    Item(EditorIndex).Data2 = chkIcon.Value
End Sub

Private Sub chkPacote_Click()
    Item(EditorIndex).Data1 = chkPacote.Value
End Sub

Private Sub cmbAmbiente_Click()
    Item(EditorIndex).Data2 = cmbAmbiente.ListIndex
End Sub

Private Sub cmbBind_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbCanal_Click()
    Item(EditorIndex).Data2 = cmbCanal.ListIndex
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClima_Click()
    Item(EditorIndex).Data2 = cmbClima.ListIndex
End Sub

Private Sub cmbItem_Change()
    txtItemNum.Text = cmbItem.ListIndex
End Sub

Private Sub cmbPlanetChange_Change()
    cmbPlanetChange_Click
End Sub

Private Sub cmbPlanetChange_Click()
    Item(EditorIndex).Data1 = cmbPlanetChange.ListIndex
    
    LoadItemConfigs
End Sub



Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Command1_Click()
frmEsoterica.visible = False
End Sub

Private Sub Command2_Click()
    Dim i As Long
    i = lstItens.ListIndex + 1
    If i > 0 Then
        Item(EditorIndex).LuckySlot(i).itemNum = cmbItem.ListIndex
        Item(EditorIndex).LuckySlot(i).Quant = Val(txtQuant.Text)
        Item(EditorIndex).LuckySlot(i).Chance = Val(txtChance.Text)
        PopulateList
    End If
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPic.max = numitems
    scrlAnim.max = MAX_ANIMATIONS
    scrlPaperdoll.max = NumPaperdolls
    scrlEffect.max = MAX_EFFECTS
    scrlProjectileNum.max = NumProjectiles
    scrlProjectileAmmo.max = MAX_ITEMS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub PopulateList()
    Dim i As Long
    cmbItem.Clear
    cmbItem.AddItem "<Nenhum>"
    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) <> vbNullString Then
            cmbItem.AddItem i & ": " & Trim$(Item(i).Name)
        Else
            cmbItem.AddItem i & ": <Sem nome>"
        End If
    Next i
    lstItens.Clear
    Dim total As Long
    For i = 1 To 40
        If Item(EditorIndex).LuckySlot(i).itemNum > 0 And Item(EditorIndex).LuckySlot(i).itemNum < MAX_ITEMS Then
            lstItens.AddItem i & ": " & Item(EditorIndex).LuckySlot(i).Quant & " (" & Item(EditorIndex).LuckySlot(i).itemNum & ")" & Trim$(Item(Item(EditorIndex).LuckySlot(i).itemNum).Name) & " Chance: " & Item(EditorIndex).LuckySlot(i).Chance & "%"
            total = total + Item(EditorIndex).LuckySlot(i).Chance
        Else
            lstItens.AddItem i & ": <Nenhum>"
        End If
    Next i
    lblChanceTotal.Caption = "Total: " & total & "%"
    If total <> 100 Then
        lblChanceTotal.ForeColor = QBColor(Red)
    Else
        lblChanceTotal.ForeColor = QBColor(Green)
    End If
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Or (cmbType.ListIndex = ITEM_TYPE_SCOUTER) Then
        fraEquipment.visible = True
        chkStackable.visible = False
        Item(EditorIndex).Stackable = 0
        If (cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
            fraWeapon.visible = True
            fraArmor.visible = False
        Else
            fraWeapon.visible = False
            fraArmor.visible = True
        End If
        'scrlDamage_Change
    Else
        fraEquipment.visible = False
        chkStackable.visible = True
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_EXTRATOR Then
        frmExtrator.visible = True
        'scrlVitalMod_Change
    Else
        frmExtrator.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_NAVE Then
        frmNave.visible = True
        frmNave.ZOrder (0)
        'scrlVitalMod_Change
    Else
        frmNave.visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.visible = True
    Else
        fraSpell.visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_ESOTERICA) Then
        frmEsoterica.visible = True
    Else
        frmEsoterica.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_DRAGONBALL Then
        frmDragonball.visible = True
    Else
        frmDragonball.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_TITULO Then
        frmTitulo.visible = True
    Else
        frmTitulo.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_COMBUSTIVEL Then
        FrmCombustivel.visible = True
    Else
        FrmCombustivel.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_BAU Then
        frmBau.visible = True
        If Item(EditorIndex).Data1 > 1 Then Item(EditorIndex).Data1 = 0
        chkPacote.Value = Item(EditorIndex).Data1
        PopulateList
    Else
        frmBau.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_VIP Then
        frmVIP.visible = True
    Else
        frmVIP.visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_PLANETCHANGE Then
        frmPlanetChange.visible = True
    Else
        frmPlanetChange.visible = False
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chk2Handed_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Handed = chk2Handed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chk2Handed_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkStackable_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Stackable = chkStackable.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkStackable_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstItens_Click()
    Command2.Enabled = True
    Dim i As Long
    i = lstItens.ListIndex + 1
    If Item(EditorIndex).LuckySlot(i).itemNum > 0 Then
        cmbItem.ListIndex = Item(EditorIndex).LuckySlot(i).itemNum
        txtQuant.Text = Item(EditorIndex).LuckySlot(i).Quant
        txtChance.Text = Item(EditorIndex).LuckySlot(i).Chance
    Else
        cmbItem.ListIndex = 0
        txtQuant.Text = vbNullString
        txtChance.Text = vbNullString
    End If
End Sub

Private Sub scrlAccessReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "Access Req: " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHp.Value
    Item(EditorIndex).AddHP = scrlAddHp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddEXP = scrlAddExp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.Value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.Value).Name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBonus_Change()
    lblBonus.Caption = "Bonus: " & scrlBonus.Value & "%"
    Item(EditorIndex).Data1 = scrlBonus.Value
End Sub

Private Sub scrlDragonball_Change()
    lblNumber.Caption = "Number: " & scrlDragonball.Value
    Item(EditorIndex).Dragonball = scrlDragonball.Value
End Sub

Private Sub scrlExtratorNum_Change()
    If scrlExtratorNum.Value > 0 Then
        lblExtratorNum.Caption = "Resource: " & scrlExtratorNum.Value & " " & Trim$(Resource(scrlExtratorNum).Name)
    Else
        lblExtratorNum.Caption = "Resource: Nenhum"
    End If
    Item(EditorIndex).Data2 = scrlExtratorNum.Value
End Sub

Private Sub scrlNaveSpeed_Change()
    lblNaveSpeed.Caption = "Velocidade de movimento: " & scrlNaveSpeed.Value
    Item(EditorIndex).Data2 = scrlNaveSpeed.Value
End Sub

Private Sub scrlNaveSprite_Change()
    lblNaveSprite.Caption = "Sprite: " & scrlNaveSprite.Value
    Item(EditorIndex).Data1 = scrlNaveSprite.Value
End Sub

Private Sub scrlNPC_Change()
    If scrlNPC.Value > 0 Then
        lblNPC.Caption = "NPC: " & scrlNPC.Value & " " & Trim$(Npc(scrlNPC.Value).Name)
    Else
        lblNPC.Caption = "NPC: <Nenhum>"
    End If
    Item(EditorIndex).Data2 = scrlNPC.Value
End Sub

Private Sub scrlProjectileAmmo_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileAmmo.Caption = "Munição: " & scrlProjectileAmmo.Value
    Item(EditorIndex).Ammo = scrlProjectileAmmo.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileAmmo_Change", "frmEditor_Item", Err.Ammober, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectileNum_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileNum.Caption = "Num: " & scrlProjectileNum.Value
    Item(EditorIndex).Projectile = scrlProjectileNum.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileNum_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectileRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRange.Caption = "Alcance: " & scrlProjectileRange.Value
    Item(EditorIndex).Range = scrlProjectileRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_Item", Err.Rangeber, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectileRotation_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRotation.Caption = "Rotação: " & scrlProjectileRotation.Value * 0.5
    Item(EditorIndex).Rotation = scrlProjectileRotation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRotation_Change", "frmEditor_Item", Err.Rotationber, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlResource_Change()
    If scrlResource.Value > 0 Then
        lblResource.Caption = "Resource: " & scrlResource.Value & " " & Trim$(Resource(scrlResource.Value).Name)
    Else
        lblResource.Caption = "Resource: <Nenhum>"
    End If
    Item(EditorIndex).Data2 = scrlResource.Value
End Sub

Private Sub scrlSpellQuant_Change()
    lblQuant.Caption = "Quantidade: " & scrlSpellQuant.Value
    Item(EditorIndex).Data2 = scrlSpellQuant.Value
End Sub

Private Sub scrlTituloCor_Change()
    picCor.BackColor = QBColor(scrlTituloCor.Value)
    Item(EditorIndex).Data1 = scrlTituloCor.Value
End Sub

Private Sub scrlTom_Change()
    Item(EditorIndex).Data3 = scrlTom.Value
    lblTom.Caption = "Tom: " & scrlTom.Value
End Sub

Private Sub txtBonus_Change()
    Item(EditorIndex).EsotericaBonus = Val(txtBonus)
End Sub

Private Sub txtDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data2 = Val(txtDamage.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
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
    Item(EditorIndex).Effect = scrlEffect.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Img: " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDiasVIP_Change()
    Item(EditorIndex).Data1 = Val(txtDiasVIP)
End Sub

Private Sub txtIndex_Change()
    Item(EditorIndex).Data2 = Val(txtIndex)
End Sub

Private Sub txtItemNum_Change()
    If IsNumeric(txtItemNum) Then
        On Error Resume Next
        cmbItem.ListIndex = Val(txtItemNum)
    End If
End Sub

Private Sub txtLimit_Change()
    Item(EditorIndex).Data3 = Val(txtLimit)
End Sub

Private Sub txtPrice_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Price = Val(txtPrice.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value / 1000
    Item(EditorIndex).speed = scrlSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
Dim Text As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            Text = "+ For: "
        Case 2
            Text = "+ Con: "
        Case 3
            Text = "+ KI: "
        Case 4
            Text = "+ Des: "
        Case 5
            Text = "+ Tec: "
    End Select
            
    lblStatBonus(Index).Caption = Text & scrlStatBonus(Index).Value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim Text As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            Text = "For: "
        Case 2
            Text = "Con: "
        Case 3
            Text = "KI: "
        Case 4
            Text = "Des: "
        Case 5
            Text = "Tec: "
    End Select
    
    lblStatReq(Index).Caption = Text & scrlStatReq(Index).Value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If Len(Trim$(Spell(scrlSpell.Value).Name)) > 0 Then
        lblSpellName.Caption = "Nome: " & Trim$(Spell(scrlSpell.Value).Name)
    Else
        lblSpellName.Caption = "Nome: Nenhum"
    End If
    
    lblSpell.Caption = "Magia: " & scrlSpell.Value
    
    Item(EditorIndex).Data1 = scrlSpell.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDefence_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data2 = Val(txtDefence.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDefence_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtResourceLimit_Change()
    Item(EditorIndex).Data3 = Val(txtResourceLimit)
End Sub

Private Sub txtSecs_Change()
    If IsNumeric(txtSecs) Then
        Label9.Caption = Int(Val(txtSecs) / 60) & "h"
    Else
        Label9.Caption = "Erro na conta"
    End If
    
    Item(EditorIndex).EsotericaTime = Val(txtSecs)
End Sub
