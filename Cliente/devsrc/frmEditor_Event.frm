VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditor_Events 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Events Editor"
   ClientHeight    =   9600
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   640
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEditCommand 
      Caption         =   "Editar comando"
      Height          =   4695
      Left            =   9480
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton cmdEditOk 
         Caption         =   "Close"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   4200
         Width           =   5535
      End
      Begin VB.Frame fraChangeLevel 
         Caption         =   "Level"
         Height          =   3735
         Left            =   120
         TabIndex        =   247
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlChangeLevel 
            Height          =   255
            Left            =   120
            TabIndex        =   251
            Top             =   600
            Width           =   3855
         End
         Begin VB.OptionButton optLevelAction 
            Caption         =   "Sub"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   250
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton optLevelAction 
            Caption         =   "Add"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   249
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton optLevelAction 
            Caption         =   "Set"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   248
            Top             =   960
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.Label lblChangeLevel 
            Caption         =   "Level: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   252
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame fraChangeVariable 
         Caption         =   "Var"
         Height          =   3735
         Left            =   120
         TabIndex        =   86
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   3960
            TabIndex        =   94
            Text            =   "0"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   93
            Text            =   "0"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtVariableData 
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   92
            Text            =   "0"
            Top             =   960
            Width           =   3495
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Random"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   91
            Top             =   1560
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Sub"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   90
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Add"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   89
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Set"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   88
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.ComboBox cmbVariable 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Max"
            Height          =   255
            Index           =   37
            Left            =   3480
            TabIndex        =   97
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Min"
            Height          =   255
            Index           =   13
            Left            =   1560
            TabIndex        =   96
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Var:"
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   95
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame fraChangeSex 
         Caption         =   "Sex"
         Height          =   3735
         Left            =   120
         TabIndex        =   186
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.OptionButton optChangeSex 
            Caption         =   "Male"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   188
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optChangeSex 
            Caption         =   "Female"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   187
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame fraChangeClass 
         Caption         =   "Class"
         Height          =   3735
         Left            =   120
         TabIndex        =   182
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbChangeClass 
            Height          =   315
            Left            =   1440
            TabIndex        =   183
            Text            =   "cmbChangeClass"
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label Label3 
            Caption         =   "Change:"
            Height          =   255
            Left            =   120
            TabIndex        =   184
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame fraMenu 
         Caption         =   "Choices"
         Height          =   3735
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlMenuOptDest 
            Height          =   255
            Left            =   240
            Max             =   10
            Min             =   1
            TabIndex        =   82
            Top             =   3360
            Value           =   1
            Width           =   5175
         End
         Begin VB.TextBox txtMenuOptText 
            Height          =   285
            Left            =   1440
            TabIndex        =   81
            Top             =   2760
            Width           =   3855
         End
         Begin VB.CommandButton cmdRemoveMenuOption 
            Caption         =   "Remove"
            Height          =   375
            Left            =   3960
            TabIndex        =   80
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton cmdModifyMenuOption 
            Caption         =   "Change"
            Height          =   375
            Left            =   2040
            TabIndex        =   79
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddMenuOption 
            Caption         =   "Add"
            Height          =   375
            Left            =   120
            TabIndex        =   78
            Top             =   2280
            Width           =   1335
         End
         Begin VB.ListBox lstMenuOptions 
            Height          =   1035
            Left            =   120
            TabIndex        =   77
            Top             =   1200
            Width           =   5295
         End
         Begin VB.TextBox txtMenuQuery 
            Height          =   645
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   76
            Top             =   480
            Width           =   5325
         End
         Begin VB.Label lblMenuOptDest 
            Caption         =   "Destination: 1"
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   3120
            Width           =   5175
         End
         Begin VB.Label Label5 
            Caption         =   "Text:"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Answer"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame fraGiveItem 
         Caption         =   "Items"
         Height          =   3735
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlGiveItemAmount 
            Height          =   255
            Left            =   120
            Max             =   32000
            Min             =   1
            TabIndex        =   72
            Top             =   1080
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlGiveItemID 
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   480
            Width           =   5295
         End
         Begin VB.OptionButton optItemOperation 
            Caption         =   "Take"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   70
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton optItemOperation 
            Caption         =   "Trade"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   69
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton optItemOperation 
            Caption         =   "Give"
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   68
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblGiveItemAmount 
            Caption         =   "Amount: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   840
            Width           =   5295
         End
         Begin VB.Label lblGiveItemID 
            Caption         =   "Item: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame frmQuest 
         Caption         =   "Quest"
         Height          =   3735
         Left            =   120
         TabIndex        =   253
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbState 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0000
            Left            =   840
            List            =   "frmEditor_Event.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   257
            Top             =   1080
            Width           =   4455
         End
         Begin VB.HScrollBar scrlQuest 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   255
            Top             =   600
            Width           =   5055
         End
         Begin VB.Label Label7 
            Caption         =   "State:"
            Height          =   255
            Left            =   240
            TabIndex        =   256
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblQuest 
            Caption         =   "Quest:"
            Height          =   255
            Left            =   240
            TabIndex        =   254
            Top             =   360
            Width           =   5055
         End
      End
      Begin VB.Frame fraOpenShop 
         Caption         =   "Shop"
         Height          =   3735
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlOpenShop 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   65
            Top             =   480
            Value           =   1
            Width           =   5295
         End
         Begin VB.Label lblOpenShop 
            Caption         =   "Abrir Shop: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame fraMapWarp 
         Caption         =   "Teleport"
         Height          =   3735
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlWarpY 
            Height          =   255
            Left            =   120
            Max             =   250
            TabIndex        =   60
            Top             =   1680
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlWarpX 
            Height          =   255
            Left            =   120
            Max             =   250
            TabIndex        =   59
            Top             =   1080
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlWarpMap 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   58
            Top             =   480
            Value           =   1
            Width           =   5295
         End
         Begin VB.Label lblWarpY 
            Caption         =   "Y: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1440
            Width           =   5295
         End
         Begin VB.Label lblWarpX 
            Caption         =   "X: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   840
            Width           =   5295
         End
         Begin VB.Label lblWarpMap 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame fraChatbubble 
         Caption         =   "Bubble"
         Height          =   3735
         Left            =   120
         TabIndex        =   131
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.OptionButton optChatBubbleTarget 
            Caption         =   "NPC"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   134
            Top             =   1440
            Width           =   735
         End
         Begin VB.ComboBox cmbChatBubbleTarget 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0037
            Left            =   120
            List            =   "frmEditor_Event.frx":0039
            Style           =   2  'Dropdown List
            TabIndex        =   133
            Top             =   1800
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.TextBox txtChatbubbleText 
            Height          =   1005
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   132
            Top             =   360
            Width           =   3735
         End
         Begin VB.OptionButton optChatBubbleTarget 
            Caption         =   "Player"
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   135
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Text:"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   137
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Target type:"
            Height          =   255
            Index           =   39
            Left            =   120
            TabIndex        =   136
            Top             =   1440
            Width           =   1335
         End
      End
      Begin VB.Frame fraPlayerText 
         Caption         =   "Message"
         Height          =   3735
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlMessageSprite 
            Height          =   255
            Left            =   1800
            TabIndex        =   130
            Top             =   3360
            Width           =   3615
         End
         Begin VB.TextBox txtPlayerText 
            Height          =   3015
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   56
            Top             =   240
            Width           =   5355
         End
         Begin VB.Label lblMessageSprite 
            Caption         =   "Sprite: Player"
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   3360
            Width           =   1335
         End
      End
      Begin VB.Frame fraAddText 
         Caption         =   "Add Text"
         Height          =   3735
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtAddText_Text 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   240
            Width           =   5295
         End
         Begin VB.HScrollBar scrlAddText_Colour 
            Height          =   255
            Left            =   120
            Max             =   18
            TabIndex        =   50
            Top             =   2400
            Width           =   5295
         End
         Begin VB.OptionButton optChannel 
            Caption         =   "Player"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   49
            Top             =   2760
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optChannel 
            Caption         =   "Map"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   48
            Top             =   2760
            Width           =   1095
         End
         Begin VB.OptionButton optChannel 
            Caption         =   "Global"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   47
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label lblAddText_Colour 
            Caption         =   "Colour: Black"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   2160
            Width           =   3255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Channel"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   52
            Top             =   2760
            Width           =   1575
         End
      End
      Begin VB.Frame fraGoTo 
         Caption         =   "Goto"
         Height          =   3735
         Left            =   120
         TabIndex        =   110
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlGOTO 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   111
            Top             =   480
            Value           =   1
            Width           =   5295
         End
         Begin VB.Label lblGOTO 
            Caption         =   "Goto: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame fraAnimation 
         Caption         =   "Animation"
         Height          =   3735
         Left            =   120
         TabIndex        =   103
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlPlayAnimationY 
            Height          =   255
            Left            =   120
            Max             =   250
            Min             =   -1
            TabIndex        =   106
            Top             =   1680
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlPlayAnimationX 
            Height          =   255
            Left            =   120
            Max             =   250
            Min             =   -1
            TabIndex        =   105
            Top             =   1080
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlPlayAnimationAnim 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   104
            Top             =   480
            Value           =   1
            Width           =   5295
         End
         Begin VB.Label lblPlayAnimationY 
            Caption         =   "Y: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   1440
            Width           =   5295
         End
         Begin VB.Label lblPlayAnimationX 
            Caption         =   "X: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   840
            Width           =   5295
         End
         Begin VB.Label lblPlayAnimationAnim 
            Caption         =   "Animação: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.Frame fraChangeSwitch 
         Caption         =   "Switch"
         Height          =   3735
         Left            =   120
         TabIndex        =   98
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbPlayerSwitchSet 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":003B
            Left            =   1200
            List            =   "frmEditor_Event.frx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   795
            Width           =   4215
         End
         Begin VB.ComboBox cmbSwitch 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Switch:"
            Height          =   255
            Index           =   23
            Left            =   360
            TabIndex        =   102
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Change for:"
            Height          =   255
            Index           =   22
            Left            =   240
            TabIndex        =   101
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame fraBranch 
         Caption         =   "Condição"
         Height          =   3735
         Left            =   120
         TabIndex        =   139
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtBranchItemAmount 
            Height          =   285
            Left            =   3480
            TabIndex        =   246
            Top             =   960
            Width           =   1815
         End
         Begin VB.HScrollBar scrlNegative 
            Height          =   255
            Left            =   120
            TabIndex        =   160
            Top             =   3360
            Value           =   1
            Width           =   5295
         End
         Begin VB.HScrollBar scrlPositive 
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   2760
            Value           =   1
            Width           =   5295
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Var"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   155
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Switch"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   154
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Item"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   153
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Race"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   152
            Top             =   1320
            Width           =   1215
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Skill"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   151
            Top             =   1680
            Width           =   1215
         End
         Begin VB.ComboBox cmbBranchVar 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0056
            Left            =   1560
            List            =   "frmEditor_Event.frx":0058
            TabIndex        =   150
            Text            =   "cmbBranchVar"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtBranchVarReq 
            Height          =   285
            Left            =   4320
            TabIndex        =   149
            Text            =   "0"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cmbVarReqOperator 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":005A
            Left            =   3480
            List            =   "frmEditor_Event.frx":0070
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Level"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   147
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox cmbLevelReqOperator 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":00D6
            Left            =   1560
            List            =   "frmEditor_Event.frx":00D8
            TabIndex        =   146
            Text            =   "cmbLevelReqOperator"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtBranchLevelReq 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            TabIndex        =   145
            Text            =   "0"
            Top             =   2040
            Width           =   855
         End
         Begin VB.ComboBox cmbBranchSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":00DA
            Left            =   1560
            List            =   "frmEditor_Event.frx":00DC
            TabIndex        =   144
            Text            =   "cmbBranchSwitch"
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox cmbBranchSwitchReq 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":00DE
            Left            =   3480
            List            =   "frmEditor_Event.frx":00E8
            TabIndex        =   143
            Text            =   "cmbBranchSwitchReq"
            Top             =   600
            Width           =   1815
         End
         Begin VB.ComboBox cmbBranchItem 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":00F9
            Left            =   1560
            List            =   "frmEditor_Event.frx":00FB
            TabIndex        =   142
            Text            =   "cmbBranchItem"
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox cmbBranchClass 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":00FD
            Left            =   1560
            List            =   "frmEditor_Event.frx":00FF
            TabIndex        =   141
            Text            =   "cmbBranchClass"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.ComboBox cmbBranchSkill 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0101
            Left            =   1560
            List            =   "frmEditor_Event.frx":0103
            TabIndex        =   140
            Text            =   "cmbBranchSkill"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblNegative 
            Caption         =   "Negative: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   161
            Top             =   3120
            Width           =   5295
         End
         Begin VB.Label lblPositive 
            Caption         =   "Positive: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   2520
            Width           =   5295
         End
         Begin VB.Label lblRandomLabel 
            Alignment       =   2  'Center
            Caption         =   "is"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   157
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Alignment       =   2  'Center
            Caption         =   "is"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   156
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.Frame fraChangeSkill 
         Caption         =   "Mudar skill do player"
         Height          =   3735
         Left            =   120
         TabIndex        =   163
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbChangeSkills 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0105
            Left            =   720
            List            =   "frmEditor_Event.frx":0107
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Top             =   360
            Width           =   4695
         End
         Begin VB.OptionButton optChangeSkills 
            Caption         =   "Learn"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   165
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optChangeSkills 
            Caption         =   "Remove"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   164
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Skill:"
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   167
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraSpawnNPC 
         Caption         =   "Spawn"
         Height          =   3735
         Left            =   120
         TabIndex        =   178
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbSpawnNPC 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0109
            Left            =   120
            List            =   "frmEditor_Event.frx":010B
            Style           =   2  'Dropdown List
            TabIndex        =   179
            Top             =   480
            Width           =   5295
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "NPC:"
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   180
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame fraChangePK 
         Caption         =   "PK"
         Height          =   3735
         Left            =   120
         TabIndex        =   175
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.OptionButton optChangePK 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   177
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton optChangePK 
            Caption         =   "Yes"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   176
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraChangeSprite 
         Caption         =   "Sprite"
         Height          =   3735
         Left            =   120
         TabIndex        =   171
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlChangeSprite 
            Height          =   255
            Left            =   1200
            Max             =   100
            TabIndex        =   172
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label lblChangeSprite 
            Caption         =   "Sprite: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   173
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraOpenEvent 
         Caption         =   "Event"
         Height          =   3735
         Left            =   120
         TabIndex        =   236
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbOpenEventType 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":010D
            Left            =   3120
            List            =   "frmEditor_Event.frx":011A
            TabIndex        =   244
            Text            =   "cmbOpenEventType"
            Top             =   1200
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.OptionButton optOpenEventType 
            Caption         =   "Close"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   242
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.OptionButton optOpenEventType 
            Caption         =   "Open"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   241
            Top             =   1200
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.HScrollBar scrlOpenEventY 
            Height          =   255
            Left            =   1560
            Max             =   255
            TabIndex        =   239
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.HScrollBar scrlOpenEventX 
            Height          =   255
            Left            =   2040
            Max             =   1000
            TabIndex        =   237
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label lblOpenEventY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   240
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblOpenEventX 
            Caption         =   "Event: None."
            Height          =   255
            Left            =   120
            TabIndex        =   238
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame fraCustomScript 
         Caption         =   "Miscellaneos"
         Height          =   3735
         Left            =   120
         TabIndex        =   232
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.HScrollBar scrlCustomScript 
            Height          =   255
            Left            =   1560
            Max             =   255
            TabIndex        =   233
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblCustomScript 
            Caption         =   "Case: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   234
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraSetAccess 
         Caption         =   "Set access"
         Height          =   3735
         Left            =   120
         TabIndex        =   228
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbSetAccess 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0150
            Left            =   240
            List            =   "frmEditor_Event.frx":0163
            Style           =   2  'Dropdown List
            TabIndex        =   229
            Top             =   360
            Width           =   5055
         End
      End
      Begin VB.Frame fraPlayBGM 
         Caption         =   "BGM"
         Height          =   3735
         Left            =   120
         TabIndex        =   224
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbPlayBGM 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01A6
            Left            =   120
            List            =   "frmEditor_Event.frx":01A8
            Style           =   2  'Dropdown List
            TabIndex        =   225
            Top             =   360
            Width           =   5295
         End
      End
      Begin VB.Frame fraPlaySound 
         Caption         =   "Sound"
         Height          =   3735
         Left            =   120
         TabIndex        =   221
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbPlaySound 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01AA
            Left            =   240
            List            =   "frmEditor_Event.frx":01AC
            Style           =   2  'Dropdown List
            TabIndex        =   222
            Top             =   360
            Width           =   5055
         End
      End
      Begin VB.Frame fraSpecialEffect 
         Caption         =   "Special Effect"
         Height          =   3735
         Left            =   120
         TabIndex        =   197
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbEffectType 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":01AE
            Left            =   1440
            List            =   "frmEditor_Event.frx":01C4
            TabIndex        =   198
            Text            =   "cmbEffectType"
            Top             =   360
            Width           =   3855
         End
         Begin VB.Frame fraMapOverlay 
            Caption         =   "Map Color"
            Height          =   2415
            Left            =   240
            TabIndex        =   207
            Top             =   840
            Visible         =   0   'False
            Width           =   5055
            Begin VB.HScrollBar scrlMapTintData 
               Height          =   255
               Index           =   2
               Left            =   120
               Max             =   255
               TabIndex        =   211
               Top             =   1200
               Width           =   1935
            End
            Begin VB.HScrollBar scrlMapTintData 
               Height          =   255
               Index           =   1
               Left            =   2640
               Max             =   255
               TabIndex        =   210
               Top             =   600
               Width           =   1935
            End
            Begin VB.HScrollBar scrlMapTintData 
               Height          =   255
               Index           =   0
               Left            =   120
               Max             =   255
               TabIndex        =   209
               Top             =   600
               Width           =   1935
            End
            Begin VB.HScrollBar scrlMapTintData 
               Height          =   255
               Index           =   3
               Left            =   2640
               Max             =   255
               TabIndex        =   208
               Top             =   1200
               Width           =   1935
            End
            Begin VB.Label lblMapTintData 
               Caption         =   "Blue: 0"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   215
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label lblMapTintData 
               Caption         =   "Green: 0"
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   214
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label lblMapTintData 
               Caption         =   "Red: 0"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   213
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label lblMapTintData 
               Caption         =   "Opacity: 0"
               Height          =   255
               Index           =   3
               Left            =   2640
               TabIndex        =   212
               Top             =   960
               Width           =   1935
            End
         End
         Begin VB.Frame fraSetFog 
            Caption         =   "Set fog"
            Height          =   2415
            Left            =   240
            TabIndex        =   200
            Top             =   840
            Visible         =   0   'False
            Width           =   5055
            Begin VB.HScrollBar ScrlFogData 
               Height          =   255
               Index           =   2
               Left            =   120
               Max             =   255
               TabIndex        =   203
               Top             =   1740
               Width           =   4815
            End
            Begin VB.HScrollBar ScrlFogData 
               Height          =   255
               Index           =   0
               Left            =   120
               Max             =   255
               TabIndex        =   202
               Top             =   600
               Width           =   4815
            End
            Begin VB.HScrollBar ScrlFogData 
               Height          =   255
               Index           =   1
               Left            =   120
               Max             =   255
               TabIndex        =   201
               Top             =   1170
               Width           =   4815
            End
            Begin VB.Label lblFogData 
               Caption         =   "Fog Opacity: 0"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   206
               Top             =   1500
               Width           =   1815
            End
            Begin VB.Label lblFogData 
               Caption         =   "Fog: None"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   205
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label lblFogData 
               Caption         =   "Fog Speed: 0"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   204
               Top             =   930
               Width           =   1815
            End
         End
         Begin VB.Frame fraSetWeather 
            Caption         =   "Weather"
            Height          =   2415
            Left            =   240
            TabIndex        =   216
            Top             =   840
            Visible         =   0   'False
            Width           =   5055
            Begin VB.ComboBox CmbWeather 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":020F
               Left            =   120
               List            =   "frmEditor_Event.frx":0228
               Style           =   2  'Dropdown List
               TabIndex        =   218
               Top             =   600
               Width           =   4695
            End
            Begin VB.HScrollBar scrlWeatherIntensity 
               Height          =   255
               Left            =   120
               Max             =   100
               TabIndex        =   217
               Top             =   1320
               Width           =   4695
            End
            Begin VB.Label lblRandomLabel 
               AutoSize        =   -1  'True
               Caption         =   "Weather Type"
               Height          =   195
               Index           =   43
               Left            =   120
               TabIndex        =   220
               Top             =   360
               Width           =   1020
            End
            Begin VB.Label lblWeatherIntensity 
               Caption         =   "Intensity: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   219
               Top             =   1080
               Width           =   1455
            End
         End
         Begin VB.Label Label6 
            Caption         =   "Effect Type:"
            Height          =   255
            Left            =   240
            TabIndex        =   199
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame fraChangeExp 
         Caption         =   "Exp"
         Height          =   3735
         Left            =   120
         TabIndex        =   190
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtExp 
            Height          =   285
            Left            =   120
            TabIndex        =   259
            Top             =   600
            Width           =   3135
         End
         Begin VB.OptionButton optExpAction 
            Caption         =   "Sub"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   194
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton optExpAction 
            Caption         =   "Add"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   193
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton optExpAction 
            Caption         =   "Set"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   192
            Top             =   960
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label lblChangeExp 
            Caption         =   "Exp:"
            Height          =   255
            Left            =   120
            TabIndex        =   191
            Top             =   360
            Width           =   3735
         End
      End
   End
   Begin VB.Frame fraLabeling 
      Caption         =   "Labeling Variables and Switches"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      Begin VB.Frame fraRenaming 
         Caption         =   "Renaming Variable/Switch"
         Height          =   7455
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   8895
         Begin VB.Frame fraRandom 
            Caption         =   "Editing Variable/Switch"
            Height          =   2295
            Index           =   10
            Left            =   1920
            TabIndex        =   18
            Top             =   2640
            Width           =   5055
            Begin VB.CommandButton cmdRename_Ok 
               Caption         =   "Ok"
               Height          =   375
               Left            =   2280
               TabIndex        =   21
               Top             =   1800
               Width           =   1215
            End
            Begin VB.CommandButton cmdRename_Cancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   3720
               TabIndex        =   20
               Top             =   1800
               Width           =   1215
            End
            Begin VB.TextBox txtRename 
               Height          =   375
               Left            =   120
               TabIndex        =   19
               Top             =   720
               Width           =   4815
            End
            Begin VB.Label lblEditing 
               Caption         =   "Naming Variable #1"
               Height          =   375
               Left            =   120
               TabIndex        =   22
               Top             =   360
               Width           =   4815
            End
         End
      End
      Begin VB.CommandButton cmbLabel_Ok 
         Caption         =   "OK"
         Height          =   375
         Left            =   5760
         TabIndex        =   170
         Top             =   7200
         Width           =   1455
      End
      Begin VB.CommandButton cmdLabel_Cancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7320
         TabIndex        =   169
         Top             =   7200
         Width           =   1455
      End
      Begin VB.ListBox lstVariables 
         Height          =   5520
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   4335
      End
      Begin VB.ListBox lstSwitches 
         Height          =   5520
         Left            =   4560
         TabIndex        =   25
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton cmdRenameVariable 
         Caption         =   "Rename Variable"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   6240
         Width           =   4335
      End
      Begin VB.CommandButton cmdRenameSwitch 
         Caption         =   "Rename Switch"
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   6240
         Width           =   4455
      End
      Begin VB.Label lblRandomLabel 
         Alignment       =   2  'Center
         Caption         =   "Player Variables"
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblRandomLabel 
         Alignment       =   2  'Center
         Caption         =   "Player Switches"
         Height          =   255
         Index           =   36
         Left            =   4560
         TabIndex        =   27
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.Frame fraCommands 
      Caption         =   "Add Command"
      Height          =   4695
      Left            =   9480
      TabIndex        =   113
      Top             =   4800
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton cmdAddOk 
         Caption         =   "Fechar"
         Height          =   375
         Left            =   120
         TabIndex        =   114
         Top             =   4200
         Width           =   5535
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3735
         Left            =   120
         TabIndex        =   115
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6588
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "1"
         TabPicture(0)   =   "frmEditor_Event.frx":0268
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdCommand(12)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdCommand(11)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdCommand(10)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdCommand(0)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdCommand(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdCommand(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdCommand(3)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdCommand(4)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdCommand(5)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmdCommand(6)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmdCommand(7)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "cmdCommand(8)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "cmdCommand(9)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "cmdCommand(13)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "cmdCommand(14)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "cmdCommand(15)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "cmdCommand(16)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "cmdCommand(17)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).ControlCount=   18
         TabCaption(1)   =   "2"
         TabPicture(1)   =   "frmEditor_Event.frx":0284
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "cmdCommand(18)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "cmdCommand(19)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdCommand(20)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdCommand(21)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmdCommand(22)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "cmdCommand(23)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "cmdCommand(24)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "cmdCommand(25)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "cmdCommand(26)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "cmdCommand(27)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "cmdCommand(28)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "cmdCommand(29)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "cmdCommand(30)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).ControlCount=   13
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Quest"
            Height          =   375
            Index           =   30
            Left            =   3360
            TabIndex        =   258
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change PK"
            Height          =   375
            Index           =   17
            Left            =   -71400
            TabIndex        =   245
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Open/Close"
            Height          =   375
            Index           =   29
            Left            =   1800
            TabIndex        =   243
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Misc"
            Height          =   375
            Index           =   28
            Left            =   1800
            TabIndex        =   235
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Set access"
            Height          =   375
            Index           =   27
            Left            =   1800
            TabIndex        =   231
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Stop BGM"
            Height          =   375
            Index           =   26
            Left            =   1800
            TabIndex        =   230
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Stop Sound"
            Height          =   375
            Index           =   25
            Left            =   1800
            TabIndex        =   227
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "BGM"
            Height          =   375
            Index           =   24
            Left            =   1800
            TabIndex        =   226
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Play sound"
            Height          =   375
            Index           =   23
            Left            =   240
            TabIndex        =   223
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Special"
            Height          =   375
            Index           =   22
            Left            =   240
            TabIndex        =   196
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change XP"
            Height          =   375
            Index           =   21
            Left            =   240
            TabIndex        =   195
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Sex"
            Height          =   375
            Index           =   20
            Left            =   240
            TabIndex        =   189
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Class"
            Height          =   375
            Index           =   19
            Left            =   240
            TabIndex        =   185
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Spawn NPC"
            Height          =   375
            Index           =   18
            Left            =   240
            TabIndex        =   181
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Chance Sprite"
            Height          =   375
            Index           =   16
            Left            =   -71400
            TabIndex        =   174
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Chance skill"
            Height          =   375
            Index           =   15
            Left            =   -71400
            TabIndex        =   168
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "IF"
            Height          =   375
            Index           =   14
            Left            =   -71400
            TabIndex        =   162
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Bubble"
            Height          =   375
            Index           =   13
            Left            =   -71400
            TabIndex        =   138
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Goto Event"
            Height          =   375
            Index           =   9
            Left            =   -73080
            TabIndex        =   128
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Teleport"
            Height          =   375
            Index           =   8
            Left            =   -73080
            TabIndex        =   127
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Animation"
            Height          =   375
            Index           =   7
            Left            =   -73080
            TabIndex        =   126
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change level"
            Height          =   375
            Index           =   6
            Left            =   -73080
            TabIndex        =   125
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change item"
            Height          =   375
            Index           =   5
            Left            =   -74760
            TabIndex        =   124
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Open bank"
            Height          =   375
            Index           =   4
            Left            =   -74760
            TabIndex        =   123
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Open Shop"
            Height          =   375
            Index           =   3
            Left            =   -74760
            TabIndex        =   122
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Exit event"
            Height          =   375
            Index           =   2
            Left            =   -74760
            TabIndex        =   121
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Choices"
            Height          =   375
            Index           =   1
            Left            =   -74760
            TabIndex        =   120
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Message"
            Height          =   375
            Index           =   0
            Left            =   -74760
            TabIndex        =   119
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Switch"
            Height          =   375
            Index           =   10
            Left            =   -73080
            TabIndex        =   118
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Var"
            Height          =   375
            Index           =   11
            Left            =   -73080
            TabIndex        =   117
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Add Text"
            Height          =   375
            Index           =   12
            Left            =   -71400
            TabIndex        =   116
            Top             =   480
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton cmdSwitchesVariables 
      Caption         =   "Switch/Variavel"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Info"
      Height          =   7695
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      Begin VB.Frame Frame1 
         Caption         =   "Tile"
         Height          =   1575
         Left            =   4560
         TabIndex        =   41
         Top             =   600
         Width           =   1335
         Begin VB.CheckBox chkWalkthrought 
            Caption         =   "Pass through"
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbTrigger 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":02A0
            Left            =   120
            List            =   "frmEditor_Event.frx":02AA
            TabIndex        =   42
            Text            =   "cmbTrigger"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Mark:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.Frame fraRandom 
         Caption         =   "Conditions"
         Height          =   1575
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   4335
         Begin VB.TextBox txtPlayerVariable 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   38
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cmbPlayerVar 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":02C3
            Left            =   1200
            List            =   "frmEditor_Event.frx":02C5
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkPlayerVar 
            Caption         =   "Var"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbPlayerSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":02C7
            Left            =   1200
            List            =   "frmEditor_Event.frx":02C9
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chkPlayerSwitch 
            Caption         =   "Switch"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox cmbHasItem 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":02CB
            Left            =   1200
            List            =   "frmEditor_Event.frx":02CD
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CheckBox chkHasItem 
            Caption         =   "Item"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   975
         End
         Begin VB.ComboBox cmbPlayerSwitchCompare 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":02CF
            Left            =   2880
            List            =   "frmEditor_Event.frx":02D9
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox cmbPlayerVarCompare 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":02EA
            Left            =   2880
            List            =   "frmEditor_Event.frx":02EC
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblRandomLabel 
            Alignment       =   2  'Center
            Caption         =   "é"
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   40
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Alignment       =   2  'Center
            Caption         =   "="
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   39
            Top             =   795
            Width           =   255
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Commands"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   6840
         Width           =   5775
         Begin VB.CommandButton cmdSubEventEdit 
            Caption         =   "Edit"
            Height          =   375
            Left            =   2280
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdSubEventUp 
            Caption         =   "/\"
            Height          =   375
            Left            =   5160
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdSubEventDown 
            Caption         =   "\/"
            Height          =   375
            Left            =   4560
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdSubEventRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1200
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdSubEventAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   5055
      End
      Begin VB.ListBox lstSubEvents 
         Height          =   4545
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7275
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ListIndex As Long

Private Sub cmbBranchClass_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbBranchClass.ListIndex + 1
End Sub

Private Sub cmbBranchItem_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbBranchItem.ListIndex + 1
End Sub

Private Sub cmbBranchSkill_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbBranchSkill.ListIndex + 1
End Sub

Private Sub cmbBranchSwitch_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(5) = cmbBranchSwitch.ListIndex + 1
End Sub

Private Sub cmbBranchSwitchReq_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbBranchSwitchReq.ListIndex
End Sub

Private Sub cmbBranchVar_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(6) = cmbBranchVar.ListIndex + 1
End Sub

Private Sub cmbEffectType_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    fraSetFog.visible = False
    fraSetWeather.visible = False
    fraMapOverlay.visible = False
    Select Case cmbEffectType.ListIndex
        Case 3: fraSetFog.visible = True
        Case 4: fraSetWeather.visible = True
        Case 5: fraMapOverlay.visible = True
    End Select
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbEffectType.ListIndex
End Sub

Private Sub cmbHasItem_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).HasItemIndex = cmbHasItem.ListIndex + 1
End Sub

Private Sub cmbChangeClass_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbChangeClass.ListIndex + 1
End Sub

Private Sub cmbChangeSkills_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbChangeSkills.ListIndex + 1
End Sub

Private Sub cmbChatBubbleTarget_click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbChatBubbleTarget.ListIndex + 1
End Sub

Private Sub cmbLabel_Ok_Click()
    fraLabeling.visible = False
    SendSwitchesAndVariables
End Sub

Private Sub cmbLevelReqOperator_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(5) = cmbLevelReqOperator.ListIndex
End Sub

Private Sub cmbPlayBGM_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = Trim$(musicCache(cmbPlayBGM.ListIndex + 1))
End Sub

Private Sub cmbPlayerSwitch_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).SwitchIndex = cmbPlayerSwitch.ListIndex
End Sub

Private Sub cmbPlayerSwitchCompare_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).SwitchCompare = cmbPlayerSwitchCompare.ListIndex
End Sub

Private Sub cmbPlayerSwitchSet_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbPlayerSwitchSet.ListIndex
End Sub

Private Sub cmbPlayerVar_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).VariableIndex = cmbPlayerVar.ListIndex
End Sub

Private Sub cmbPlayerVarCompare_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).VariableCompare = cmbPlayerVarCompare.ListIndex
End Sub

Private Sub cmbPlaySound_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = Trim$(soundCache(cmbPlaySound.ListIndex + 1))
End Sub

Private Sub cmbSetAccess_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbSetAccess.ListIndex
End Sub

Private Sub cmbSpawnNPC_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbSpawnNPC.ListIndex + 1
End Sub

Private Sub cmbState_Click()
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbState.ListIndex
End Sub

Private Sub cmbSwitch_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbSwitch.ListIndex + 1
End Sub

Private Sub cmbTrigger_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).Trigger = cmbTrigger.ListIndex
End Sub

Private Sub cmbVariable_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbVariable.ListIndex + 1
End Sub

Private Sub cmbVarReqOperator_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(5) = cmbVarReqOperator.ListIndex
End Sub

Private Sub CmbWeather_click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = CmbWeather.ListIndex
End Sub

Private Sub cmdAddMenuOption_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Dim optIdx As Long
    With Events(EditorIndex).SubEvents(ListIndex)
        ReDim Preserve .Data(1 To UBound(.Data) + 1)
        ReDim Preserve .Text(1 To UBound(.Data) + 1)
        .Data(UBound(.Data)) = 1
    End With
    lstMenuOptions.AddItem ": " & 1
End Sub

Private Sub cmdAddOk_Click()
    fraCommands.visible = False
End Sub

Private Sub cmdCommand_Click(Index As Integer)
    Dim Count As Long
    If Not (Events(EditorIndex).HasSubEvents) Then
        ReDim Events(EditorIndex).SubEvents(1 To 1)
        Events(EditorIndex).HasSubEvents = True
    Else
        Count = UBound(Events(EditorIndex).SubEvents) + 1
        ReDim Preserve Events(EditorIndex).SubEvents(1 To Count)
    End If
    Call Events_SetSubEventType(EditorIndex, UBound(Events(EditorIndex).SubEvents), Index)
    Call PopulateSubEventList
    fraCommands.visible = False
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex <= 0 Or EditorIndex > MAX_EVENTS Then Exit Sub
    ListIndex = 0
    ClearEvent EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Trim$(Events(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    Event_Changed(EditorIndex) = True
    EventEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdEditOk_Click()
    Call PopulateSubEventList
    fraEditCommand.visible = False
End Sub

Private Sub cmdLabel_Cancel_Click()
    fraLabeling.visible = False
    RequestSwitchesAndVariables
End Sub

Private Sub cmdModifyMenuOption_Click()
    Dim tempIndex As Long, optIdx As Long
    tempIndex = lstSubEvents.ListIndex + 1
    optIdx = lstMenuOptions.ListIndex + 1
    If optIdx < 1 Or optIdx > UBound(Events(EditorIndex).SubEvents(ListIndex).Data) Then Exit Sub
    
    Events(EditorIndex).SubEvents(ListIndex).Text(optIdx + 1) = Trim$(txtMenuOptText.Text)
    Events(EditorIndex).SubEvents(ListIndex).Data(optIdx) = scrlMenuOptDest.Value
    lstMenuOptions.List(optIdx - 1) = Trim$(txtMenuOptText.Text) & ": " & scrlMenuOptDest.Value
End Sub

Private Sub cmdRemoveMenuOption_Click()
    Dim Index As Long, i As Long
    
    Index = lstMenuOptions.ListIndex + 1
    If Index > 0 And Index < lstMenuOptions.ListCount And lstMenuOptions.ListCount > 0 Then
        For i = Index + 1 To lstMenuOptions.ListCount
            Events(EditorIndex).SubEvents(ListIndex).Data(i - 1) = Events(EditorIndex).SubEvents(ListIndex).Data(i)
            Events(EditorIndex).SubEvents(ListIndex).Text(i) = Events(EditorIndex).SubEvents(ListIndex).Text(i + 1)
        Next i
        ReDim Preserve Events(EditorIndex).SubEvents(ListIndex).Data(1 To UBound(Events(EditorIndex).SubEvents(ListIndex).Data) - 1)
        ReDim Preserve Events(EditorIndex).SubEvents(ListIndex).Text(1 To UBound(Events(EditorIndex).SubEvents(ListIndex).Text) - 1)
        Call PopulateSubEventConfig
    End If
End Sub

Private Sub cmdRename_Cancel_Click()
    Dim i As Long
    fraRenaming.visible = False
    RenameType = 0
    RenameIndex = 0
    lstSwitches.Clear
    For i = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(i) & ". " & Trim$(Switches(i))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For i = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(i) & ". " & Trim$(Variables(i))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRename_Ok_Click()
    Dim i As Long
    Select Case RenameType
        Case 1
            'Variable
            If RenameIndex > 0 And RenameIndex <= MAX_VARIABLES + 1 Then
                Variables(RenameIndex) = txtRename.Text
                fraRenaming.visible = False
                RenameType = 0
                RenameIndex = 0
            End If
        Case 2
            'Switch
            If RenameIndex > 0 And RenameIndex <= MAX_SWITCHES + 1 Then
                Switches(RenameIndex) = txtRename.Text
                fraRenaming.visible = False
                RenameType = 0
                RenameIndex = 0
            End If
    End Select
    
    lstSwitches.Clear
    For i = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(i) & ". " & Trim$(Switches(i))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For i = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(i) & ". " & Trim$(Variables(i))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRenameSwitch_Click()
    If lstSwitches.ListIndex > -1 And lstSwitches.ListIndex < MAX_SWITCHES Then
        fraRenaming.visible = True
        lblEditing.Caption = "Editing Switch #" & CStr(lstSwitches.ListIndex + 1)
        txtRename.Text = Switches(lstSwitches.ListIndex + 1)
        RenameType = 2
        RenameIndex = lstSwitches.ListIndex + 1
    End If
End Sub

Private Sub cmdRenameVariable_Click()
    If lstVariables.ListIndex > -1 And lstVariables.ListIndex < MAX_VARIABLES Then
        fraRenaming.visible = True
        lblEditing.Caption = "Editing Variable #" & CStr(lstVariables.ListIndex + 1)
        txtRename.Text = Variables(lstVariables.ListIndex + 1)
        RenameType = 1
        RenameIndex = lstVariables.ListIndex + 1
    End If
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call EventEditorOk
    ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call EventEditorCancel
    ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub cmdSubEventAdd_Click()
    fraCommands.visible = True
End Sub

Private Sub cmdSubEventDown_Click()
    Dim Index As Long
    Index = lstSubEvents.ListIndex + 1
    If Index > 0 And Index < lstSubEvents.ListCount Then
        Dim temp As SubEventRec
        temp = Events(EditorIndex).SubEvents(Index)
        Events(EditorIndex).SubEvents(Index) = Events(EditorIndex).SubEvents(Index + 1)
        Events(EditorIndex).SubEvents(Index + 1) = temp
        Call PopulateSubEventList
    End If
End Sub

Private Sub cmdSubEventEdit_Click()
    If ListIndex > 0 And ListIndex <= lstSubEvents.ListCount Then
        fraEditCommand.visible = True
        PopulateSubEventConfig
    End If
End Sub

Private Sub cmdSubEventRemove_Click()
    Dim Index As Long, i As Long
    If ListIndex > 0 And ListIndex <= lstSubEvents.ListCount Then
        For i = ListIndex + 1 To lstSubEvents.ListCount
            Events(EditorIndex).SubEvents(i - 1) = Events(EditorIndex).SubEvents(i)
        Next i
        If lstSubEvents.ListCount = 1 Then
            Events(EditorIndex).HasSubEvents = False
            Erase Events(EditorIndex).SubEvents
        Else
            ReDim Preserve Events(EditorIndex).SubEvents(1 To lstSubEvents.ListCount - 1)
        End If
        Call PopulateSubEventList
    End If
End Sub

Private Sub cmdSubEventUp_Click()
    Dim Index As Long
    Index = lstSubEvents.ListIndex + 1
    If Index > 1 And Index <= lstSubEvents.ListCount Then
        Dim temp As SubEventRec
        temp = Events(EditorIndex).SubEvents(Index)
        Events(EditorIndex).SubEvents(Index) = Events(EditorIndex).SubEvents(Index - 1)
        Events(EditorIndex).SubEvents(Index - 1) = temp
        Call PopulateSubEventList
    End If
End Sub

Private Sub cmdSwitchesVariables_Click()
Dim i As Long
    fraLabeling.visible = True
    lstSwitches.Clear
    For i = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(i) & ". " & Trim$(Switches(i))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For i = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(i) & ". " & Trim$(Variables(i))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Dim i As Long, cap As Long
    'Move windows to right places
    frmEditor_Events.Width = 9600
    frmEditor_Events.Height = 8835
    fraEditCommand.Left = 232
    fraEditCommand.Top = 152
    fraCommands.Left = 232
    fraCommands.Top = 152
    fraLabeling.Width = 609
    fraLabeling.Height = 513
    
    ListIndex = 0

    scrlOpenShop.max = MAX_SHOPS
    scrlGiveItemID.max = MAX_ITEMS
    scrlPlayAnimationAnim.max = MAX_ANIMATIONS
    scrlWarpMap.max = MAX_MAPS
    scrlMessageSprite.max = NumCharacters
    ScrlFogData(0).max = NumFogs
    
    cmbLevelReqOperator.Clear
    cmbPlayerVarCompare.Clear
    cmbVarReqOperator.Clear
    For i = 0 To ComparisonOperator_Count - 1
        cmbLevelReqOperator.AddItem GetComparisonOperatorName(i)
        cmbPlayerVarCompare.AddItem GetComparisonOperatorName(i)
        cmbVarReqOperator.AddItem GetComparisonOperatorName(i)
    Next
    
    cmbHasItem.Clear
    cmbBranchItem.Clear
    For i = 1 To MAX_ITEMS
        cmbHasItem.AddItem Trim$(Item(i).Name)
        cmbBranchItem.AddItem Trim$(Item(i).Name)
    Next
    
    cmbSwitch.Clear
    cmbPlayerSwitch.Clear
    cmbBranchSwitch.Clear
    For i = 1 To MAX_SWITCHES
        cmbSwitch.AddItem i & ". " & Switches(i)
        cmbPlayerSwitch.AddItem i & ". " & Switches(i)
        cmbBranchSwitch.AddItem i & ". " & Switches(i)
    Next
    
    cmbVariable.Clear
    cmbPlayerVar.Clear
    cmbBranchVar.Clear
    For i = 1 To MAX_VARIABLES
        cmbVariable.AddItem i & ". " & Variables(i)
        cmbPlayerVar.AddItem i & ". " & Variables(i)
        cmbBranchVar.AddItem i & ". " & Variables(i)
    Next
    
    cmbBranchClass.Clear
    cmbChangeClass.Clear
    For i = 1 To Max_Classes
        cmbBranchClass.AddItem Trim$(Class(i).Name)
        cmbChangeClass.AddItem Trim$(Class(i).Name)
    Next
    
    cmbBranchSkill.Clear
    cmbChangeSkills.Clear
    For i = 1 To MAX_SPELLS
        cmbBranchClass.AddItem Trim$(Spell(i).Name)
        cmbChangeSkills.AddItem Trim$(Spell(i).Name)
    Next
    
    cmbChatBubbleTarget.Clear
    cmbSpawnNPC.Clear
    For i = 1 To MAX_MAP_NPCS
        If Map.Npc(i) <= 0 Then
            cmbChatBubbleTarget.AddItem CStr(i) & ". "
            cmbSpawnNPC.AddItem CStr(i) & ". "
        Else
            cmbChatBubbleTarget.AddItem CStr(i) & ". " & Trim$(Npc(Map.Npc(i)).Name)
            cmbSpawnNPC.AddItem CStr(i) & ". " & Trim$(Npc(Map.Npc(i)).Name)
        End If
    Next
    
    cmbPlaySound.Clear
    For i = 1 To UBound(soundCache)
        cmbPlaySound.AddItem (soundCache(i))
        cmbPlayBGM.AddItem (soundCache(i))
    Next
    
    cmbPlayBGM.Clear
    For i = 1 To UBound(musicCache)
        cmbPlayBGM.AddItem (musicCache(i))
    Next
End Sub

Private Sub chkHasItem_Click()
    If chkHasItem.Value = 0 Then cmbHasItem.Enabled = False Else cmbHasItem.Enabled = True
    Events(EditorIndex).chkHasItem = chkHasItem.Value
End Sub

Private Sub chkPlayerSwitch_Click()
    If chkPlayerSwitch.Value = 0 Then
        cmbPlayerSwitch.Enabled = False
        cmbPlayerSwitchCompare.Enabled = False
    Else
        cmbPlayerSwitch.Enabled = True
        cmbPlayerSwitchCompare.Enabled = True
    End If
    Events(EditorIndex).chkSwitch = chkPlayerSwitch.Value
End Sub

Private Sub chkPlayerVar_Click()
    If chkPlayerVar.Value = 0 Then
        cmbPlayerVar.Enabled = False
        txtPlayerVariable.Enabled = False
        cmbPlayerVarCompare.Enabled = False
    Else
        cmbPlayerVar.Enabled = True
        txtPlayerVariable.Enabled = True
        cmbPlayerVarCompare.Enabled = True
    End If
    Events(EditorIndex).chkVariable = chkPlayerVar.Value
End Sub

Private Sub chkWalkthrought_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).WalkThrought = chkWalkthrought.Value
End Sub

Private Sub lstIndex_Click()
    EventEditorInit
End Sub

Private Sub lstMenuOptions_Click()
    Dim tempIndex As Long, optIdx As Long
    tempIndex = lstSubEvents.ListIndex + 1
    optIdx = lstMenuOptions.ListIndex + 1
    If optIdx < 1 Or optIdx > UBound(Events(EditorIndex).SubEvents(ListIndex).Data) Then Exit Sub
    
    txtMenuOptText.Text = Trim$(Events(EditorIndex).SubEvents(ListIndex).Text(optIdx + 1))
    If Events(EditorIndex).SubEvents(ListIndex).Data(optIdx) <= 0 Then Events(EditorIndex).SubEvents(ListIndex).Data(optIdx) = 1
    
    If scrlMenuOptDest.max >= Events(EditorIndex).SubEvents(ListIndex).Data(optIdx) Then
        scrlMenuOptDest.Value = Events(EditorIndex).SubEvents(ListIndex).Data(optIdx)
    Else
        scrlMenuOptDest.Value = 1
    End If
End Sub

Private Sub lstSubEvents_Click()
    ListIndex = lstSubEvents.ListIndex + 1
    If ListIndex > 0 And ListIndex < lstSubEvents.ListCount Then
        cmdSubEventDown.Enabled = True
    Else
        cmdSubEventDown.Enabled = False
    End If
    If ListIndex > 1 And ListIndex <= lstSubEvents.ListCount Then
        cmdSubEventUp.Enabled = True
    Else
        cmdSubEventUp.Enabled = False
    End If
End Sub

Private Sub lstSubEvents_DblClick()
    If ListIndex > 0 And ListIndex <= lstSubEvents.ListCount Then
        fraEditCommand.visible = True
        PopulateSubEventConfig
    End If
End Sub

Private Sub optCondition_Index_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub

    cmbBranchVar.Enabled = False
    cmbVarReqOperator.Enabled = False
    txtBranchVarReq.Enabled = False
    cmbBranchSwitch.Enabled = False
    cmbBranchSwitchReq.Enabled = False
    cmbBranchItem.Enabled = False
    txtBranchItemAmount.Enabled = False
    cmbBranchClass.Enabled = False
    cmbBranchSkill.Enabled = False
    cmbLevelReqOperator.Enabled = False
    txtBranchLevelReq.Enabled = False
    
    Select Case Index
        Case 0
            cmbBranchVar.Enabled = True
            cmbVarReqOperator.Enabled = True
            txtBranchVarReq.Enabled = True
        Case 1
            cmbBranchSwitch.Enabled = True
            cmbBranchSwitchReq.Enabled = True
        Case 2
            cmbBranchItem.Enabled = True
            txtBranchItemAmount.Enabled = True
        Case 3
            cmbBranchClass.Enabled = True
        Case 4
            cmbBranchSkill.Enabled = True
        Case 5
            cmbLevelReqOperator.Enabled = True
            txtBranchLevelReq.Enabled = True
    End Select

    Events(EditorIndex).SubEvents(ListIndex).Data(1) = Index
End Sub

Private Sub optExpAction_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
End Sub

Private Sub optChangePK_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = Index
End Sub

Private Sub optChangeSex_Click(Index As Integer)
 If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = Index
End Sub

Private Sub optChangeSkills_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
End Sub

Private Sub optChannel_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
End Sub

Private Sub optChatBubbleTarget_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If Index = 0 Then
        cmbChatBubbleTarget.visible = False
    ElseIf Index = 1 Then
        cmbChatBubbleTarget.visible = True
    End If
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = Index
End Sub

Private Sub optItemOperation_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = Index
End Sub

Private Sub optLevelAction_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
End Sub

Private Sub optOpenEventType_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = Index
End Sub

Private Sub optVariableAction_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
    Select Case Index
        Case 0, 1, 2
            txtVariableData(0).Enabled = True
            txtVariableData(1).Enabled = False
            txtVariableData(2).Enabled = False
        Case 3
            txtVariableData(0).Enabled = False
            txtVariableData(1).Enabled = True
            txtVariableData(2).Enabled = True
    End Select
End Sub

Private Sub scrlAddText_Colour_Change()
    frmEditor_Events.lblAddText_Colour.Caption = "Color: " & GetColorString(frmEditor_Events.scrlAddText_Colour.Value)
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlAddText_Colour.Value
End Sub

Private Sub scrlCustomScript_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblCustomScript.Caption = "Script: " & scrlCustomScript.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlCustomScript.Value
End Sub

Private Sub ScrlFogData_Change(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Select Case Index
        Case 0
            lblFogData(Index).Caption = "Fog: " & ScrlFogData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(2) = ScrlFogData(Index).Value
        Case 1
            lblFogData(Index).Caption = "Speed: " & ScrlFogData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(3) = ScrlFogData(Index).Value
        Case 2
            lblFogData(Index).Caption = "Opacity: " & ScrlFogData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(4) = ScrlFogData(Index).Value
    End Select
End Sub

Private Sub scrlGiveItemAmount_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblGiveItemAmount.Caption = "Amount: " & scrlGiveItemAmount.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlGiveItemAmount.Value
End Sub

Private Sub scrlGiveItemID_Change()
    Dim tempIndex As Long
    On Error Resume Next
    tempIndex = lstSubEvents.ListIndex + 1
    lblGiveItemID.Caption = "Item: " & scrlGiveItemID.Value & "-" & Item(scrlGiveItemID.Value).Name
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlGiveItemID.Value
End Sub

Private Sub scrlGOTO_Change()
    lblGOTO.Caption = "Line Number: " & scrlGOTO.Value
    Dim tempIndex As Long
    tempIndex = lstSubEvents.ListIndex + 1
    
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlGOTO.Value
End Sub

Private Sub scrlChangeLevel_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeLevel.Caption = "Level: " & scrlChangeLevel.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlChangeLevel.Value
End Sub

Private Sub scrlChangeSprite_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeSprite.Caption = "Sprite: " & scrlChangeSprite.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlChangeSprite.Value
End Sub

Private Sub scrlMapTintData_Change(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Select Case Index
        Case 0
            lblMapTintData(Index).Caption = "Red: " & scrlMapTintData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlMapTintData(Index).Value
        Case 1
            lblMapTintData(Index).Caption = "Green: " & scrlMapTintData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlMapTintData(Index).Value
        Case 2
            lblMapTintData(Index).Caption = "Blue: " & scrlMapTintData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(4) = scrlMapTintData(Index).Value
        Case 3
            lblMapTintData(Index).Caption = "Opacity: " & scrlMapTintData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(5) = scrlMapTintData(Index).Value
    End Select
End Sub

Private Sub scrlMenuOptDest_Change()
    lblMenuOptDest.Caption = "Destination: " & scrlMenuOptDest.Value
End Sub

Private Sub scrlMessageSprite_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If scrlMessageSprite.Value = 0 Then
        lblMessageSprite.Caption = "Sprite: Player"
    Else
        lblMessageSprite.Caption = "Sprite: " & scrlMessageSprite.Value
    End If
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlMessageSprite.Value
End Sub

Private Sub scrlOpenEventX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If scrlOpenEventX.Value > 0 Then
        lblOpenEventX.Caption = "Event: " & scrlOpenEventX.Value & " " & Trim$(Events(scrlOpenEventX.Value).Name)
    Else
        lblOpenEventX.Caption = "Event: " & scrlOpenEventX.Value & " Nenhum"
    End If
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlOpenEventX.Value
End Sub
Private Sub scrlOpenEventY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblOpenEventY.Caption = "Y: " & scrlOpenEventY.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlOpenEventY.Value
End Sub

Private Sub scrlOpenShop_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblOpenShop.Caption = "Shop: " & scrlOpenShop.Value & "-" & Shop(scrlOpenShop.Value).Name
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlOpenShop.Value
End Sub

Private Sub scrlPlayAnimationAnim_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblPlayAnimationAnim.Caption = "Animation: " & scrlPlayAnimationAnim.Value & "-" & Animation(scrlPlayAnimationAnim.Value).Name
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlPlayAnimationAnim.Value
End Sub

Private Sub scrlPlayAnimationX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If scrlPlayAnimationX.Value >= 0 Then
        lblPlayAnimationX.Caption = "X: " & scrlPlayAnimationX.Value
    Else
        lblPlayAnimationX.Caption = "X: Player's X Position"
    End If
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlPlayAnimationX.Value
End Sub

Private Sub scrlPlayAnimationY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If scrlPlayAnimationY.Value >= 0 Then
        lblPlayAnimationY.Caption = "Y: " & scrlPlayAnimationY.Value
    Else
        lblPlayAnimationY.Caption = "Y: Player's Y Position"
    End If
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlPlayAnimationY.Value
End Sub

Private Sub scrlPositive_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblPositive.Caption = "Positive: " & scrlPositive.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlPositive.Value
End Sub
Private Sub scrlNegative_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblNegative.Caption = "Negative: " & scrlNegative.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(4) = scrlNegative.Value
End Sub

Private Sub scrlQuest_Change()
    If scrlQuest.Value > 0 Then
        lblQuest.Caption = "Quest: " & Trim$(Quest(scrlQuest.Value).Name)
    Else
        lblQuest.Caption = "Quest: Nenhuma"
    End If
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlQuest.Value
End Sub

Private Sub scrlWarpMap_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWarpMap.Caption = "Map: " & scrlWarpMap.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlWarpMap.Value
End Sub

Private Sub scrlWarpX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWarpX.Caption = "X: " & scrlWarpX.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlWarpX.Value
End Sub

Private Sub scrlWarpY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWarpY.Caption = "Y: " & scrlWarpY.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlWarpY.Value
End Sub

Private Sub scrlWeatherIntensity_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWeatherIntensity.Caption = "Intensity: " & scrlWeatherIntensity.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlWeatherIntensity.Value
End Sub

Private Sub Set_Click(Index As Integer)

End Sub

Private Sub txtAddText_Text_Change()
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = Trim$(txtAddText_Text.Text)
End Sub

Private Sub txtBranchItemAmount_Change()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(5) = Val(txtBranchItemAmount.Text)
End Sub

Private Sub txtBranchLevelReq_Change()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Val(txtBranchLevelReq.Text)
End Sub

Private Sub txtBranchVarReq_Change()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Val(txtBranchVarReq.Text)
End Sub

Private Sub txtChatbubbleText_Change()
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = txtChatbubbleText.Text
End Sub

Private Sub txtEXP_Change()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If IsNumeric(txtEXP) Then
        Events(EditorIndex).SubEvents(ListIndex).Data(1) = txtEXP
    End If
End Sub

Private Sub txtMenuQuery_Change()
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = txtMenuQuery.Text
End Sub

Private Sub txtName_Change()
    If EditorIndex <= 0 Or EditorIndex > MAX_EVENTS Then Exit Sub
    Events(EditorIndex).Name = txtName.Text
End Sub

Public Sub PopulateSubEventList()
    Dim tempIndex As Long, i As Long
    tempIndex = lstSubEvents.ListIndex
    
    lstSubEvents.Clear
    If Events(EditorIndex).HasSubEvents Then
        For i = 1 To UBound(Events(EditorIndex).SubEvents)
            lstSubEvents.AddItem i & ": " & GetEventTypeName(EditorIndex, i)
        Next i
    End If
    cmdSubEventRemove.Enabled = Events(EditorIndex).HasSubEvents
    
    If tempIndex >= 0 And tempIndex < lstSubEvents.ListCount - 1 Then lstSubEvents.ListIndex = tempIndex
    Call PopulateSubEventConfig
End Sub

Public Sub PopulateSubEventConfig()
    Dim i As Long, cap As Long
    If Not (fraEditCommand.visible) Then Exit Sub
    If ListIndex = 0 Then Exit Sub
    On Error Resume Next
    HideMenus
    'Ensure Capacity
    Call Events_SetSubEventType(EditorIndex, ListIndex, Events(EditorIndex).SubEvents(ListIndex).Type)
    
    With Events(EditorIndex).SubEvents(ListIndex)
        Select Case .Type
            Case Evt_Message
                txtPlayerText.Text = Trim$(.Text(1))
                scrlMessageSprite.Value = .Data(1)
                fraPlayerText.visible = True
            Case Evt_Menu
                txtMenuQuery.Text = Trim$(.Text(1))
                lstMenuOptions.Clear
                For i = 2 To UBound(.Text)
                    lstMenuOptions.AddItem Trim$(.Text(i)) & ": " & .Data(i - 1)
                Next i
                scrlMenuOptDest.max = UBound(Events(EditorIndex).SubEvents)
                fraMenu.visible = True
            Case Evt_OpenShop
                If .Data(1) < 1 Or .Data(1) > MAX_SHOPS Then .Data(1) = 1
                
                scrlOpenShop.Value = .Data(1)
                Call scrlOpenShop_Change
                fraOpenShop.visible = True
            Case Evt_GiveItem
                If .Data(1) < 1 Or .Data(1) > MAX_ITEMS Then .Data(1) = 1
                If .Data(2) < 1 Then .Data(2) = 1
                optItemOperation(.Data(3)).Value = True
                scrlGiveItemID.Value = .Data(1)
                scrlGiveItemAmount.Value = .Data(2)
                Call scrlGiveItemID_Change
                Call scrlGiveItemAmount_Change
                fraGiveItem.visible = True
            Case Evt_PlayAnimation
                If .Data(1) < 1 Or .Data(1) > MAX_ANIMATIONS Then .Data(1) = 1
                
                scrlPlayAnimationAnim.Value = .Data(1)
                scrlPlayAnimationX.Value = .Data(2)
                scrlPlayAnimationY.Value = .Data(3)
                Call scrlPlayAnimationAnim_Change
                Call scrlPlayAnimationX_Change
                Call scrlPlayAnimationY_Change
                fraAnimation.visible = True
            Case Evt_Warp
                If .Data(1) < 1 Or .Data(1) > MAX_MAPS Then .Data(1) = 1
                
                scrlWarpMap.Value = .Data(1)
                scrlWarpX.Value = .Data(2)
                scrlWarpY.Value = .Data(3)
                Call scrlWarpMap_Change
                Call scrlWarpX_Change
                Call scrlWarpY_Change
                fraMapWarp.visible = True
            Case Evt_GOTO
                If .Data(1) < 1 Or .Data(1) > UBound(Events(EditorIndex).SubEvents) Then .Data(1) = 1
                
                scrlGOTO.max = UBound(Events(EditorIndex).SubEvents)
                scrlGOTO.Value = .Data(1)
                Call scrlGOTO_Change
                fraGoTo.visible = True
            Case Evt_Switch
                cmbSwitch.ListIndex = .Data(1) - 1
                cmbPlayerSwitchSet.ListIndex = .Data(2)
                fraChangeSwitch.visible = True
            Case Evt_Variable
                cmbVariable.ListIndex = .Data(1) - 1
                optVariableAction(.Data(2)).Value = True
                If .Data(2) = 3 Then
                    txtVariableData(1) = .Data(3)
                    txtVariableData(2) = .Data(4)
                Else
                    txtVariableData(0) = .Data(3)
                End If
                fraChangeVariable.visible = True
            Case Evt_AddText
                txtAddText_Text.Text = Trim$(.Text(1))
                scrlAddText_Colour.Value = .Data(1)
                optChannel(.Data(2)).Value = True
                fraAddText.visible = True
            Case Evt_Chatbubble
                txtChatbubbleText.Text = Trim$(.Text(1))
                optChatBubbleTarget(.Data(1)).Value = True
                cmbChatBubbleTarget.ListIndex = .Data(2) - 1
                fraChatbubble.visible = True
            Case Evt_Branch
                scrlPositive.max = UBound(Events(EditorIndex).SubEvents)
                scrlNegative.max = UBound(Events(EditorIndex).SubEvents)
                scrlPositive.Value = .Data(3)
                scrlNegative.Value = .Data(4)
                optCondition_Index(.Data(1)) = True
                Select Case .Data(1)
                    Case 0
                        cmbBranchVar.ListIndex = .Data(6) - 1
                        txtBranchVarReq.Text = .Data(2)
                        cmbVarReqOperator.ListIndex = .Data(5)
                    Case 1
                        cmbBranchSwitch.ListIndex = .Data(5) - 1
                        cmbBranchSwitchReq.ListIndex = .Data(2)
                    Case 2
                        cmbBranchItem.ListIndex = .Data(2) - 1
                        txtBranchItemAmount.Text = .Data(5)
                    Case 3
                        cmbBranchClass.ListIndex = .Data(2) - 1
                    Case 4
                        cmbBranchSkill.ListIndex = .Data(2) - 1
                    Case 5
                        cmbLevelReqOperator.ListIndex = .Data(5)
                        txtBranchLevelReq.Text = .Data(2)
                End Select
                fraBranch.visible = True
            Case Evt_ChangeSkill
                cmbChangeSkills.ListIndex = .Data(1) - 1
                optChangeSkills(.Data(2)).Value = True
                fraChangeSkill.visible = True
            Case Evt_ChangeLevel
                scrlChangeLevel.Value = .Data(1)
                optLevelAction(.Data(2)).Value = True
                fraChangeLevel.visible = True
            Case Evt_ChangeSprite
                scrlChangeSprite.Value = .Data(1)
                fraChangeSprite.visible = True
            Case Evt_ChangePK
                optChangePK(.Data(1)).Value = True
                fraChangePK.visible = True
            Case Evt_SpawnNPC
                cmbSpawnNPC.ListIndex = .Data(1) - 1
                fraSpawnNPC.visible = True
            Case Evt_ChangeClass
                cmbChangeClass.ListIndex = .Data(1) - 1
                fraChangeClass.visible = True
            Case Evt_ChangeSex
                optChangeSex(.Data(1)).Value = True
                fraChangeSex.visible = True
            Case Evt_ChangeExp
                txtEXP.Text = .Data(1)
                optExpAction(.Data(2)).Value = True
                fraChangeExp.visible = True
            Case Evt_SpecialEffect
                cmbEffectType.ListIndex = .Data(1)
                Select Case .Data(1)
                    Case 3
                        ScrlFogData(0).Value = .Data(2)
                        ScrlFogData(1).Value = .Data(3)
                        ScrlFogData(2).Value = .Data(4)
                    Case 4
                        CmbWeather.ListIndex = .Data(2)
                        scrlWeatherIntensity.Value = .Data(3)
                    Case 5
                        scrlMapTintData(0).Value = .Data(2)
                        scrlMapTintData(1).Value = .Data(3)
                        scrlMapTintData(2).Value = .Data(4)
                        scrlMapTintData(3).Value = .Data(5)
                End Select
                fraSpecialEffect.visible = True
            Case Evt_PlaySound
                For i = 1 To UBound(soundCache())
                    If soundCache(i) = Trim$(.Text(1)) Then
                        cmbPlaySound.ListIndex = i - 1
                    End If
                Next
                fraPlaySound.visible = True
            Case Evt_PlayBGM
                For i = 1 To UBound(musicCache())
                    If musicCache(i) = Trim$(.Text(1)) Then
                        cmbPlayBGM.ListIndex = i - 1
                    End If
                Next
                fraPlayBGM.visible = True
            Case Evt_SetAccess
                cmbSetAccess.ListIndex = .Data(1)
                fraSetAccess.visible = True
            Case Evt_CustomScript
                scrlCustomScript.Value = .Data(1)
                fraCustomScript.visible = True
            Case Evt_OpenEvent
                scrlOpenEventX.Value = .Data(1)
                scrlOpenEventY.Value = .Data(2)
                optOpenEventType(.Data(3)).Value = True
                cmbOpenEventType.ListIndex = .Data(4)
                fraOpenEvent.visible = True
            Case Evt_Quest
                scrlQuest.Value = .Data(1)
                cmbState.ListIndex = .Data(2)
                frmQuest.visible = True
        End Select
    End With
End Sub
Private Sub HideMenus()
    fraPlayerText.visible = False
    fraMenu.visible = False
    fraOpenShop.visible = False
    fraGiveItem.visible = False
    fraAnimation.visible = False
    fraMapWarp.visible = False
    fraGoTo.visible = False
    fraChangeSwitch.visible = False
    fraChangeVariable.visible = False
    fraAddText.visible = False
    fraChatbubble.visible = False
    fraBranch.visible = False
    fraChangeLevel.visible = False
    fraChangeSkill.visible = False
    fraChangeSprite.visible = False
    fraChangePK.visible = False
    fraSpawnNPC.visible = False
    fraChangeClass.visible = False
    fraChangeSex.visible = False
    fraSpecialEffect.visible = False
    fraPlaySound.visible = False
    fraPlayBGM.visible = False
    fraSetAccess.visible = False
    fraCustomScript.visible = False
    fraOpenEvent.visible = False
    fraChangeExp.visible = False
    frmQuest.visible = False
End Sub

Private Sub txtPlayerText_Change()
    Dim tempIndex As Long
    tempIndex = lstSubEvents.ListIndex + 1
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = txtPlayerText.Text
End Sub

Private Sub txtPlayerVariable_Change()
    Events(EditorIndex).VariableCondition = Val(txtPlayerVariable.Text)
End Sub

Private Sub txtVariableData_Change(Index As Integer)
    Select Case Index
        Case 0: Events(EditorIndex).SubEvents(ListIndex).Data(3) = Val(txtVariableData(0))
        Case 1: Events(EditorIndex).SubEvents(ListIndex).Data(3) = Val(txtVariableData(1))
        Case 2: Events(EditorIndex).SubEvents(ListIndex).Data(4) = Val(txtVariableData(2))
    End Select
End Sub
