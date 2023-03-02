VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10335
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSpellLinear 
      Caption         =   "Linear"
      Height          =   3135
      Left            =   6840
      TabIndex        =   95
      Top             =   3600
      Visible         =   0   'False
      Width           =   3255
      Begin VB.HScrollBar scrlLargura 
         Height          =   255
         Left            =   120
         Max             =   2
         TabIndex        =   110
         Top             =   2280
         Width           =   3015
      End
      Begin VB.HScrollBar scrlSpellLinearAnim 
         Height          =   255
         Index           =   0
         Left            =   120
         Max             =   1000
         TabIndex        =   99
         Top             =   480
         Width           =   3015
      End
      Begin VB.HScrollBar scrlSpellLinearAnim 
         Height          =   255
         Index           =   1
         Left            =   120
         Max             =   1000
         TabIndex        =   98
         Top             =   1080
         Width           =   3015
      End
      Begin VB.HScrollBar scrlSpellLinearAnim 
         Height          =   255
         Index           =   2
         Left            =   120
         Max             =   1000
         TabIndex        =   97
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salvar"
         Height          =   375
         Left            =   1800
         TabIndex        =   96
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblLargura 
         Caption         =   "Largura:"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblSpellLinearAnim 
         Caption         =   "Base:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblSpellLinearAnim 
         Caption         =   "Corpo:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   101
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblSpellLinearAnim 
         Caption         =   "Cabeça:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   100
         Top             =   1440
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "List"
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6900
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   7335
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      Begin VB.Frame frmTrans 
         Caption         =   "Data"
         Height          =   6375
         Left            =   3480
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
         Begin VB.Frame Frame4 
            Caption         =   "Vital"
            Height          =   1215
            Left            =   120
            TabIndex        =   79
            Top             =   3960
            Width           =   3015
            Begin VB.HScrollBar scrlTransVital 
               Height          =   135
               Index           =   1
               Left            =   120
               Max             =   100
               TabIndex        =   83
               Top             =   960
               Width           =   2775
            End
            Begin VB.HScrollBar scrlTransVital 
               Height          =   135
               Index           =   0
               Left            =   120
               Max             =   100
               TabIndex        =   81
               Top             =   480
               Width           =   2775
            End
            Begin VB.Label lblTransVital 
               Caption         =   "MP:"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   82
               Top             =   720
               Width           =   2775
            End
            Begin VB.Label lblTransVital 
               Caption         =   "HP:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   80
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.HScrollBar scrlChangeHair 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   78
            Top             =   5520
            Width           =   3015
         End
         Begin VB.HScrollBar scrlPDL 
            Height          =   255
            Left            =   120
            Max             =   5000
            TabIndex        =   74
            Top             =   1680
            Width           =   3015
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Close"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   6000
            Width           =   3015
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   135
            Index           =   4
            Left            =   120
            Max             =   200
            Min             =   -200
            TabIndex        =   71
            Top             =   3720
            Width           =   3015
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   135
            Index           =   3
            Left            =   120
            Max             =   200
            Min             =   -200
            TabIndex        =   69
            Top             =   3360
            Width           =   3015
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   135
            Index           =   2
            Left            =   120
            Max             =   200
            Min             =   -200
            TabIndex        =   67
            Top             =   3000
            Width           =   3015
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   135
            Index           =   1
            Left            =   120
            Max             =   200
            Min             =   -200
            TabIndex        =   65
            Top             =   2640
            Width           =   3015
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   135
            Index           =   0
            Left            =   120
            Max             =   200
            Min             =   -200
            TabIndex        =   63
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlTransAnim 
            Height          =   255
            Left            =   120
            Max             =   200
            TabIndex        =   61
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlSprite 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   59
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblChangeHair 
            Caption         =   "Mudar cabelo:"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   5280
            Width           =   3015
         End
         Begin VB.Label lblPDL 
            Caption         =   "PDL bonus:"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label lblStat 
            Caption         =   "Técnica:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   70
            Top             =   3480
            Width           =   3015
         End
         Begin VB.Label lblStat 
            Caption         =   "Destreza:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   68
            Top             =   3120
            Width           =   3015
         End
         Begin VB.Label lblStat 
            Caption         =   "KI:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   66
            Top             =   2760
            Width           =   3015
         End
         Begin VB.Label lblStat 
            Caption         =   "Constituição:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   64
            Top             =   2400
            Width           =   3015
         End
         Begin VB.Label lblStat 
            Caption         =   "Força:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   2040
            Width           =   3015
         End
         Begin VB.Label lblTransAnim 
            Caption         =   "Animation:"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblSprite 
            Caption         =   "Avermelhar sprite:"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Projectile"
         Height          =   255
         Left            =   2040
         TabIndex        =   85
         Top             =   6240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Projectile"
         Height          =   1815
         Left            =   120
         TabIndex        =   84
         Top             =   4320
         Visible         =   0   'False
         Width           =   3255
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
            Left            =   2640
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   90
            Top             =   240
            Width           =   480
         End
         Begin VB.HScrollBar scrlRotate 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   89
            Top             =   1440
            Width           =   3015
         End
         Begin VB.HScrollBar scrlProjectile 
            Height          =   255
            Left            =   120
            Max             =   50
            TabIndex        =   87
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblRotate 
            Caption         =   "Rotação:"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label lblProjectile 
            Caption         =   "Imagem:"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   600
            Width           =   3015
         End
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   6840
         Width           =   1215
      End
      Begin VB.TextBox txtDesc 
         Height          =   495
         Left            =   1440
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   6720
         Width           =   5295
      End
      Begin VB.Frame Frame2 
         Caption         =   "Basic"
         Height          =   5895
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3255
         Begin VB.CommandButton Command6 
            Caption         =   ">"
            Height          =   255
            Left            =   2880
            TabIndex        =   111
            Top             =   2040
            Width           =   255
         End
         Begin VB.CommandButton Command4 
            Caption         =   "..."
            Height          =   255
            Left            =   2880
            TabIndex        =   93
            Top             =   480
            Width           =   255
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
            Left            =   2640
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   48
            Top             =   5160
            Width           =   480
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   5400
            Width           =   2415
         End
         Begin VB.HScrollBar scrlCool 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   31
            Top             =   4680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   29
            Top             =   4080
            Width           =   3015
         End
         Begin VB.ComboBox cmbClass 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   3480
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   25
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   23
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMP 
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   3015
         End
         Begin VB.ComboBox cmbType 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":0000
            Left            =   120
            List            =   "frmEditor_Spell.frx":001F
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label lblIcon 
            Caption         =   "Icon: None"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   5160
            Width           =   3015
         End
         Begin VB.Label lblCool 
            Caption         =   "Cooldown Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   4440
            Width           =   2535
         End
         Begin VB.Label lblCast 
            Caption         =   "Casting Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Class:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label lblAccess 
            Caption         =   "Access Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblMP 
            Caption         =   "MP Cost: None"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data"
         Height          =   6375
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   3255
         Begin VB.Frame frmAoE 
            Caption         =   "AoE Effects"
            Height          =   1695
            Left            =   120
            TabIndex        =   103
            Top             =   4560
            Visible         =   0   'False
            Width           =   3015
            Begin VB.CommandButton Command5 
               Caption         =   "Salvar"
               Height          =   255
               Left            =   1560
               TabIndex        =   108
               Top             =   1200
               Width           =   1335
            End
            Begin VB.HScrollBar scrlEffectTick 
               Height          =   135
               Left            =   120
               Max             =   100
               TabIndex        =   107
               Top             =   960
               Width           =   2775
            End
            Begin VB.HScrollBar scrlAoeDuration 
               Height          =   135
               Left            =   120
               Max             =   100
               TabIndex        =   105
               Top             =   480
               Width           =   2775
            End
            Begin VB.Label lblEffectTick 
               Caption         =   "Effect Tick:"
               Height          =   255
               Left            =   120
               TabIndex        =   106
               Top             =   720
               Width           =   2775
            End
            Begin VB.Label lblAoEDuration 
               Caption         =   "Duration: "
               Height          =   255
               Left            =   120
               TabIndex        =   104
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.TextBox txtPrice 
            Height          =   270
            Left            =   2040
            TabIndex        =   94
            Top             =   1680
            Width           =   1095
         End
         Begin VB.HScrollBar scrlUpgrade 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   17
            Top             =   480
            Width           =   3015
         End
         Begin VB.HScrollBar scrlPlayerAnim 
            Height          =   135
            Left            =   120
            Max             =   5
            TabIndex        =   76
            Top             =   5520
            Width           =   2895
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlEffect 
            Height          =   135
            Left            =   120
            Max             =   5
            TabIndex        =   55
            Top             =   4800
            Width           =   2895
         End
         Begin VB.HScrollBar scrlStun 
            Height          =   135
            Left            =   120
            TabIndex        =   49
            Top             =   5160
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnim 
            Height          =   135
            Left            =   120
            TabIndex        =   47
            Top             =   4440
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnimCast 
            Height          =   135
            Left            =   120
            TabIndex        =   45
            Top             =   4080
            Width           =   2895
         End
         Begin VB.CheckBox chkAOE 
            Caption         =   "AoE"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   3120
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAOE 
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   3600
            Width           =   3015
         End
         Begin VB.HScrollBar scrlDuration 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   35
            Top             =   1680
            Width           =   1815
         End
         Begin VB.HScrollBar scrlVital 
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.HScrollBar scrlRequisite 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   19
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   15
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.HScrollBar scrlImpact 
            Height          =   135
            Left            =   120
            Max             =   8
            TabIndex        =   92
            Top             =   5880
            Width           =   2895
         End
         Begin VB.Label lblUpgrade 
            Caption         =   "Evolução: Nenhuma"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label lblImpact 
            Caption         =   "Impacto: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   5640
            Width           =   2895
         End
         Begin VB.Label lblPlayerAnim 
            Caption         =   "Player Anim: None"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label lblEffect 
            AutoSize        =   -1  'True
            Caption         =   "Effect: None"
            Height          =   180
            Left            =   120
            TabIndex        =   56
            Top             =   4560
            Width           =   930
         End
         Begin VB.Label lblStun 
            Caption         =   "Stun Duration: None"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   4920
            Width           =   2895
         End
         Begin VB.Label lblAnim 
            Caption         =   "Animation: None"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   4200
            Width           =   2895
         End
         Begin VB.Label lblAnimCast 
            Caption         =   "Cast Anim: None"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   3840
            Width           =   2895
         End
         Begin VB.Label lblAOE 
            Caption         =   "AoE: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   3360
            Width           =   3015
         End
         Begin VB.Label lblRange 
            Caption         =   "Range: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   2640
            Width           =   3015
         End
         Begin VB.Label lblInterval 
            Caption         =   "Price:"
            Height          =   255
            Left            =   2040
            TabIndex        =   36
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblDuration 
            Caption         =   "Item: Nenhum"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblVital 
            Caption         =   "Vital: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   2040
            Width           =   3015
         End
         Begin VB.Label lblDir 
            Caption         =   "Dir: Down"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblRequisite 
            Caption         =   "Requisito:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   6240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmTrans.visible = False
End Sub

Private Sub Command2_Click()
    fraSpellLinear.visible = False
End Sub

Private Sub Command3_Click()
    Frame5.visible = Not Frame5.visible
End Sub

Private Sub Command4_Click()
    Dim i As String
    i = InputBox("Copy spell:", "Universo Z")
    i = Val(i)
    If IsNumeric(i) Then
        If i > 0 And i <= MAX_SPELLS Then
            Spell(EditorIndex) = Spell(i)
            SpellEditorInit
            'SpellEditorInit
        End If
    End If
End Sub

Private Sub Command5_Click()
    frmAoE.visible = False
End Sub

Private Sub Command6_Click()
    If Spell(EditorIndex).Requisite > 0 Then
        scrlLevel.Value = Item(Spell(EditorIndex).Requisite).LevelReq
    End If
End Sub

Private Sub Form_Load()
    ' set max values
    scrlAnimCast.max = MAX_ANIMATIONS
    scrlAnim.max = MAX_ANIMATIONS
    scrlAOE.max = MAX_BYTE
    scrlRange.max = MAX_BYTE
    scrlMap.max = MAX_MAPS
    scrlEffect.max = MAX_EFFECTS
    scrlTransAnim.max = MAX_ANIMATIONS
    scrlIcon.max = NumSpellIcons
End Sub

Private Sub chkAOE_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkAOE.Value = 0 Then
        Spell(EditorIndex).IsAoE = False
    Else
        Spell(EditorIndex).IsAoE = True
    End If
    
    frmAoE.visible = Spell(EditorIndex).IsAoE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkAOE_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClass_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).ClassReq = cmbClass.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClass_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).Type = cmbType.ListIndex
    
    frmTrans.visible = (Spell(EditorIndex).Type = SPELL_TYPE_TRANS)
    fraSpellLinear.visible = (Spell(EditorIndex).Type = SPELL_TYPE_LINEAR)
    Command3.visible = (Spell(EditorIndex).Type = SPELL_TYPE_DAMAGEHP)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearSpell EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccess_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAccess.Value > 0 Then
        lblAccess.Caption = "Access req: " & scrlAccess.Value
    Else
        lblAccess.Caption = "Access req: None"
    End If
    Spell(EditorIndex).AccessReq = scrlAccess.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccess_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnim.Value > 0 Then
        lblAnim.Caption = "Animation: " & Trim$(Animation(scrlAnim.Value).Name)
    Else
        lblAnim.Caption = "Animation: None"
    End If
    Spell(EditorIndex).SpellAnim = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnimCast.Value > 0 Then
        lblAnimCast.Caption = "Cast Anim: " & Trim$(Animation(scrlAnimCast.Value).Name)
    Else
        lblAnimCast.Caption = "Cast Anim: None"
    End If
    Spell(EditorIndex).CastAnim = scrlAnimCast.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAOE_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAOE.Value > 0 Then
        lblAOE.Caption = "Area: " & scrlAOE.Value & " tiles."
    Else
        lblAOE.Caption = "Area: Apenas em si"
    End If
    Spell(EditorIndex).AoE = scrlAOE.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAOE_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAoeDuration_Change()
    lblAoEDuration.Caption = "Duration: " & (scrlAoeDuration.Value * 100) & "ms"
    Spell(EditorIndex).AoEDuration = (scrlAoeDuration.Value * 100)
End Sub

Private Sub scrlCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCast.Caption = "Casting: " & scrlCast.Value & "s"
    Spell(EditorIndex).CastTime = scrlCast.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub scrlChangeHair_Change()
    Dim HairStyle As String

    Spell(EditorIndex).HairChange = scrlChangeHair.Value
    
    Select Case scrlChangeHair.Value
        Case 0: HairStyle = "Normal"
        Case 1: HairStyle = "Super sayan 1"
        Case 2: HairStyle = "Super sayan 2"
        Case 3: HairStyle = "Super sayan 3"
        Case 4: HairStyle = "Super sayan 4"
        Case 5: HairStyle = "Oozaru"
    End Select
    
    lblChangeHair.Caption = "Hair: " & HairStyle
End Sub

Private Sub scrlCool_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCool.Caption = "Cooldown: " & scrlCool.Value & "s"
    Spell(EditorIndex).CDTime = scrlCool.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCool_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDir_Change()
Dim sDir As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlDir.Value
        Case DIR_UP
            sDir = "Cima"
        Case DIR_DOWN
            sDir = "Baixo"
        Case DIR_RIGHT
            sDir = "Direita"
        Case DIR_LEFT
            sDir = "Esquerda"
    End Select
    lblDir.Caption = "Dir: " & sDir
    Spell(EditorIndex).Dir = scrlDir.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDir_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDuration_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlDuration.Value > 0 Then
        If Trim$(Item(scrlDuration.Value).Name) <> "" Then
            lblDuration.Caption = "Item: (" & scrlDuration.Value & ") " & Trim$(Item(scrlDuration.Value).Name)
        Else
            lblDuration.Caption = "Item: Nenhum"
        End If
    Else
        lblDuration.Caption = "Item: Nenhum"
    End If
    Spell(EditorIndex).Item = scrlDuration.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDuration_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    Spell(EditorIndex).Effect = scrlEffect.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEffect_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlEffectTick_Change()
    lblEffectTick.Caption = "Effect Tick: " & (scrlEffectTick.Value * 100) & "ms"
    Spell(EditorIndex).AoETick = (scrlEffectTick.Value * 100)
End Sub

Private Sub scrlIcon_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlIcon.Value > 0 Then
        lblIcon.Caption = "Icon: " & scrlIcon.Value
    Else
        lblIcon.Caption = "Icon: None"
    End If
    Spell(EditorIndex).Icon = scrlIcon.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlIcon_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlImpact_Change()
    lblImpact.Caption = "Impact: " & scrlImpact.Value
    Spell(EditorIndex).Impact = scrlImpact.Value
End Sub

Private Sub scrlLargura_Change()
    Spell(EditorIndex).LinearRange = scrlLargura.Value
    lblLargura.Caption = "Largura: " & scrlLargura.Value
End Sub

Private Sub scrlLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlLevel.Value > 0 Then
        lblLevel.Caption = "Level req: " & scrlLevel.Value
    Else
        lblLevel.Caption = "Level req: None"
    End If
    Spell(EditorIndex).LevelReq = scrlLevel.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevel_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMap_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblMap.Caption = "Mapa: " & scrlMap.Value
    Spell(EditorIndex).Map = scrlMap.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMap_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlMP.Value > 0 Then
        lblMP.Caption = "MP: " & scrlMP.Value
    Else
        lblMP.Caption = "MP: None"
    End If
    Spell(EditorIndex).MPCost = scrlMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMP_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPDL_Change()
    lblPDL.Caption = "PDL (PL) Bonus: " & scrlPDL.Value & "%"
    Spell(EditorIndex).PDLBonus = scrlPDL.Value
End Sub

Private Sub scrlPlayerAnim_Change()
    Dim Anim As String

    Spell(EditorIndex).CastPlayerAnim = scrlPlayerAnim.Value
    
    Select Case scrlPlayerAnim.Value
        Case 0: Anim = "None."
        Case 1: Anim = "Kamehameha"
        Case 2: Anim = "Spirit bomb"
        Case 3: Anim = "Big bang attack"
        Case 4: Anim = "Super sayajin"
        Case 5: Anim = "Normal attack"
    End Select
    
    lblPlayerAnim.Caption = "Animation attack: " & Anim
End Sub

Private Sub scrlProjectile_Change()
    lblProjectile.Caption = "Image: " & scrlProjectile.Value
    Spell(EditorIndex).Projectile = scrlProjectile.Value
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlRange.Value > 0 Then
        lblRange.Caption = "Range: " & scrlRange.Value & " tiles."
    Else
        lblRange.Caption = "Range: Self-cast"
    End If
    Spell(EditorIndex).Range = scrlRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRequisite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlRequisite.Value > 0 Then
        If Trim$(Item(scrlRequisite.Value).Name) <> "" Then
            lblRequisite.Caption = "Requisite: (" & scrlRequisite.Value & ") " & Trim$(Item(scrlRequisite.Value).Name)
        Else
            lblRequisite.Caption = "Requisite: Nenhum"
        End If
    Else
        lblRequisite.Caption = "Requisite: Nenhum"
    End If
    Spell(EditorIndex).Requisite = scrlRequisite.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlY_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRotate_Change()
    lblRotate.Caption = "Rotation: " & scrlRotate
    Spell(EditorIndex).RotateSpeed = scrlRotate
End Sub

Private Sub scrlSpellLinearAnim_Change(Index As Integer)
    Spell(EditorIndex).SpellLinearAnim(Index + 1) = scrlSpellLinearAnim(Index).Value
    
    Dim Caption As String
    
    Select Case Index
        Case 0: Caption = "Base: "
        Case 1: Caption = "Body: "
        Case 2: Caption = "Head: "
    End Select
    
    If scrlSpellLinearAnim(Index).Value > 0 And scrlSpellLinearAnim(Index).Value < MAX_ANIMATIONS Then
        lblSpellLinearAnim(Index).Caption = Caption & Animation(scrlSpellLinearAnim(Index).Value).Name
    Else
        lblSpellLinearAnim(Index).Caption = Caption & "None"
    End If
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = "Sprite red: " & scrlSprite.Value
    Spell(EditorIndex).SpriteTrans = scrlSprite.Value
End Sub

Private Sub scrlStat_Change(Index As Integer)
    Dim StatName As String
    
    Select Case Index
        Case 0: StatName = "Força (STR)"
        Case 1: StatName = "Constituição (CON)"
        Case 2: StatName = "KI"
        Case 3: StatName = "Destreza (DEX)"
        Case 4: StatName = "Técnica"
    End Select
    
    Spell(EditorIndex).Add_Stat(Index + 1) = scrlStat(Index).Value
    lblStat(Index).Caption = StatName & ": " & scrlStat(Index).Value
    
End Sub

Private Sub scrlStun_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlStun.Value > 0 Then
        lblStun.Caption = "Stun: " & scrlStun.Value & "s"
    Else
        lblStun.Caption = "Stuna: None"
    End If
    Spell(EditorIndex).StunDuration = scrlStun.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStun_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTransAnim_Change()
    If scrlTransAnim.Value > 0 Then
    lblTransAnim.Caption = "Animation: " & Trim$(Animation(scrlTransAnim.Value).Name)
    Spell(EditorIndex).TransAnim = scrlTransAnim.Value
    Else
    lblTransAnim.Caption = "Animation: None"
    Spell(EditorIndex).TransAnim = scrlTransAnim.Value
    End If
End Sub

Private Sub scrlTransVital_Change(Index As Integer)
    Dim VitalName As String
    Spell(EditorIndex).TransVital(Index + 1) = scrlTransVital(Index).Value
    
    Select Case Index
        Case 0: VitalName = "HP: "
        Case 1: VitalName = "MP: "
    End Select
    
    lblTransVital(Index).Caption = VitalName & scrlTransVital(Index).Value
    
End Sub

Private Sub scrlUpgrade_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlUpgrade.Value > 0 Then
        If Trim$(Spell(scrlUpgrade.Value).Name) <> "" Then
            lblUpgrade.Caption = "Evolution: (" & scrlUpgrade.Value & ") " & Trim$(Spell(scrlUpgrade.Value).Name)
        Else
            lblUpgrade.Caption = "Evolution: Nenhuma"
        End If
    Else
        lblUpgrade.Caption = "Evolução: Nenhuma"
    End If
    Spell(EditorIndex).Upgrade = scrlUpgrade.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlX_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlVital_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblVital.Caption = "Vital: " & scrlVital.Value
    Spell(EditorIndex).Vital = scrlVital.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlVital_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).Desc = txtDesc.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Spell(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Spell(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Spell(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtPrice_Change()
    If IsNumeric(txtPrice) Then
        If Val(txtPrice) >= 0 And Val(txtPrice) < MAX_LONG Then
            Spell(EditorIndex).Price = Val(txtPrice)
        End If
    End If
End Sub
