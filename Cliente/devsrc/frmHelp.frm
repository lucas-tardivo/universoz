VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEvents 
      Caption         =   "Events"
      Height          =   6135
      Left            =   2280
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Label lblEventInfo 
         Caption         =   "6th - Use ""map tile only"" options only for map tiles (set from Map Editor)"
         Height          =   375
         Index           =   8
         Left            =   480
         TabIndex        =   42
         Top             =   4440
         Width           =   4575
      End
      Begin VB.Label lblEventInfo 
         Caption         =   "4th - Optional - You can set conditions to run event at top of Event Editor."
         Height          =   375
         Index           =   6
         Left            =   480
         TabIndex        =   41
         Top             =   3480
         Width           =   4575
      End
      Begin VB.Label lblEventInfo 
         Caption         =   $"frmHelp.frx":0000
         Height          =   615
         Index           =   5
         Left            =   480
         TabIndex        =   40
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Label lblEventInfo 
         Caption         =   "2nd - Click -commands- EDIT and fill command details"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   39
         Top             =   2400
         Width           =   4695
      End
      Begin VB.Label lblEventInfo 
         Caption         =   "1st - Click -commands- ADD and choose command from list."
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   38
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label lblEventInfo 
         Caption         =   "How to use:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   37
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Label lblEventInfo 
         Alignment       =   2  'Center
         Caption         =   $"frmHelp.frx":00BD
         Height          =   855
         Index           =   1
         Left            =   480
         TabIndex        =   36
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblEventInfo 
         Alignment       =   2  'Center
         Caption         =   "Basic Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   35
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblEventInfo 
         Caption         =   "5th - To edit Switches and Variables, use button at bottom-left of Event editor"
         Height          =   375
         Index           =   7
         Left            =   480
         TabIndex        =   34
         Top             =   3960
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame fraHelpType 
      Caption         =   "Select help"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optHelp 
         Caption         =   "Map Report"
         Enabled         =   0   'False
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   11
         Top             =   3960
         Width           =   1815
      End
      Begin VB.OptionButton optHelp 
         Caption         =   "Spells"
         Enabled         =   0   'False
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   1815
      End
      Begin VB.OptionButton optHelp 
         Caption         =   "Shops"
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   1815
      End
      Begin VB.OptionButton optHelp 
         Caption         =   "Resources"
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   1815
      End
      Begin VB.OptionButton optHelp 
         Caption         =   "NPCs"
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   1815
      End
      Begin VB.OptionButton optHelp 
         Caption         =   "Map Properties"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1815
      End
      Begin VB.OptionButton optHelp 
         Caption         =   "Map Editor"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton optHelp 
         Caption         =   "Items"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton optHelp 
         Caption         =   "Events"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optHelp 
         Caption         =   "Effects"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optHelp 
         Caption         =   "Animations"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame fraEffects 
      Caption         =   "Effects"
      Height          =   6135
      Left            =   2280
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Label lblEffectInfo 
         Caption         =   "2nd - Choose Effect for each multi-particle"
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   32
         Top             =   5160
         Width           =   4695
      End
      Begin VB.Label lblEffectInfo 
         Caption         =   "1st - Set amount of multi-particles"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   31
         Top             =   4800
         Width           =   4695
      End
      Begin VB.Label lblEffectInfo 
         Caption         =   "How to use (multi-particle):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   30
         Top             =   4440
         Width           =   4695
      End
      Begin VB.Label lblEffectInfo 
         Caption         =   "5th - Fill ""Graphic Data"" with values what do you want."
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   29
         Top             =   4080
         Width           =   4575
      End
      Begin VB.Label lblEffectInfo 
         Alignment       =   2  'Center
         Caption         =   "Basic Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   28
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblEffectInfo 
         Alignment       =   2  'Center
         Caption         =   $"frmHelp.frx":0189
         Height          =   975
         Index           =   1
         Left            =   480
         TabIndex        =   27
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblEffectInfo 
         Caption         =   "How to use (single-particle):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   26
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label lblEffectInfo 
         Caption         =   "1st - You need to set the sprite, for every sprite effect is a little different."
         Height          =   495
         Index           =   3
         Left            =   480
         TabIndex        =   25
         Top             =   2160
         Width           =   4695
      End
      Begin VB.Label lblEffectInfo 
         Caption         =   "2nd - Choose basic effect type."
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   24
         Top             =   2640
         Width           =   4695
      End
      Begin VB.Label lblEffectInfo 
         Caption         =   "3rd - Set the amount of particles (more particles, more ""shiny and bigger"" effect."
         Height          =   495
         Index           =   5
         Left            =   480
         TabIndex        =   23
         Top             =   3000
         Width           =   4695
      End
      Begin VB.Label lblEffectInfo 
         Caption         =   "4th - Set size. Size is size of each particle (you specified Particle count above)"
         Height          =   375
         Index           =   6
         Left            =   480
         TabIndex        =   22
         Top             =   3600
         Width           =   4575
      End
   End
   Begin VB.Frame fraAnimations 
      Caption         =   "Animations"
      Height          =   6135
      Left            =   2280
      TabIndex        =   12
      Top             =   120
      Width           =   5655
      Begin VB.Label lblAnimInfo 
         Caption         =   "4th - Loop time is the time it takes to loop throught each frame."
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   19
         Top             =   3600
         Width           =   4575
      End
      Begin VB.Label lblAnimInfo 
         Caption         =   $"frmHelp.frx":02CE
         Height          =   735
         Index           =   5
         Left            =   480
         TabIndex        =   18
         Top             =   2880
         Width           =   4695
      End
      Begin VB.Label lblAnimInfo 
         Caption         =   "2nd - Set the amount of times you want the animation to loop."
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   17
         Top             =   2520
         Width           =   4695
      End
      Begin VB.Label lblAnimInfo 
         Caption         =   $"frmHelp.frx":036B
         Height          =   615
         Index           =   3
         Left            =   480
         TabIndex        =   16
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label lblAnimInfo 
         Caption         =   "How to use:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   15
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label lblAnimInfo 
         Alignment       =   2  'Center
         Caption         =   $"frmHelp.frx":03F4
         Height          =   735
         Index           =   1
         Left            =   480
         TabIndex        =   14
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblAnimInfo 
         Alignment       =   2  'Center
         Caption         =   "Basic Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    frmHelp.visible = False
End Sub

Private Sub optHelp_Click(Index As Integer)
    fraAnimations.visible = False
    fraEffects.visible = False
    fraEvents.visible = False
    
    Select Case Index
        Case 0: fraAnimations.visible = True
        Case 1: fraEffects.visible = True
        Case 2: fraEvents.visible = True
    End Select
End Sub
