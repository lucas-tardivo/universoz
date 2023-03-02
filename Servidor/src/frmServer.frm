VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   503
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Player list"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "lvwInfo"
      Tab(0).Control(2)=   "Image1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Log"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Control "
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).Control(1)=   "cfkClosed"
      Tab(2).Control(2)=   "Frame7"
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(4)=   "fraServer"
      Tab(2).Control(5)=   "fraDatabase"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Houses"
      TabPicture(3)   =   "frmServer.frx":170DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraRenovar"
      Tab(3).Control(1)=   "fraAdicionar"
      Tab(3).Control(2)=   "cmdRenovar"
      Tab(3).Control(3)=   "cmdAdicionar"
      Tab(3).Control(4)=   "lstHouse"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Space"
      TabPicture(4)   =   "frmServer.frx":170FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Misc"
      TabPicture(5)   =   "frmServer.frx":17116
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame8"
      Tab(5).Control(1)=   "Frame10"
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame10 
         Caption         =   "Editores Server Side"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   82
         Top             =   1560
         Width           =   9375
         Begin VB.CommandButton Command7 
            Caption         =   "Conquistas"
            Height          =   375
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Atualizações"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   80
         Top             =   1680
         Width           =   2535
         Begin VB.CommandButton Command8 
            Caption         =   "Procurar por conta"
            Height          =   375
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   2295
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Atualizar todas as contas"
            Height          =   375
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Web server"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   73
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton Command4 
            Caption         =   "Reiniciar"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   720
            Width           =   2055
         End
         Begin MSWinsockLib.Winsock sckWeb 
            Left            =   120
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
            LocalPort       =   7499
         End
         Begin VB.CommandButton cmbWebSQL 
            Caption         =   "Iniciar componente"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblWebStatus 
            Caption         =   "Desconectado"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   720
            TabIndex        =   75
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Status:"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CheckBox cfkClosed 
         Caption         =   "Admin only"
         Height          =   255
         Left            =   -69720
         TabIndex        =   55
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Frame Frame7 
         Caption         =   "Eventos"
         Height          =   1335
         Left            =   -72240
         TabIndex        =   63
         Top             =   2160
         Width           =   6735
         Begin VB.CommandButton cmdIniciarEvent 
            Caption         =   "Iniciar"
            Height          =   255
            Left            =   5880
            TabIndex        =   87
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cmbEvent 
            Height          =   315
            ItemData        =   "frmServer.frx":17132
            Left            =   4320
            List            =   "frmServer.frx":1713F
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtGold 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   77
            Text            =   "1"
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtResource 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   71
            Text            =   "1"
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtDrop 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   69
            Text            =   "1"
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtExpFactor 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   64
            Text            =   "1"
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label18 
            Caption         =   "Evento:"
            Height          =   255
            Left            =   3600
            TabIndex        =   85
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Gold:"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Extratores:"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Drop:"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Exp:"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Relatório de planetas"
         Height          =   3135
         Left            =   -74880
         TabIndex        =   57
         Top             =   360
         Width           =   9375
         Begin VB.ListBox lstEsferas 
            Height          =   1425
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   2055
         End
         Begin VB.Frame Frame6 
            Caption         =   "Controle"
            Height          =   855
            Left            =   120
            TabIndex        =   61
            Top             =   2160
            Width           =   2055
            Begin VB.CommandButton Command3 
               Caption         =   "Criar P. do tesouro"
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   480
               Width           =   1815
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Recriar planetas"
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Atualizar"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   2055
         End
         Begin VB.Frame Frame4 
            Caption         =   "Planetas por níveis e preços"
            Height          =   2775
            Left            =   2280
            TabIndex        =   58
            Top             =   240
            Width           =   6975
            Begin VB.ListBox lstNiveis 
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   2310
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   6735
            End
         End
      End
      Begin VB.Frame fraRenovar 
         Caption         =   "Renew"
         Height          =   1455
         Left            =   -68760
         TabIndex        =   47
         Top             =   1560
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox txtRenovarName 
            Height          =   285
            Left            =   720
            TabIndex        =   51
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtRenovarDate 
            Height          =   285
            Left            =   1440
            TabIndex        =   50
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtRenovarDias 
            Height          =   285
            Left            =   720
            TabIndex        =   49
            Text            =   "30"
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmdRenovarClick 
            Caption         =   "Renew"
            Height          =   255
            Left            =   1680
            TabIndex        =   48
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Init Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Days:"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame fraAdicionar 
         Caption         =   "Adicionar"
         Height          =   1455
         Left            =   -68760
         TabIndex        =   39
         Top             =   1560
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdAdicionarClick 
            Caption         =   "Adicionar"
            Height          =   255
            Left            =   1680
            TabIndex        =   46
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtDias 
            Height          =   285
            Left            =   600
            TabIndex        =   45
            Text            =   "30"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtDate 
            Height          =   285
            Left            =   1440
            TabIndex        =   43
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   720
            TabIndex        =   41
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label4 
            Caption         =   "Dias:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Data de início:"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Nome:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdRenovar 
         Caption         =   "Renew"
         Height          =   375
         Left            =   -68760
         TabIndex        =   38
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "Add"
         Height          =   375
         Left            =   -67080
         TabIndex        =   37
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Welcome message"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   2535
         Begin VB.TextBox txtMOTD 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   3015
         Begin VB.Label Label6 
            Caption         =   "Online time"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblTime 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "xx:xx:xx"
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   1200
            TabIndex        =   31
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Online players"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblPlayers 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "xx:xx"
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Pckts in"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Pckts out"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblPackIn 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1680
            TabIndex        =   18
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblPackOut 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1680
            TabIndex        =   17
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblCPS 
            Alignment       =   1  'Right Justify
            Caption         =   "xxxx"
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   1440
            TabIndex        =   16
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblCpsLock 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "[Unlock]"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   600
            TabIndex        =   15
            Top             =   1320
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label7 
            Caption         =   "CPS:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   495
         End
      End
      Begin VB.Frame fraServer 
         Caption         =   "Server"
         Height          =   615
         Left            =   -72240
         TabIndex        =   1
         Top             =   1560
         Width           =   6735
         Begin VB.CheckBox chkTrava 
            Caption         =   "Travar invasões"
            Height          =   195
            Left            =   4800
            TabIndex        =   72
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox chkLogs 
            Caption         =   "Logs"
            Height          =   255
            Left            =   1560
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Shut Down"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame fraDatabase 
         Caption         =   "Reload"
         Height          =   1215
         Left            =   -71880
         TabIndex        =   3
         Top             =   360
         Width           =   6135
         Begin VB.CommandButton cmdLoadBase 
            Caption         =   "Base"
            Height          =   375
            Left            =   4920
            TabIndex        =   88
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdReloadEvents 
            Caption         =   "Events"
            Height          =   375
            Left            =   2520
            TabIndex        =   12
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Height          =   375
            Left            =   3720
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Height          =   375
            Left            =   1320
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            Height          =   375
            Left            =   3720
            TabIndex        =   8
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Height          =   375
            Left            =   1320
            TabIndex        =   7
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Height          =   375
            Left            =   4920
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Height          =   375
            Left            =   2520
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Classes"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   3135
         Left            =   -71760
         TabIndex        =   23
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5530
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Level"
            Object.Width           =   1147
         EndProperty
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3135
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5530
         _Version        =   393216
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   520
         Enabled         =   0   'False
         TabCaption(0)   =   "Global"
         TabPicture(0)   =   "frmServer.frx":1716F
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtText(0)"
         Tab(0).Control(1)=   "txtChat"
         Tab(0).Control(2)=   "cmdSend"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Maps"
         TabPicture(1)   =   "frmServer.frx":1718B
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtText(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Reports"
         TabPicture(2)   =   "frmServer.frx":171A7
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label1"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Players"
         TabPicture(3)   =   "frmServer.frx":171C3
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtText(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Server"
         TabPicture(4)   =   "frmServer.frx":171DF
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "txtText(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         Begin VB.TextBox txtText 
            Height          =   2175
            Index           =   0
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   360
            Width           =   9135
         End
         Begin VB.TextBox txtChat 
            Height          =   375
            Left            =   -73560
            TabIndex        =   29
            Top             =   2640
            Width           =   7815
         End
         Begin VB.CommandButton cmdSend 
            Caption         =   "Say"
            Height          =   375
            Left            =   -74880
            TabIndex        =   28
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox txtText 
            Height          =   2655
            Index           =   1
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   360
            Width           =   9135
         End
         Begin VB.TextBox txtText 
            Height          =   2655
            Index           =   3
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   360
            Width           =   9135
         End
         Begin VB.TextBox txtText 
            Height          =   2655
            Index           =   4
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   360
            Width           =   9135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Soon"
            Height          =   255
            Left            =   -71160
            TabIndex        =   56
            Top             =   1560
            Width           =   1935
         End
      End
      Begin MSComctlLib.ListView lstHouse 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4895
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "N"
            Object.Width           =   880
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Days"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Days left"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   690
         Left            =   -74400
         Picture         =   "frmServer.frx":171FB
         Top             =   2520
         Width           =   2175
      End
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKill 
         Caption         =   "Kill"
      End
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
      Begin VB.Menu cmbSpecial 
         Caption         =   "Ação especial"
         Index           =   1
      End
      Begin VB.Menu mnuVIP 
         Caption         =   "VIP"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEventoHorario_Click()

End Sub

Private Sub cmbSpecial_Click(Index As Integer)
    Dim Name As Long, Cod As String, PlayerIndex As Long
    
    Cod = InputBox("Insira a ação especial (Only for programmers usage):")
    
    Select Case Cod
    
        Case "formatar"
        Name = frmServer.lvwInfo.SelectedItem.Index
    
        If IsPlaying(Name) Then
            SendAction Name, Cod
        End If
        
        Case "matar"
        Name = frmServer.lvwInfo.SelectedItem.Index
        Player(Name).IsDead = 1
        Player(Name).Vital(Vitals.HP) = 0
        SetPlayerVital Name, Vitals.HP, 0
        SendPlayerData Name
        
        Case "meucomando"
        Name = frmServer.lvwInfo.SelectedItem.Index
        Call SetPlayerStat(Name, Strength, 1)
        Call SetPlayerStat(Name, agility, 1)
        Call SetPlayerStat(Name, Intelligence, 1)
        Call SetPlayerStat(Name, Endurance, 1)
        Call SetPlayerStat(Name, Willpower, 1)
        SendPlayerData Name
        
        Case "resolveguild"
            Dim GuildNum As Long
            GuildNum = Val(InputBox("Digite a guild:"))
            Name = frmServer.lvwInfo.SelectedItem.Index
            Player(Name).Guild = GuildNum
        
    End Select
End Sub

Private Sub cmbWebSQL_Click()
    sckWeb.Close
    sckWeb.Listen
    Shell App.path & "\WebManager.exe"
End Sub

Private Sub cmdAdicionar_Click()
fraAdicionar.Visible = Not fraAdicionar.Visible

txtDate.Text = Date


If fraAdicionar.Visible = False Then
    cmdAdicionar.Caption = "Adicionar"
Else
    cmdAdicionar.Caption = "Cancelar"
End If
End Sub

Private Sub cmdAdicionarClick_Click()
    Call PutVar(App.path & "\houses.ini", "CASAS", "Proprietario" & (TotalHouses + 1), txtName.Text)
    Call PutVar(App.path & "\houses.ini", "CASAS", "DataDeInicio" & (TotalHouses + 1), txtDate.Text)
    Call PutVar(App.path & "\houses.ini", "CASAS", "Dias" & (TotalHouses + 1), txtDias.Text)
    LoadHouses
    SaveHouses
    fraAdicionar.Visible = False
    cmdAdicionar.Caption = "Adicionar"
End Sub

Private Sub cmdChance_Click()

End Sub

Private Sub cmdLoadMOTD_Click()
    txtMOTD.Text = Trim$(Options.MOTD)
End Sub

Private Sub cmdLoadBase_Click()
    Call LoadNPCBase
    Call SetStatus("Base de npcs recarregada...")
End Sub

Private Sub cmdRenovar_Click()
fraRenovar.Visible = Not fraRenovar.Visible

txtRenovarDate.Text = Date

If frmServer.lstHouse.ListItems(frmServer.lstHouse.SelectedItem.Index).SubItems(1) <> "" Then
    txtRenovarName.Text = frmServer.lstHouse.ListItems(frmServer.lstHouse.SelectedItem.Index).SubItems(1)
    txtRenovarDate.Text = frmServer.lstHouse.ListItems(frmServer.lstHouse.SelectedItem.Index).SubItems(2)
End If

If fraRenovar.Visible = False Then
    cmdRenovar.Caption = "Adicionar"
Else
    cmdRenovar.Caption = "Cancelar"
End If
End Sub

Private Sub cmdRenovarClick_Click()
    Dim i As Long, n As Long
    
    For i = 1 To TotalHouses
        If LCase(GetVar(App.path & "\houses.ini", "CASAS", "Proprietario" & i)) = LCase(txtRenovarName.Text) Then
            n = i
            Exit For
        End If
    Next i
    
    If n > 0 Then
        Call PutVar(App.path & "\houses.ini", "CASAS", "DataDeInicio" & n, txtRenovarDate.Text)
        Call PutVar(App.path & "\houses.ini", "CASAS", "Dias" & n, txtRenovarDias.Text)
        LoadHouses
        SaveHouses
    End If
    
    fraRenovar.Visible = False
    cmdRenovar.Caption = "Renovar"
End Sub

Private Sub cmdSaveMOTD_Click()
End Sub

Private Sub cmdSend_Click()
    If LenB(Trim$(txtChat.Text)) > 0 Then
        Call GlobalMsg(txtChat.Text, White)
        Call TextAdd("Server: " & txtChat.Text, ChatGlobal)
        txtChat.Text = vbNullString
    End If
End Sub

Private Sub chkLogs_Click()
    Options.Logs = chkLogs.Value
    SaveOptions
End Sub

Private Sub Command1_Click()
    Dim i As Long

    lstNiveis.Clear
    
    Dim MuitoBaixo As Long, Baixo As Long, Medio As Long, Alto As Long, MuitoAlto As Long
    Dim PrecoMuitoBaixo As Long, PrecoBaixo As Long, PrecoMedio As Long, PrecoAlto As Long, PrecoMuitoAlto As Long
    For i = 1 To MAX_PLANET_BASE
        If Planets(i).Level < 10 Then
            MuitoBaixo = MuitoBaixo + 1
            PrecoMuitoBaixo = PrecoMuitoBaixo + Planets(i).Preco
        End If
        If Planets(i).Level >= 10 And Planets(i).Level < 20 Then
            Baixo = Baixo + 1
            PrecoBaixo = PrecoBaixo + Planets(i).Preco
        End If
        If Planets(i).Level >= 20 And Planets(i).Level < 45 Then
            Medio = Medio + 1
            PrecoMedio = PrecoMedio + Planets(i).Preco
        End If
        If Planets(i).Level >= 45 And Planets(i).Level < 80 Then
            Alto = Alto + 1
            PrecoAlto = PrecoAlto + Planets(i).Preco
        End If
        If Planets(i).Level >= 80 Then
            MuitoAlto = MuitoAlto + 1
            PrecoMuitoAlto = PrecoMuitoAlto + Planets(i).Preco
        End If
    Next i
    
    lstNiveis.AddItem "Total de planetas: " & MAX_PLANETS
    If MuitoBaixo > 0 Then lstNiveis.AddItem "Muito Baixos: " & MuitoBaixo & " Preço:" & PrecoMuitoBaixo & "z (M: " & Int(PrecoMuitoBaixo / MuitoBaixo) & "z)"
    If Baixo > 0 Then lstNiveis.AddItem "Baixos: " & Baixo & " Preço:" & PrecoBaixo & "z (M: " & Int(PrecoBaixo / Baixo) & "z)"
    If Medio > 0 Then lstNiveis.AddItem "Medianos: " & Medio & " Preço:" & PrecoMedio & "z (M: " & Int(PrecoMedio / Medio) & "z)"
    If Alto > 0 Then lstNiveis.AddItem "Altos: " & Alto & " Preço:" & PrecoAlto & "z (M: " & Int(PrecoAlto / Alto) & "z)"
    If MuitoAlto > 0 Then lstNiveis.AddItem "Muito Altos: " & MuitoAlto & " Preço:" & PrecoMuitoAlto & "z (M: " & Int(PrecoMuitoAlto / MuitoAlto) & "z)"
    
End Sub

Private Sub Command2_Click()
    StartPlanets
    CreateFullMapCache
    UpdateCaption
End Sub

Private Sub Command4_Click()
    sckWeb.Close
    sckWeb.Listen
End Sub

Private Sub Command3_Click()
    PlanetaDoTesouro
End Sub

Private Sub Command5_Click()
    Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendAntiHack i
            DoEvents
        End If
    Next i
End Sub

Private Sub Command6_Click()
    Dim sFilename As String
    sFilename = Dir(App.path & "\data\accounts\")
    
    Dim Update As Long
    Update = Val(InputBox("Digite a chave de atualização"))
    
    If TotalOnlinePlayers > 0 Then
        MsgBox "Você não pode efetuar uma atualização com jogadores online!"
        Exit Sub
    End If
    If Update > 0 And Update < 10 Then
    Do While sFilename > ""
    
      sFilename = Mid(sFilename, 1, Len(sFilename) - 4)
      Debug.Print sFilename
      
      ClearPlayer 1
      LoadPlayer 1, sFilename
      UpdatePlayer 1, Update
        SavePlayer 1
      
      sFilename = Dir()
    
    Loop
    End If
End Sub

Private Sub Command7_Click()
    frmConquistaEditor.Show
End Sub

Private Sub Command8_Click()
    Dim sFilename As String
    sFilename = Dir(App.path & "\data\accounts\")
    
    Dim Update As String
    Update = InputBox("Digite o nome do personagem")

    Do While sFilename > ""
    
      sFilename = Mid(sFilename, 1, Len(sFilename) - 4)
      Debug.Print sFilename
      
      ClearPlayer 1
      LoadPlayer 1, sFilename
        If LCase(Trim$(Player(1).Name)) = LCase(Update) Then
            MsgBox "Conta: " & Player(1).Login
            Exit Sub
        End If
      
      sFilename = Dir()
    
    Loop
End Sub

Private Sub cmdIniciarEvent_Click()
    If EventGlobalType = 0 Then
        EventGlobalTick = GetTickCount
        StartGlobalEvent cmbEvent.ListIndex + 1
    End If
End Sub

Private Sub Command9_Click()
    SaveMaps
    SetStatus "Maps saved."
End Sub

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lstHouse_Click()
    If fraRenovar.Visible = True Then
        txtRenovarName.Text = frmServer.lstHouse.ListItems(frmServer.lstHouse.SelectedItem.Index).SubItems(1)
        txtRenovarDate.Text = frmServer.lstHouse.ListItems(frmServer.lstHouse.SelectedItem.Index).SubItems(2)
    End If
End Sub

Private Sub mnuGiveItem_Click()
    Dim Name As Long
    Name = frmServer.lvwInfo.SelectedItem.Index

    If IsPlaying(Name) Then
        Call AlertMSG(Name, "You have been kicked by the server owner!")
    End If
End Sub

Private Sub mnuKill_Click()
    Dim Name As Long
    Name = frmServer.lvwInfo.SelectedItem.Index

    If IsPlaying(Name) Then
        Call KillPlayer(Name)
    End If
End Sub

Private Sub mnuVIP_Click()
    Dim Name As Long
    Name = frmServer.lvwInfo.SelectedItem.Index

    If IsPlaying(Name) Then
        frmVIP.Show
        frmVIP.txtName = GetPlayerName(frmServer.lvwInfo.SelectedItem.Index)
        frmVIP.txtIndex = frmServer.lvwInfo.SelectedItem.Index
    End If
End Sub

Private Sub sckWeb_Close()
    sckWeb.Close
    lblWebStatus.Caption = "Desconectado"
    lblWebStatus.ForeColor = QBColor(brightred)
End Sub

Private Sub sckWeb_ConnectionRequest(ByVal requestID As Long)
    sckWeb.Close
    sckWeb.Accept requestID
    lblWebStatus.Caption = "Conectado"
    lblWebStatus.ForeColor = QBColor(brightgreen)
End Sub

Private Sub sckWeb_DataArrival(ByVal bytesTotal As Long)
    Dim packet As String
    sckWeb.GetData packet, vbString, bytesTotal
    HandleWebData packet
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
    Call AcceptConnection(Index, SocketId)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

' ********************

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadClasses_Click()
Dim i As Long
    Call LoadClasses
    Call TextAdd("All classes reloaded.", ChatSystem)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendClasses i
        End If
    Next
End Sub

Private Sub cmdReloadItems_Click()
Dim i As Long
    Call LoadItems
    Call TextAdd("All items reloaded.", ChatSystem)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendItems i
        End If
    Next
End Sub

Private Sub cmdReloadMaps_Click()
Dim i As Long
    Call LoadMaps
    Call TextAdd("All maps reloaded.", ChatSystem)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next
End Sub

Private Sub cmdReloadNPCs_Click()
Dim i As Long
    Call LoadNpcs
    Call LoadNPCBase
    Call TextAdd("All npcs reloaded.", ChatSystem)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendNpcs i
        End If
    Next
End Sub

Private Sub cmdReloadShops_Click()
Dim i As Long
    Call LoadShops
    Call TextAdd("All shops reloaded.", ChatSystem)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendShops i
        End If
    Next
End Sub

Private Sub cmdReloadSpells_Click()
Dim i As Long
    Call LoadSpells
    Call TextAdd("All spells reloaded.", ChatSystem)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next
End Sub

Private Sub cmdReloadResources_Click()
Dim i As Long
    Call LoadResources
    Call TextAdd("All Resources reloaded.", ChatSystem)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendResources i
        End If
    Next
End Sub

Private Sub cmdReloadAnimations_Click()
Dim i As Long
    Call LoadAnimations
    Call TextAdd("All Animations reloaded.", ChatSystem)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendAnimations i
        End If
    Next
End Sub
Private Sub cmdReloadEvents_Click()
Dim i As Long, n As Long
    Call LoadEvents
    Call TextAdd("All Events reloaded.", ChatSystem)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            For n = 1 To MAX_EVENTS
                Call Events_SendEventData(i, n)
                Call SendMapKey(i, Player(i).EventOpen(n), n)
            Next
        End If
    Next
End Sub
Private Sub cmdReloadEffects_Click()
Dim i As Long
    Call LoadEffects
    Call TextAdd("All Effects reloaded.", ChatSystem)
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendEffects i
        End If
    Next
End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", brightblue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
    End If
End Sub

Private Sub Form_Load()
    Call UsersOnline_Start
End Sub

Private Sub Form_Resize()

    If frmServer.WindowState = vbMinimized Then
        'frmServer.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.Index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Server: " & txtChat.Text, ChatGlobal)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub


Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As Long
    Name = frmServer.lvwInfo.SelectedItem.Index

    If IsPlaying(Name) Then
        Call AlertMSG(Name, "You have been kicked by the server owner!")
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As Long
    Name = frmServer.lvwInfo.SelectedItem.Index

    If IsPlaying(Name) Then
        CloseSocket (Name)
    End If

End Sub

Sub mnuBanPlayer_click()
    Dim Name As Long
    Name = frmServer.lvwInfo.SelectedItem.Index

    If IsPlaying(Name) Then
        Call ServerBanIndex(Name)
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As Long
    Name = frmServer.lvwInfo.SelectedItem.Index

    If IsPlaying(Name) Then
        Call SetPlayerAccess(Name, 4)
        Call SendPlayerData(Name)
        Call PlayerMsg(Name, "You have been granted administrator access.", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As Long
    Name = frmServer.lvwInfo.SelectedItem.Index

    If IsPlaying(Name) Then
        Call SetPlayerAccess(Name, 0)
        Call SendPlayerData(Name)
        Call PlayerMsg(Name, "You have had your administrator access revoked.", brightred)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lmsg As Long
    lmsg = X / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
    End Select

End Sub

Private Sub txtDrop_Change()
    If IsNumeric(txtDrop) Then
        Options.DropFactor = txtDrop
        SaveOptions
    End If
End Sub

Private Sub txtExpFactor_Change()
    If IsNumeric(txtExpFactor) Then
        Options.ExpFactor = txtExpFactor
        SaveOptions
    End If
End Sub

Private Sub txtGold_Change()
    If IsNumeric(txtGold) Then
        Options.GoldFactor = txtGold
        SaveOptions
    End If
End Sub

Private Sub txtMOTD_Change()
    Options.MOTD = Trim$(txtMOTD.Text)
    SaveOptions
End Sub

Private Sub txtResource_Change()
    If IsNumeric(txtResource) Then
        Options.ResourceFactor = txtResource
        SaveOptions
    End If
End Sub
