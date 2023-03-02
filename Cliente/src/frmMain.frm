VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picGuildAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   3000
      ScaleHeight     =   4905
      ScaleWidth      =   5985
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtMOTD 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CommandButton cmdMOTD 
         Caption         =   "Editar Notícia"
         Height          =   255
         Left            =   3120
         TabIndex        =   71
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame frmInfo 
         BackColor       =   &H00FFFF80&
         Caption         =   "Informação"
         Height          =   1815
         Left            =   240
         TabIndex        =   65
         Top             =   2040
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CommandButton Command5 
            Caption         =   "Fechar"
            Height          =   255
            Left            =   360
            TabIndex        =   70
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label lblDonations 
            BackStyle       =   0  'Transparent
            Caption         =   "Doações:"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   960
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label lblGuildExp 
            BackStyle       =   0  'Transparent
            Caption         =   "Guild Exp:"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label lblLevel 
            BackStyle       =   0  'Transparent
            Caption         =   "Nível:"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label lblPlayerName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "%Rank% %Name%"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame frmInvite 
         BackColor       =   &H00FFFF80&
         Caption         =   "Convidar para a guild"
         Height          =   1815
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   2895
         Begin VB.CommandButton Command4 
            Caption         =   "Cancelar"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   2655
         End
         Begin VB.CommandButton cmdSendInvite 
            Caption         =   "Convidar"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtPlayerName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame frmMembers 
         BackColor       =   &H00FFFF80&
         Caption         =   "Membros (Info: Duplo clique)"
         Height          =   1815
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   2895
         Begin VB.ListBox lstMembers 
            Height          =   1425
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "X"
         Height          =   255
         Left            =   5640
         TabIndex        =   64
         Top             =   120
         Width           =   255
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF80&
         Caption         =   "Missão"
         Height          =   1695
         Left            =   120
         TabIndex        =   60
         Top             =   5400
         Visible         =   0   'False
         Width           =   5775
         Begin VB.PictureBox picMission 
            BackColor       =   &H008080FF&
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1275
            ScaleWidth      =   5475
            TabIndex        =   61
            Top             =   240
            Width           =   5535
            Begin VB.CommandButton Command2 
               Caption         =   "Comprar missão"
               Height          =   255
               Left            =   1320
               TabIndex        =   63
               Top             =   960
               Width           =   2895
            End
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   $"frmMain.frx":0CCA
               Height          =   855
               Left            =   120
               TabIndex        =   62
               Top             =   120
               Width           =   5175
            End
         End
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   55
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   480
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   54
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   840
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   53
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   52
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   51
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   50
         Top             =   480
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   480
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   49
         Top             =   480
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   840
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   48
         Top             =   480
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   47
         Top             =   480
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   46
         Top             =   480
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   45
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   480
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   44
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   840
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   43
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   13
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   42
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   14
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   41
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   15
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   40
         Top             =   1200
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   16
         Left            =   480
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   39
         Top             =   1200
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   17
         Left            =   840
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   38
         Top             =   1200
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   18
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   37
         Top             =   1200
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   19
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   36
         Top             =   1200
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   20
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   35
         Top             =   1560
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   21
         Left            =   480
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   34
         Top             =   1560
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   22
         Left            =   840
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   33
         Top             =   1560
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   23
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   32
         Top             =   1560
         Width           =   375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   24
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   31
         Top             =   1560
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Banco da guild"
         Height          =   855
         Left            =   120
         TabIndex        =   29
         Top             =   3960
         Width           =   5775
         Begin VB.Frame frmDoar 
            BackColor       =   &H00FFFF80&
            Caption         =   "Doar"
            Height          =   615
            Left            =   3240
            TabIndex        =   73
            Top             =   120
            Width           =   2415
            Begin VB.TextBox txtQuant 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   81
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txtQuant 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   80
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txtQuant 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   79
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txtQuant 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   3
               Left            =   120
               TabIndex        =   78
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton cmdDonate 
               Caption         =   "Doar"
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   77
               Top             =   960
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.CommandButton cmdDonate 
               Caption         =   "Doar"
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   76
               Top             =   960
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.CommandButton cmdDonate 
               Caption         =   "Doar"
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   75
               Top             =   960
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.CommandButton cmdDonate 
               Caption         =   "Doar"
               Height          =   255
               Index           =   3
               Left            =   1680
               TabIndex        =   74
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Doar"
            Height          =   255
            Left            =   4800
            TabIndex        =   82
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblGold 
            BackStyle       =   0  'Transparent
            Caption         =   "Moedas Z:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label lblYellow 
            BackStyle       =   0  'Transparent
            Caption         =   "Especiaria amarela:"
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Visible         =   0   'False
            Width           =   5535
         End
         Begin VB.Label lblBlue 
            BackStyle       =   0  'Transparent
            Caption         =   "Especiaria azul:"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   960
            Visible         =   0   'False
            Width           =   5535
         End
         Begin VB.Label lblRed 
            BackStyle       =   0  'Transparent
            Caption         =   "Especiaria vermelha:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   960
            Visible         =   0   'False
            Width           =   5535
         End
      End
      Begin VB.Frame frmBasicActions 
         BackColor       =   &H00FFFF80&
         Caption         =   "Ações"
         Height          =   1815
         Left            =   3120
         TabIndex        =   10
         Top             =   2040
         Width           =   2775
         Begin VB.Frame frmActions 
            BackColor       =   &H00FFFF80&
            Caption         =   "Mestre"
            Height          =   1095
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2535
            Begin VB.CommandButton Command6 
               Caption         =   "?"
               Height          =   255
               Left            =   2160
               TabIndex        =   84
               Top             =   720
               Width           =   255
            End
            Begin VB.CommandButton cmdPromote 
               Caption         =   "Promover"
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   26
               Top             =   480
               Width           =   1095
            End
            Begin VB.CommandButton cmdRevoke 
               Caption         =   "Rebaixar"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   27
               Top             =   480
               Width           =   1215
            End
            Begin VB.CommandButton cmdKick 
               Caption         =   "Expulsar"
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   25
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton cmdInvite 
               Caption         =   "Convidar"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   1215
            End
            Begin VB.CheckBox chkBlock 
               BackColor       =   &H00FFFF80&
               Caption         =   "Bloquear UP"
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   720
               Width           =   2295
            End
         End
         Begin VB.Frame frmActions 
            BackColor       =   &H00FFFF80&
            Caption         =   "Major"
            Height          =   1095
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2535
            Begin VB.CommandButton cmdKick 
               Caption         =   "Expulsar membro"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   720
               Width           =   2295
            End
            Begin VB.CommandButton cmdPromote 
               Caption         =   "Promover"
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   22
               Top             =   480
               Width           =   1095
            End
            Begin VB.CommandButton cmdRevoke 
               Caption         =   "Rebaixar"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   480
               Width           =   1215
            End
            Begin VB.CommandButton cmdInvite 
               Caption         =   "Convidar para a guild"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   2295
            End
         End
         Begin VB.Frame frmActions 
            BackColor       =   &H00FFFF80&
            Caption         =   "Capitão"
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2535
            Begin VB.CommandButton cmdInvite 
               Caption         =   "Convidar para a guild"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   2295
            End
         End
         Begin VB.CommandButton cmdLeave 
            Caption         =   "Sair da guild"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   2535
         End
      End
      Begin VB.PictureBox picGuildExpBarMold 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         ScaleHeight     =   225
         ScaleWidth      =   3825
         TabIndex        =   4
         Top             =   1680
         Width           =   3855
         Begin VB.PictureBox picGuildExpBar 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   495
            TabIndex        =   7
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.Label lblGuildMOTD 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "%MOTD%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   30
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblGuildMembers 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Membros: %Actual%/%Limit%"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label lblGuildLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Guild Level: %Level%"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label lblGuildName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "%GuildName%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.PictureBox picTroca 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   2760
      ScaleHeight     =   3465
      ScaleWidth      =   6585
      TabIndex        =   130
      Top             =   2280
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   4440
         TabIndex        =   155
         Top             =   2160
         Width           =   2055
         Begin VB.TextBox txtEspQuant 
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   157
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdSell 
            Caption         =   "Vender"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   156
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Quant.:"
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblReceber 
            BackStyle       =   0  'Transparent
            Caption         =   "Á receber: 0z"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   158
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   2280
         TabIndex        =   150
         Top             =   2160
         Width           =   2055
         Begin VB.TextBox txtEspQuant 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   152
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdSell 
            Caption         =   "Vender"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   151
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Quant.:"
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblReceber 
            BackStyle       =   0  'Transparent
            Caption         =   "Á receber: 0z"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   153
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   120
         TabIndex        =   145
         Top             =   2160
         Width           =   2055
         Begin VB.CommandButton cmdSell 
            Caption         =   "Vender"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   149
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox txtEspQuant 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   147
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblReceber 
            BackStyle       =   0  'Transparent
            Caption         =   "Á receber: 0z"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   148
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Quant.:"
            Height          =   255
            Left            =   120
            TabIndex        =   146
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   520
         Left            =   5280
         Picture         =   "frmMain.frx":0D78
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   138
         Top             =   480
         Width           =   520
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   520
         Left            =   3000
         Picture         =   "frmMain.frx":79BC
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   137
         Top             =   480
         Width           =   520
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   520
         Left            =   960
         Picture         =   "frmMain.frx":E600
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   136
         Top             =   480
         Width           =   520
      End
      Begin VB.Label lblAlta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Valor: 0%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   162
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblAlta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Valor: 0%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   161
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblAlta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Valor: 0%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   160
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblEspPrice 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Preço: 0z"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   144
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblEspPrice 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Preço: 0z"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   143
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblEspPrice 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Preço: 0z"
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
         Index           =   0
         Left            =   120
         TabIndex        =   142
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblTotalAcumulado 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Acumulado: %Valor%"
         ForeColor       =   &H00008080&
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   141
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblTotalAcumulado 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Acumulado: %Valor%"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   140
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblTotalAcumulado 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Acumulado: %Valor%"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   139
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Especiaria Amarela"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   4440
         TabIndex        =   135
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Especiaria Azul"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2280
         TabIndex        =   134
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Especiaria Vermelha"
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
         Left            =   120
         TabIndex        =   133
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
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
         Left            =   6120
         TabIndex        =   132
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda de especiarias"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   131
         Top             =   120
         Width           =   2115
      End
   End
   Begin VB.PictureBox picExercito 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   3720
      ScaleHeight     =   4665
      ScaleWidth      =   4665
      TabIndex        =   116
      Top             =   1800
      Visible         =   0   'False
      Width           =   4695
      Begin VB.ListBox lstEFila 
         Height          =   1815
         Left            =   120
         TabIndex        =   124
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "X"
         Height          =   255
         Left            =   4320
         TabIndex        =   117
         Top             =   120
         Width           =   255
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00008000&
         Caption         =   "Produzir"
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   120
         TabIndex        =   118
         Top             =   2280
         Width           =   4455
         Begin VB.CommandButton Command11 
            Caption         =   "Produzir"
            Height          =   255
            Left            =   840
            TabIndex        =   120
            Top             =   1920
            Width           =   2895
         End
         Begin VB.TextBox txtEQuant 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            TabIndex        =   119
            Text            =   "1"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblAlloc 
            BackStyle       =   0  'Transparent
            Caption         =   "Residencias: 0/70"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lblESaibaman 
            BackStyle       =   0  'Transparent
            Caption         =   "1 Semente de saibaman nivel X (Contém: X)"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   1080
            Width           =   4215
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "- Requisitos -"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   840
            Width           =   4215
         End
         Begin VB.Label lblEGold 
            BackStyle       =   0  'Transparent
            Caption         =   "Moedas:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   1320
            Width           =   4215
         End
         Begin VB.Label lblETime 
            BackStyle       =   0  'Transparent
            Caption         =   "Tempo:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   1560
            Width           =   4215
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exército"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.PictureBox picFabrica 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   3720
      ScaleHeight     =   4665
      ScaleWidth      =   4665
      TabIndex        =   103
      Top             =   1800
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command9 
         Caption         =   "X"
         Height          =   255
         Left            =   4320
         TabIndex        =   113
         Top             =   120
         Width           =   255
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFF80&
         Caption         =   "Produzir"
         Height          =   2295
         Left            =   120
         TabIndex        =   106
         Top             =   2280
         Width           =   4455
         Begin VB.TextBox txtSQuant 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            TabIndex        =   115
            Text            =   "1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Produzir"
            Height          =   255
            Left            =   840
            TabIndex        =   111
            Top             =   1920
            Width           =   2895
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade:"
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblSemente 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label lblTime 
            BackStyle       =   0  'Transparent
            Caption         =   "Tempo:"
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   1560
            Width           =   4215
         End
         Begin VB.Label lblSYellow 
            BackStyle       =   0  'Transparent
            Caption         =   "Esp. Amarela:"
            ForeColor       =   &H0000C0C0&
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   1320
            Width           =   4215
         End
         Begin VB.Label lblSBlue 
            BackStyle       =   0  'Transparent
            Caption         =   "Esp. Azul:"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   1080
            Width           =   4215
         End
         Begin VB.Label lblSRed 
            BackStyle       =   0  'Transparent
            Caption         =   "Esp. Vermelha:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   840
            Width           =   4215
         End
      End
      Begin VB.ListBox lstFila 
         Height          =   1815
         Left            =   120
         TabIndex        =   105
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fábrica de sementes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.PictureBox picArena 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   3240
      ScaleHeight     =   4425
      ScaleWidth      =   5505
      TabIndex        =   85
      Top             =   1560
      Visible         =   0   'False
      Width           =   5535
      Begin VB.ComboBox cmbType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMain.frx":15244
         Left            =   1560
         List            =   "frmMain.frx":1524B
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Desafiar!"
         Height          =   255
         Left            =   1800
         TabIndex        =   100
         Top             =   3960
         Width           =   2175
      End
      Begin VB.TextBox txtMoedas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   98
         Text            =   "0"
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Jogadores"
         Height          =   1455
         Left            =   120
         TabIndex        =   89
         Top             =   1080
         Width           =   5295
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2760
            TabIndex        =   95
            Text            =   "Nome"
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   94
            Text            =   "Nome"
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2760
            TabIndex        =   93
            Text            =   "Nome"
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   92
            Text            =   "Nome"
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   91
            Text            =   "Nome"
            Top             =   240
            Width           =   2415
         End
         Begin VB.Line Line1 
            X1              =   2640
            X2              =   2640
            Y1              =   240
            Y2              =   1320
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Você"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.ComboBox cmbPlayers 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMain.frx":1525A
         Left            =   2160
         List            =   "frmMain.frx":15267
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de desafio:"
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total para cada jogador: 10000z"
         Height          =   255
         Left            =   240
         TabIndex        =   99
         Top             =   3240
         Width           =   5055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Aposta em moedas:"
         Height          =   255
         Left            =   240
         TabIndex        =   97
         Top             =   2880
         Width           =   5055
      End
      Begin VB.Label lblPrice 
         BackStyle       =   0  'Transparent
         Caption         =   "Taxa individual para o uso da arena: 10.000z"
         Height          =   255
         Left            =   240
         TabIndex        =   96
         Top             =   2640
         Width           =   4935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de jogadores:"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   720
         Width           =   5295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Arena"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.PictureBox picLoad 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
      Begin VB.Label lblLoad 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carregando..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   135
      End
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long

Private Sub chkBlock_Click()
    SendGuildUpBlock chkBlock.Value
End Sub

Private Sub cmbPlayers_Change()
    cmbPlayers_Click
End Sub

Private Sub cmbPlayers_Click()
    Dim i As Long

    For i = 1 To 4
        If i <= cmbPlayers.ListIndex * 2 Then
            txtName(i).Enabled = True
        Else
            txtName(i).Enabled = False
        End If
    Next i
End Sub

Private Sub cmdDonate_Click(Index As Integer)
    If IsNumeric(txtQuant(Index).Text) Then
        If Val(txtQuant(Index).Text) > 0 Then
            SendGuildDonate Index, Val(txtQuant(Index).Text)
            txtQuant(Index).Text = vbNullString
        End If
    End If
End Sub

Private Sub cmdInvite_Click(Index As Integer)
    frmInvite.visible = True
End Sub

Private Sub cmdKick_Click(Index As Integer)
    If lstMembers.ListIndex >= 0 Then
        Dim i As Long
        i = MsgBox("Tem certeza que deseja expulsar " & Trim$(Guild(Player(MyIndex).Guild).Member(lstMembers.ListIndex + 1).Name) & " da guild?", vbYesNo)
        If i = vbYes Then
            SendGuildKick lstMembers.ListIndex + 1
        End If
    End If
End Sub

Private Sub cmdLeave_Click()
    Dim i As Long
    i = MsgBox("Tem certeza que deseja sair da guild?", vbYesNo)
    If i = vbYes Then
        SendGuildLeave
    End If
End Sub

Private Sub cmdMOTD_Click()
    If txtMOTD.visible = False Then
        txtMOTD.visible = True
    Else
        txtMOTD.visible = False
        lblGuildMOTD.Caption = txtMOTD.Text
        SendGuildMOTD txtMOTD.Text
    End If
End Sub

Private Sub cmdPromote_Click(Index As Integer)
    If lstMembers.ListIndex >= 0 Then
        If Player(MyIndex).Guild > 0 Then
            If Guild(Player(MyIndex).Guild).Member(lstMembers.ListIndex + 1).Rank = 2 Then
                Dim i As Long
                i = MsgBox("Ao promover este membro você estará nomeando-o como mestre da sua guild e você será rebaixado para major, tem certeza disso?", vbYesNo)
                If i = vbYes Then
                    SendGuildPromote lstMembers.ListIndex + 1
                End If
            Else
                SendGuildPromote lstMembers.ListIndex + 1
            End If
        End If
    End If
End Sub

Private Sub cmdRevoke_Click(Index As Integer)
    If lstMembers.ListIndex >= 0 Then
        SendGuildRevoke lstMembers.ListIndex + 1
    End If
End Sub

Private Sub cmdSell_Click(Index As Integer)
    If Val(txtEspQuant(Index).Text) > 0 Then
        SendSellEspeciaria Index + 1, Val(txtEspQuant(Index))
        txtQuant(Index).Text = 0
    End If
End Sub

Private Sub cmdSendInvite_Click()
    If Len(txtPlayerName.Text) > 3 Then
        SendGuildInvite txtPlayerName.Text
        frmInvite.visible = False
    End If
End Sub

Private Sub Command1_Click()
frmDoar.visible = True
End Sub

Private Sub Command10_Click()
    picExercito.visible = False
End Sub

Private Sub Command11_Click()
    If lstEFila.ListIndex >= 0 Then
        If Val(txtEQuant) > 0 And Val(txtEQuant) <= 255 Then
            SendProduceSementes lstEFila.ListIndex + 1, Val(txtEQuant)
        End If
    End If
End Sub

Private Sub Command3_Click()
    picGuildAdmin.visible = False
End Sub

Private Sub Command4_Click()
    frmInvite.visible = False
End Sub

Private Sub Command5_Click()
    frmInfo.visible = False
End Sub

Private Sub Command6_Click()
    MsgBox "Cada vez que um membro da sua guild derrota um planeta, é cobrada uma taxa do banco da sua guild em especiarias e ouro baseada no nível do planeta capturado pelo membro para que a guild possa receber experiência, caso você não queira gastar fundos da guild com evolução você pode bloquear o up nesta opção.", vbInformation
End Sub

Private Sub Command7_Click()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CChallengeArena
    buffer.WriteByte 1
    buffer.WriteByte cmbType.ListIndex
    buffer.WriteByte cmbPlayers.ListIndex
    Dim i As Long
    For i = 0 To 4
        buffer.WriteString txtName(i)
    Next i
    buffer.WriteLong Val(txtMoedas.Text)
    
    SendData buffer.ToArray()
    Set buffer = Nothing
    picArena.visible = False
End Sub

Private Sub Command8_Click()
    If lstFila.ListIndex >= 0 Then
        If Val(txtSQuant) > 0 And Val(txtSQuant) <= 255 Then
            SendProduceSementes lstFila.ListIndex + 1, Val(txtSQuant)
        End If
    End If
End Sub

Private Sub Command9_Click()
    picFabrica.visible = False
End Sub

Private Sub Form_DblClick()
    HandleDoubleClick
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseDown Button
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseUp Button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Cancel = True
    RemoveAllMapSounds
    StopAllSounds
    logoutGame
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    HandleMouseMove CLng(X), CLng(Y), Button
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub Label13_Click()
    picTroca.visible = False
End Sub

Private Sub lstEFila_Click()
    UpdateSaibamansPrice
End Sub

Private Sub UpdateSaibamansPrice()
    If lstEFila.ListIndex >= 0 Then
        lblESaibaman.Caption = Val(txtEQuant) & " Semente de saibaman nível " & lstEFila.ListIndex + 1 & " (Contém: " & Sementes(lstEFila.ListIndex + 1) & ")"
        lblEGold.Caption = "Moedas Z:" & (Val(txtEQuant) * (Fat(lstEFila.ListIndex + 3) * 100)) & "z"
        lblETime.Caption = "Tempo:" & Fat(lstEFila.ListIndex + 1) & "m"
    End If
End Sub

Private Sub lstFila_Click()
    UpdateSementesPrice
End Sub

Private Sub UpdateSementesPrice()
    lblSemente.Caption = "Semente nível " & lstFila.ListIndex + 1
    lblSRed.Caption = "Esp. Vermelha: " & (Fat(lstFila.ListIndex + 1) * 25) * Val(txtSQuant)
    lblSBlue.Caption = "Esp. Azul: " & (Fat(lstFila.ListIndex + 1) * 15) * Val(txtSQuant)
    lblSYellow.Caption = "Esp. Amarela: " & (Fat(lstFila.ListIndex + 1) * 5) * Val(txtSQuant)
    lblTime.Caption = "Tempo por unidade:" & Fat(lstFila.ListIndex + 1) & "m"
End Sub

Private Function Fat(ByVal Number As Long) As Long
    If Number <= 1 Then
        Fat = 1
    Else
        Fat = Number * Fat(Number - 1)
    End If
End Function

Private Sub lstMembers_DblClick()
    Dim MemberIndex As Long
    Dim GuildNum As Long
    Dim MemberData As GuildMemberRec
    GuildNum = Player(MyIndex).Guild
    MemberIndex = lstMembers.ListIndex + 1
    MemberData = Guild(GuildNum).Member(MemberIndex)
    
    frmInfo.visible = True
    lblPlayerName.Caption = RankName(MemberData.Rank) & " " & Trim$(MemberData.Name)
    lblLevel.Caption = "Nível: " & MemberData.Level
    lblGuildExp.Caption = "Guild Exp: " & MemberData.GuildExp
    lblDonations.Caption = "Doações: " & MemberData.Donations
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    HandleKeyUp KeyCode

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEQuant_Change()
    If Not IsNumeric(txtEQuant) Or Val(txtEQuant) <= 0 Then txtEQuant = 1
    If Val(txtEQuant) > 255 Then txtEQuant = 255
    UpdateSaibamansPrice
End Sub

Private Sub txtEspQuant_Change(Index As Integer)
    If Not IsNumeric(txtEspQuant(Index)) Then
        txtEspQuant(Index) = 0
        Exit Sub
    End If
    txtEspQuant(Index) = Int(txtEspQuant(Index))
    lblReceber(Index).Caption = "Á receber: " & (Val(Int(txtEspQuant(Index))) * EspPrice(Index + 1)) & "z"
End Sub

Private Sub txtMoedas_Change()
    If Not IsNumeric(txtMoedas) Then txtMoedas = 0
    If Val(txtMoedas) < 0 Or Val(txtMoedas) > MAX_LONG Then txtMoedas = 0
    lblTotal = "Valor total para cada jogador: " & (Val(txtMoedas) + 10000) & "z"
End Sub

Private Sub txtSQuant_Change()
    If Not IsNumeric(txtSQuant) Or Val(txtSQuant) <= 0 Then txtSQuant = 1
    If Val(txtSQuant) > 255 Then txtSQuant = 255
    UpdateSementesPrice
End Sub
