VERSION 5.00
Begin VB.Form frmExercito 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Exercito"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7950
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
   ScaleHeight     =   2415
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "Adicionar á produção"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Requisitos"
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4935
         Begin VB.Label Label1 
            Caption         =   "%Valor% Moedas Z"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   4695
         End
         Begin VB.Label lblSemente 
            Caption         =   "1 Semente de Saibaman nível %nivel%"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Label lblTime 
         Caption         =   "Tempo para conclusão: 0m"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Label lblLoc 
         Caption         =   "Locação:"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de produção"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ListBox lstProducao 
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmExercito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
