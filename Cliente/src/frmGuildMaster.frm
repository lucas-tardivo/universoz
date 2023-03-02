VERSION 5.00
Begin VB.Form frmGuildMaster 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   Caption         =   "Criar uma nova guild"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Concluído!"
      Height          =   375
      Left            =   4080
      TabIndex        =   46
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ícone"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF80&
         Caption         =   "Cores"
         Height          =   1695
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   3135
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   840
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   47
            Top             =   480
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   600
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   45
            Top             =   480
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   360
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   44
            Top             =   480
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   43
            Top             =   480
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   2760
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   42
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   2520
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   41
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   2280
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   40
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   2040
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   39
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   1800
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   38
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   1560
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   37
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   1320
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   36
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1080
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   35
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   840
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   34
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   600
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   33
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   360
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   32
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   31
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox picSelected 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1920
            ScaleHeight     =   345
            ScaleWidth      =   945
            TabIndex        =   30
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cor selecionada:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   2175
         End
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
         TabIndex        =   27
         Top             =   1680
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
         TabIndex        =   26
         Top             =   1680
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
         TabIndex        =   25
         Top             =   1680
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
         TabIndex        =   24
         Top             =   1680
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
         TabIndex        =   23
         Top             =   1680
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
         TabIndex        =   22
         Top             =   1320
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
         TabIndex        =   21
         Top             =   1320
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
         TabIndex        =   20
         Top             =   1320
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
         TabIndex        =   19
         Top             =   1320
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
         TabIndex        =   18
         Top             =   1320
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
         TabIndex        =   17
         Top             =   960
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
         TabIndex        =   16
         Top             =   960
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
         TabIndex        =   15
         Top             =   960
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
         TabIndex        =   14
         Top             =   960
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
         TabIndex        =   13
         Top             =   960
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
         TabIndex        =   12
         Top             =   600
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
         TabIndex        =   11
         Top             =   600
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
         TabIndex        =   10
         Top             =   600
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
         TabIndex        =   9
         Top             =   600
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
         TabIndex        =   8
         Top             =   600
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
         TabIndex        =   7
         Top             =   240
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
         TabIndex        =   6
         Top             =   240
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
         TabIndex        =   5
         Top             =   240
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
         TabIndex        =   4
         Top             =   240
         Width           =   375
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
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.TextBox txtGuildName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "É necessário ser level 20 e possuir 1,000,000 de Moedas Z para criar uma guild!"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5415
   End
End
Attribute VB_Name = "frmGuildMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IconColor(0 To 24) As Byte
Private SelectedColor As Byte

Private Sub Command1_Click()
    Dim GuildName As String
    GuildName = Trim(txtGuildName.Text)
    If Len(GuildName) > 3 Then
        If Len(GuildName) < NAME_LENGTH Then
            SendCreateGuild GuildName, IconColor
            Unload Me
        Else
            MsgBox "O nome da sua guild não pode exceder " & NAME_LENGTH & " caracteres"
        End If
    Else
        MsgBox "O nome de sua guild deve ter no mínimo 3 caracteres"
    End If
End Sub

Private Sub Form_Load()
    For i = 0 To 15
        picColor(i).BackColor = QBColor(i)
    Next i
    DrawIcon
End Sub

Private Sub DrawIcon()
    Dim i As Long
    For i = 0 To 24
        picIcon(i).BackColor = QBColor(IconColor(i))
    Next i
End Sub

Private Sub picColor_Click(Index As Integer)
    SelectedColor = Index
    picSelected.BackColor = QBColor(SelectedColor)
End Sub

Private Sub picIcon_Click(Index As Integer)
    IconColor(Index) = SelectedColor
    DrawIcon
End Sub
