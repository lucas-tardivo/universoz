VERSION 5.00
Begin VB.Form frmEditor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7455
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmLoad 
      Height          =   255
      Left            =   120
      TabIndex        =   71
      Top             =   480
      Width           =   255
      Begin VB.Label lblLoad 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   3120
         Width           =   6495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "e clique em ""Carregar"""
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
         TabIndex        =   73
         Top             =   3240
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Digite o nome da conta"
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
         TabIndex        =   72
         Top             =   3000
         Visible         =   0   'False
         Width           =   6495
      End
   End
   Begin VB.ComboBox txtFile 
      Height          =   315
      ItemData        =   "frmEditor.frx":2982
      Left            =   1560
      List            =   "frmEditor.frx":2984
      TabIndex        =   14
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados"
      Height          =   6495
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6735
      Begin VB.Frame Frame10 
         Caption         =   "Esotérica"
         Height          =   1455
         Left            =   120
         TabIndex        =   62
         Top             =   4920
         Width           =   3255
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   20
            Left            =   1680
            TabIndex        =   67
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   19
            Left            =   1440
            TabIndex        =   66
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblInfo 
            Caption         =   "min"
            Height          =   255
            Index           =   23
            Left            =   2760
            TabIndex        =   69
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblInfo 
            Caption         =   "%"
            Height          =   255
            Index           =   22
            Left            =   2760
            TabIndex        =   68
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblInfo 
            Caption         =   "Tempo restante:"
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   65
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblInfo 
            Caption         =   "Bonus de exp:"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   64
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblEsoterica 
            Caption         =   "Nome do item:"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Pontos"
         Height          =   3255
         Left            =   3480
         TabIndex        =   47
         Top             =   3120
         Width           =   3135
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   12
            Left            =   600
            TabIndex        =   60
            Top             =   2760
            Width           =   2415
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   18
            Left            =   840
            TabIndex        =   59
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   17
            Left            =   960
            TabIndex        =   57
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   16
            Left            =   1080
            TabIndex        =   56
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   15
            Left            =   480
            TabIndex        =   55
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   14
            Left            =   1320
            TabIndex        =   54
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   13
            Left            =   720
            TabIndex        =   53
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label lblInfo 
            Caption         =   "PDL:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   61
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "Pontos:"
            ForeColor       =   &H0000C0C0&
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "Constituição:"
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   52
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblInfo 
            Caption         =   "Técnica:"
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   51
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "KI:"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   50
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "Destreza:"
            ForeColor       =   &H000040C0&
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   49
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "Força:"
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "VIP"
         Height          =   1575
         Left            =   4440
         TabIndex        =   42
         Top             =   1560
         Width           =   2175
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   21
            Left            =   720
            TabIndex        =   77
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Level:"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblRestantes 
            Caption         =   "Dias restantes:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblInicio 
            Caption         =   "Início:"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblDias 
            Caption         =   "Dias:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lblVIP 
            Caption         =   "Jogador VIP:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Status"
         Height          =   1695
         Left            =   120
         TabIndex        =   32
         Top             =   3120
         Width           =   3255
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   11
            Left            =   480
            TabIndex        =   40
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   10
            Left            =   480
            TabIndex        =   39
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   9
            Left            =   600
            TabIndex        =   38
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   8
            Left            =   720
            TabIndex        =   37
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblExp 
            Caption         =   "/ 1000"
            Height          =   255
            Left            =   1920
            TabIndex        =   41
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblInfo 
            Caption         =   "MP:"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "HP:"
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "Exp:"
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblInfo 
            Caption         =   "Level:"
            ForeColor       =   &H0000C0C0&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Aparência"
         Height          =   1335
         Left            =   2280
         TabIndex        =   25
         Top             =   1680
         Width           =   2055
         Begin VB.TextBox txtHair 
            Height          =   285
            Left            =   1560
            TabIndex        =   78
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   6
            Left            =   840
            TabIndex        =   29
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "Sprite:"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblRaça 
            Caption         =   "Raça:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblHair 
            Caption         =   "Tipo de cabelo:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Informações"
         Height          =   1335
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   2055
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   7
            Left            =   840
            TabIndex        =   31
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "Acesso:"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblDir 
            Caption         =   "Direção:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblPK 
            Caption         =   "PK:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados pessoais"
         Height          =   975
         Left            =   1680
         TabIndex        =   18
         Top             =   600
         Width           =   2415
         Begin VB.TextBox txtData 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   840
            TabIndex        =   75
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   5
            Left            =   840
            TabIndex        =   21
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblInfo 
            Caption         =   "Senha:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "Login:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Legenda"
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1455
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "  Normal"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   120
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "  Modificado"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   120
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   4455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Localização"
         Height          =   975
         Left            =   4200
         TabIndex        =   5
         Top             =   600
         Width           =   2415
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   3
            Left            =   1440
            TabIndex        =   13
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   12
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Y:"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   8
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "X:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "Mapa:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do personagem:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Carregar"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Universo Z Online, GoPlay Games® 2014"
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
      TabIndex        =   70
      Top             =   7080
      Width           =   6735
   End
   Begin VB.Label lblInfo 
      Caption         =   ".bin"
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblInfo 
      Caption         =   "Nome da conta:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuChar 
      Caption         =   "Personagem"
      Visible         =   0   'False
      Begin VB.Menu mnuInv 
         Caption         =   "Itens"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSpells 
         Caption         =   "Magias"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuVIP 
         Caption         =   "VIP"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHotbar 
         Caption         =   "Hotbar"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSwitch 
         Caption         =   "Switches e Variaveis"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEvents 
         Caption         =   "Eventos"
         Shortcut        =   ^E
      End
      Begin VB.Menu line 
         Caption         =   "________________"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Fechar edição"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuEnd 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const Edited = &HC0FFC0

Sub LoadEditor()
    On Error Resume Next
    txtData(Editor.Name).Text = Trim$(Player.Name)
    txtData(Editor.Map).Text = Player.Map
    txtData(Editor.X).Text = Player.X
    txtData(Editor.Y).Text = Player.Y
    txtData(Editor.Login).Text = Trim$(Player.Login)
    txtData(Editor.Senha) = Trim$(Player.Password)
    txtData(Editor.Sprite) = Player.Sprite
    txtData(Editor.Acesso) = Player.Access
    txtData(Editor.Level) = Player.Level
    txtData(Editor.Experiencia) = Player.Exp
    txtData(Editor.HP) = Player.Vital(Vitals.HP)
    txtData(Editor.MP) = Player.Vital(Vitals.MP)
    txtData(Editor.PDL) = Player.PDL
    txtData(Editor.Pontos) = Player.Points
    txtData(Editor.Bonus) = Player.EsoBonus
    txtData(Editor.Tempo) = Player.EsoTime
    txtData(Editor.LevelVIP) = Player.VIP
    txtHair.Text = Player.Hair
    
    Dim i As Long
    For i = 13 To 17
        txtData(i) = Player.stat(i - 12)
    Next i
    
    'lblRaça.Caption = "Raça: " & Trim$(Class(Player.Class).Name)
    lblPK.Caption = "PK: " & Player.PK
    
    Dim PlayerDir As String
    
    Select Case Player.Dir
        Case 0: PlayerDir = "Cima"
        Case 1: PlayerDir = "Baixo"
        Case 2: PlayerDir = "Esquerda"
        Case 3: PlayerDir = "Direita"
    End Select
    
    lblDir.Caption = "Direção: " & PlayerDir
    
    PlayerDir = "Nenhum"
    If Player.VIP = 1 Then PlayerDir = "Comum"
    
    lblVIP.Caption = "Plano VIP:" & PlayerDir
    lblDias.Caption = "Dias: " & Player.VIPDias
    lblInicio.Caption = "Inicio: " & Player.VIPData
    lblEsoterica.Caption = "Nome do item: " & Trim$(Item(Player.EsoNum).Name)
    
    If Player.VIP = 1 Then lblRestantes.Caption = "Dias restantes: " & (Player.VIPDias - DateDiff("d", Player.VIPData, Date))
    
    lblExp.Caption = "/ " & Val(GetVar(App.Path & "\data\exp.ini", "EXPERIENCE", "Exp" & Player.Level))
    
    For i = 0 To Editor.TotalTexts - 1
        txtData(i).BackColor = &H80000005
    Next i
    
    txtData(Editor.Login).BackColor = &HE0E0E0
End Sub

Private Sub Form_Load()
    DoEvents
    frmLoad.Height = 6495
    frmLoad.Width = 6735
    frmEditor.Visible = True
    frmEditor.Height = 8190
    'frmEditor.Height = 1230
    frmEditor.Enabled = False
    LoadList

    frmEditor.Caption = "Carregando Itens..."
    lblLoad.Caption = frmEditor.Caption
    LoadItems
    frmEditor.Caption = "Carregando Magias..."
    lblLoad.Caption = frmEditor.Caption
    LoadSpells
    frmEditor.Caption = "Carregando Switches..."
    lblLoad.Caption = frmEditor.Caption
    LoadSwitches
    frmEditor.Caption = "Carregando Variaveis..."
    lblLoad.Caption = frmEditor.Caption
    LoadVariables
    frmEditor.Caption = "Carregando Eventos..."
    lblLoad.Caption = frmEditor.Caption
    LoadEvents
    
    Label5.Visible = True
    Label6.Visible = True
    lblLoad.Visible = False
    frmEditor.Enabled = True
    frmEditor.Caption = "(UNIVERSO Z) Editor de contas"
End Sub

Private Sub Command1_Click()
    If Dir(App.Path & "\data\accounts\" & Trim(txtFile) & ".bin", vbArchive) <> "" Then
        LoadPlayer txtFile
        LoadEditor
        AddToList txtFile
        Frame1.Caption = "Conta " & txtFile
        Command2.Enabled = True
        Command1.Enabled = False
        frmEditor.Height = 8190
        txtFile.Enabled = False
        mnuChar.Visible = True
        mnuFile.Visible = False
        frmLoad.Visible = False
    Else
        MsgBox "Conta não encontrada no banco de dados!", vbCritical
    End If
End Sub

Private Sub Command2_Click()
    SavePlayer txtFile
    Command2.Enabled = False
    Command1.Enabled = True
    txtFile.Enabled = True
    'frmEditor.Height = 1230
    mnuChar.Visible = False
    mnuFile.Visible = True
    frmLoad.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuClose_Click()
    Command1.Enabled = True
    Command2.Enabled = False
    txtFile.Enabled = True
    'frmEditor.Height = 1230
    mnuChar.Visible = False
    mnuFile.Visible = True
    frmLoad.Visible = True
End Sub

Private Sub mnuEnd_Click()
    End
End Sub

Private Sub mnuEvents_Click()
    frmEvents.Visible = True
End Sub

Private Sub mnuHotbar_Click()
    frmHotbar.Visible = True
End Sub

Private Sub mnuInv_Click()
    frmInv.Visible = True
End Sub

Private Sub mnuSpells_Click()
    frmSpells.Visible = True
End Sub

Private Sub mnuSwitch_Click()
    frmSwitches.Visible = True
End Sub

Private Sub mnuVIP_Click()
    frmVIP.Visible = True
End Sub

Private Sub txtData_Change(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case Editor.Name: Player.Name = txtData(Index).Text
        Case Editor.Map: Player.Map = txtData(Index).Text
        Case Editor.X: Player.X = txtData(Index).Text
        Case Editor.Y: Player.Y = txtData(Index).Text
        Case Editor.Login: Player.Login = txtData(Index)
        Case Editor.Senha: Player.Password = txtData(Index)
        Case Editor.Sprite: Player.Sprite = txtData(Index)
        Case Editor.Acesso: Player.Access = txtData(Index)
        Case Editor.Level: Player.Level = txtData(Index)
        Case Editor.Experiencia: Player.Exp = txtData(Index)
        Case Editor.HP: Player.Vital(Vitals.HP) = txtData(Index)
        Case Editor.MP: Player.Vital(Vitals.MP) = txtData(Index)
        Case Editor.PDL: Player.PDL = txtData(Index)
        Case Editor.Força: Player.stat(Stats.Strength) = txtData(Index)
        Case Editor.Constituiçao: Player.stat(Stats.Endurance) = txtData(Index)
        Case Editor.KI: Player.stat(Stats.Intelligence) = txtData(Index)
        Case Editor.Destreza: Player.stat(Stats.agility) = txtData(Index)
        Case Editor.Tecnica: Player.stat(Stats.Willpower) = txtData(Index)
        Case Editor.Pontos: Player.Points = txtData(Index)
        Case Editor.LevelVIP: Player.VIP = txtData(Index)
    
    End Select
    
    txtData(Index).BackColor = Edited
End Sub

Private Sub txtHair_Change()
    Player.Hair = Val(txtHair)
End Sub
