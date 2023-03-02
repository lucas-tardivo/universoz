VERSION 5.00
Begin VB.Form frmAdminPanel 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Panel"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAdminPanel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   0
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   438
      TabIndex        =   0
      Top             =   0
      Width           =   6570
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   5400
         TabIndex        =   39
         Text            =   "1"
         Top             =   2040
         Width           =   975
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   2760
         Min             =   1
         TabIndex        =   29
         Top             =   2040
         Value           =   1
         Width           =   2655
      End
      Begin VB.TextBox txtAMap 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3480
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtASprite 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3480
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtAName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtAAccess 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3480
         TabIndex        =   1
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Editores"
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   2415
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Quest"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Notícias"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   13
            Left            =   1200
            TabIndex        =   34
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuda"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   12
            Left            =   1200
            TabIndex        =   33
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Del Bans"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   28
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Atualizar"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   10
            Left            =   1200
            TabIndex        =   27
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Map Report"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Animações"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Magias"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   7
            Left            =   1200
            TabIndex        =   24
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Shop"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Recursos"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   1200
            TabIndex        =   22
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "NPC"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Eventos"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblCommand 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Efeitos"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Left            =   0
         TabIndex        =   38
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblGiveItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dar item ao player(Sem Mensagem)"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   36
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label lblGiveItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dar item ao player"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   35
         Top             =   2640
         Width           =   3615
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Item: None"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   32
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "#: 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   31
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblSpawnItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dropar item"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   30
         Top             =   2400
         Width           =   3615
      End
      Begin VB.Label lblSprite 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Sprite"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblAccess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Setar acess"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblKick 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kickar"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblWarp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Teleportar"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblBan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Banir"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblWarpMe2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ir para"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblWarp2Me 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trazer"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa#:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite#:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Acesso:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAdminPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()
    SendRequestEditQuest
End Sub

Private Sub Label2_Click()
    If App.LogMode = 0 Then
        paperdollTestin = Not paperdollTestin
        frmPaperdoll.Show
        Unload Me
    End If
End Sub

' ****************
' ** Admin Menu **
' ****************
Private Sub lblCommand_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCommand(Index).ForeColor = &HFFFF&
End Sub

Private Sub lblCommand_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 12 Then
        lblCommand(Index).ForeColor = &HFF00&
    Else
        lblCommand(Index).ForeColor = &HFFFFFF
    End If
End Sub
Private Sub lblCommand_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 0
            If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
            SendRequestEditEffect
        Case 1
            If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
            Call RequestSwitchesAndVariables
            Call Events_SendRequestEventsData
            Call Events_SendRequestEditEvents
        Case 2
            If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
            SendRequestEditItem
        Case 3
            If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
            SendRequestEditMap
        Case 4
            If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
            SendRequestEditNpc
        Case 5
            If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
            SendRequestEditResource
        Case 6
            If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
            SendRequestEditShop
        Case 7
            If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
            SendRequestEditSpell
        Case 8
            If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
            SendRequestEditAnimation
        Case 9
            If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
            SendMapReport
        Case 10
            If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
            SendMapRespawn
            Exit Sub
        Case 11
            If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
            SendBanDestroy
            Exit Sub
        Case 12
            frmHelp.visible = True
        Case 13
            If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
            SendRequestNews
            SendRequestEditNews
    End Select
    frmAdminPanel.visible = False
    frmMain.lblAdminPanel.Tag = 0
    frmMain.lblAdminPanel.Caption = "Open Admin Panel"

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCommand_Click", "frmAdminPanel", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblGiveItem_Click(Index As Integer)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    
    If Index = 0 Then
        SendSpawnItem scrlAItem.Value, Val(txtAmount), txtAName.Text, 0
    Else
        SendSpawnItem scrlAItem.Value, Val(txtAmount), txtAName.Text, 1
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblSpawnItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblKick_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then Exit Sub
    
    SendKick Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblKick_Click", "frmAdminPanel", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblWarp2Me_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblWarpMe2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblWarp_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAMap.Text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.Text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.Text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblSprite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub

    If Len(Trim$(txtASprite.Text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtASprite.Text)) Then
        Exit Sub
    End If

    SendSetSprite CLng(Trim$(txtASprite.Text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblSprite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblBan_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblBan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Or Not IsNumeric(Trim$(txtAAccess.Text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.Text), CLng(Trim$(txtAAccess.Text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblSpawnItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    SendSpawnItem scrlAItem.Value, Val(txtAmount)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblSpawnItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblWarpMeTo_Click()

End Sub

Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAAmount.Caption = "#: " & scrlAAmount.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAAmount_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAItem.Caption = "Item(" & scrlAItem.Value & "): " & Trim$(Item(scrlAItem.Value).Name)
    If Item(scrlAItem.Value).Type = ITEM_TYPE_CURRENCY Or Item(scrlAItem.Value).Stackable > 0 Then
        txtAmount.Enabled = True
        Exit Sub
    End If
    txtAmount.Enabled = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAmount_Change()
    If Not IsNumeric(txtAmount) Then txtAmount = 0
End Sub
