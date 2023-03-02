VERSION 5.00
Begin VB.Form frmAdminPanel 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Panel"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAdminPanel 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   5715
      Left            =   0
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   542
      TabIndex        =   0
      Top             =   0
      Width           =   8130
      Begin VB.TextBox txtQuant 
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         Top             =   2640
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   1560
         Min             =   1
         TabIndex        =   17
         Top             =   2040
         Value           =   1
         Width           =   5055
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         LargeChange     =   10
         Left            =   1560
         Min             =   1
         TabIndex        =   16
         Top             =   2640
         Value           =   1
         Width           =   5055
      End
      Begin VB.TextBox txtAMap 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   2400
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtASprite 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   2400
         TabIndex        =   3
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtAName 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox txtAAccess 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   2400
         TabIndex        =   1
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Image Image2 
         Height          =   630
         Left            =   4560
         Picture         =   "frmAdminPanel.frx":0000
         Top             =   3480
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblGiveItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Give to player (Silence)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   22
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label lblGiveItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Give to player (Above)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   21
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Item: None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "#: 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblSpawnItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Drop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label lblSprite 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Sprite"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblAccess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Access"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblKick 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kick"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblWarp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Teleport"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblBan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ban"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblWarpMe2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Go to"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblWarp2Me 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bring"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map#:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAdminPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    scrlAItem.max = MAX_ITEMS
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

End Sub

Private Sub lblCommand_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    
    frmAdminPanel.visible = False

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
    
    Dim Quant As Long
    If txtQuant.visible = True Then
        Quant = Val(txtQuant)
    Else
        Quant = scrlAAmount.Value
    End If
    
    If Index = 0 Then
        SendSpawnItem scrlAItem.Value, Quant, txtAName.Text, 0
    Else
        SendSpawnItem scrlAItem.Value, Quant, txtAName.Text, 1
    End If
    
    AddText "Item enviado com sucesso!", White
    
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
    
    
    Dim Quant As Long
    If txtQuant.visible = True Then
        Quant = Val(txtQuant)
    Else
        Quant = scrlAAmount.Value
    End If
    
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    SendSpawnItem scrlAItem.Value, Quant
    
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
    If scrlAAmount.Value = scrlAAmount.max Then
        txtQuant.visible = True
    End If
    
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

    If scrlAItem.Value > MAX_ITEMS Then scrlAItem.Value = MAX_ITEMS
    lblAItem.Caption = "Item(" & scrlAItem.Value & "): " & Trim$(Item(scrlAItem.Value).Name)
    txtQuant.visible = False
    If Item(scrlAItem.Value).Type = ITEM_TYPE_CURRENCY Or Item(scrlAItem.Value).Stackable > 0 Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    scrlAAmount.Enabled = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
