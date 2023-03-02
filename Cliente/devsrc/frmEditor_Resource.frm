VERSION 5.00
Begin VB.Form frmEditor_Resource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
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
   Icon            =   "frmEditor_Resource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   26
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   25
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   24
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data"
      Height          =   7575
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   52
         Top             =   240
         Width           =   255
      End
      Begin VB.Frame frmPlanetable 
         Caption         =   "Planetas próprios"
         Height          =   4215
         Left            =   120
         TabIndex        =   34
         Top             =   3240
         Visible         =   0   'False
         Width           =   4815
         Begin VB.TextBox txtCentroLevel 
            Height          =   270
            Left            =   3600
            TabIndex        =   56
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtRLevel 
            Height          =   270
            Left            =   3600
            TabIndex        =   54
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtNucleo 
            Height          =   270
            Left            =   3600
            TabIndex        =   51
            Top             =   240
            Width           =   1095
         End
         Begin VB.HScrollBar scrlEvolution 
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1080
            Width           =   4575
         End
         Begin VB.Frame Frame8 
            Caption         =   "Custo da evolução"
            Height          =   1695
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   4575
            Begin VB.TextBox txtMoedas 
               Height          =   270
               Left            =   3360
               TabIndex        =   41
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtRed 
               Height          =   270
               Left            =   3360
               TabIndex        =   40
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtBlue 
               Height          =   270
               Left            =   3360
               TabIndex        =   39
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox txtYellow 
               Height          =   270
               Left            =   3360
               TabIndex        =   38
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label lblGold 
               Caption         =   "Moedas necessárias:"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   4215
            End
            Begin VB.Label Label14 
               Caption         =   "Especiarias vermelhas necessárias:"
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   600
               Width           =   4215
            End
            Begin VB.Label Label17 
               Caption         =   "Especiarias azuis necessárias:"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   960
               Width           =   4215
            End
            Begin VB.Label Label18 
               Caption         =   "Especiarias amarelas necessárias:"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   1320
               Width           =   4215
            End
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Fechar"
            Height          =   255
            Left            =   3120
            TabIndex        =   36
            Top             =   3720
            Width           =   1575
         End
         Begin VB.TextBox txtMinutes 
            Height          =   270
            Left            =   3480
            TabIndex        =   35
            Top             =   3240
            Width           =   1095
         End
         Begin VB.CheckBox chkPlanetable 
            Caption         =   "É de planetas próprios"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Centro level:"
            Height          =   255
            Left            =   2160
            TabIndex        =   55
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Level desta construção:"
            Height          =   255
            Left            =   1680
            TabIndex        =   53
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label6 
            Caption         =   "Nucleo level:"
            Height          =   255
            Left            =   2520
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblEvolution 
            Caption         =   "Evolui para:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   840
            Width           =   4575
         End
         Begin VB.Label Label19 
            Caption         =   "Tempo em minutos para evoluir:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   3240
            Width           =   3375
         End
      End
      Begin VB.TextBox txtRespawnTime 
         Height          =   270
         Left            =   2280
         TabIndex        =   33
         Top             =   5880
         Width           =   2655
      End
      Begin VB.HScrollBar scrlEffect 
         Height          =   255
         Left            =   3000
         Max             =   5
         TabIndex        =   31
         Top             =   6720
         Width           =   1935
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   7080
         Width           =   3975
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   3000
         Max             =   6000
         TabIndex        =   27
         Top             =   6480
         Width           =   1935
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   1920
         Width           =   2295
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   2640
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   21
         Top             =   2280
         Width           =   2280
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":3332
         Left            =   960
         List            =   "frmEditor_Resource.frx":334B
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   3975
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   2295
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4320
         Width           =   4815
      End
      Begin VB.HScrollBar scrlTool 
         Height          =   255
         Left            =   120
         Max             =   3
         TabIndex        =   7
         Top             =   4920
         Width           =   4815
      End
      Begin VB.HScrollBar scrlHealth 
         Height          =   255
         Left            =   120
         Max             =   32000
         TabIndex        =   6
         Top             =   5520
         Width           =   4815
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   5
         Top             =   2280
         Width           =   2280
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtMessage2 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label lblEffect 
         AutoSize        =   -1  'True
         Caption         =   "Effect: None"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   6720
         Width           =   930
      End
      Begin VB.Label Label5 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   6480
         Width           =   1260
      End
      Begin VB.Label lblExhaustedPic 
         AutoSize        =   -1  'True
         Caption         =   "Exhausted Image: 0"
         Height          =   180
         Left            =   2640
         TabIndex        =   23
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lblNormalPic 
         AutoSize        =   -1  'True
         Caption         =   "Normal Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1470
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         Caption         =   "Item Reward: None"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   4080
         Width           =   1440
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   4680
         Width           =   1530
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         Caption         =   "Health: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   5280
         Width           =   705
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time (Seconds): "
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   5880
         Width           =   1995
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lista"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPlanetable_Click()
    Resource(EditorIndex).IsPlanetable = chkPlanetable.Value
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).ResourceType = cmbType.ListIndex
    
    If Resource(EditorIndex).ResourceType = 5 Then
        UpdateConstruction
    Else
        frmPlanetable.visible = False
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub UpdateConstruction()
    frmPlanetable.visible = True
    chkPlanetable.Value = Resource(EditorIndex).IsPlanetable
    scrlEvolution.max = MAX_RESOURCES
    scrlEvolution.Value = Resource(EditorIndex).Evolution
    txtMoedas.Text = Val(Resource(EditorIndex).ECostGold)
    txtRed.Text = Val(Resource(EditorIndex).ECostRed)
    txtBlue.Text = Val(Resource(EditorIndex).ECostBlue)
    txtYellow.Text = Val(Resource(EditorIndex).ECostYellow)
    txtMinutes.Text = Val(Resource(EditorIndex).TimeToEvolute)
    txtNucleo.Text = Val(Resource(EditorIndex).NucleoLevel)
    txtRLevel.Text = Val(Resource(EditorIndex).ResourceLevel)
    txtCentroLevel.Text = Val(Resource(EditorIndex).MinLevel)
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearResource EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ResourceEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Command1_Click()
    Dim i As String
    i = InputBox("Copy resource:", "Universo Z")
    i = Val(i)
    If IsNumeric(i) Then
        If i > 0 And i <= MAX_RESOURCES Then
            Resource(EditorIndex) = Resource(i)
            ResourceEditorInit
            'SpellEditorInit
        End If
    End If
End Sub

Private Sub Command11_Click()
    frmPlanetable.visible = False
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlReward.max = MAX_ITEMS
    scrlEffect.max = MAX_EFFECTS
    scrlAnimation.max = MAX_ANIMATIONS
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ResourceEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ResourceEditorInit
    If Resource(EditorIndex).ResourceType = 5 Then
        UpdateConstruction
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnim.Caption = "Animation: " & sString
    Resource(EditorIndex).Animation = scrlAnimation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    Resource(EditorIndex).Effect = scrlEffect.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEffect_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlEvolution_Change()
    If scrlEvolution.Value > 0 Then
        lblEvolution.Caption = "Evolui para: " & scrlEvolution.Value & " " & Trim$(Resource(scrlEvolution.Value).Name)
    Else
        lblEvolution.Caption = "Evolui para: <Nenhum>"
    End If
    Resource(EditorIndex).Evolution = scrlEvolution.Value
End Sub

Private Sub scrlExhaustedPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblExhaustedPic.Caption = "Vazio: " & scrlExhaustedPic.Value
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlExhaustedPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlHealth_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblHealth.Caption = "HP: " & scrlHealth.Value
    Resource(EditorIndex).health = scrlHealth.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlHealth_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNormalPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNormalPic.Caption = "Normal Image: " & scrlNormalPic.Value
    Resource(EditorIndex).ResourceImage = scrlNormalPic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNormalPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlReward_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlReward.Value > 0 Then
        lblReward.Caption = "Item: " & Trim$(Item(scrlReward.Value).Name)
    Else
        lblReward.Caption = "Item: None"
    End If
    
    Resource(EditorIndex).ItemReward = scrlReward.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlReward_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTool_Change()
    Dim Name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlTool.Value
        Case 0
            Name = "None"
        Case 1
            Name = "Axe"
        Case 2
            Name = "Fishing rod"
        Case 3
            Name = "Pickaxe"
    End Select

    lblTool.Caption = "Tool: " & Name
    
    Resource(EditorIndex).ToolRequired = scrlTool.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtBlue_Change()
    If IsNumeric(txtBlue) Then
        Resource(EditorIndex).ECostBlue = Val(txtBlue)
    End If
End Sub

Private Sub txtCentroLevel_Change()
    Resource(EditorIndex).MinLevel = Val(txtCentroLevel)
End Sub

Private Sub txtMessage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).SuccessMessage = Trim$(txtMessage.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage2_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).EmptyMessage = Trim$(txtMessage2.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage2_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMinutes_Change()
    Resource(EditorIndex).TimeToEvolute = Val(txtMinutes)
End Sub

Private Sub txtMoedas_Change()
    If IsNumeric(txtMoedas) Then
        Resource(EditorIndex).ECostGold = Val(txtMoedas)
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Resource(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Resource(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtNucleo_Change()
    Resource(EditorIndex).NucleoLevel = Val(txtNucleo)
End Sub

Private Sub txtRed_Change()
    If IsNumeric(txtRed) Then
        Resource(EditorIndex).ECostRed = Val(txtRed)
    End If
End Sub

Private Sub txtRespawnTime_Change()
    Resource(EditorIndex).RespawnTime = Val(txtRespawnTime.Text)
End Sub

Private Sub txtRLevel_Change()
    Resource(EditorIndex).ResourceLevel = Val(txtRLevel)
End Sub

Private Sub txtYellow_Change()
    If IsNumeric(txtYellow) Then
        Resource(EditorIndex).ECostYellow = Val(txtYellow)
    End If
End Sub
