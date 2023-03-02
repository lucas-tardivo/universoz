VERSION 5.00
Begin VB.Form frmEditor_Effect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Effects editor"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
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
   ScaleHeight     =   6975
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   615
      Left            =   3360
      TabIndex        =   33
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   34
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3480
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Type"
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Width           =   6735
      Begin VB.OptionButton optEffectType 
         Caption         =   "Multi"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optEffectType 
         Caption         =   "Effect"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "List"
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5910
         ItemData        =   "frmEditor_Effect.frx":0000
         Left            =   120
         List            =   "frmEditor_Effect.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraMultiParticle 
      Caption         =   "Multi"
      Height          =   975
      Left            =   3360
      TabIndex        =   28
      Top             =   720
      Visible         =   0   'False
      Width           =   6735
      Begin VB.HScrollBar scrlEffect 
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   480
         Value           =   1
         Width           =   4335
      End
      Begin VB.HScrollBar scrlMultiParticle 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   31
         Top             =   480
         Value           =   1
         Width           =   2055
      End
      Begin VB.Label lblEffect 
         Caption         =   "Effect: XXXXXX"
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblMultiParticle 
         Caption         =   "Multi-Particle: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraEffect 
      Caption         =   "Efeito"
      Height          =   4815
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   6735
      Begin VB.HScrollBar scrlDuration 
         Height          =   255
         Left            =   1920
         Max             =   255
         TabIndex        =   39
         Top             =   960
         Width           =   1335
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Effect.frx":0004
         Left            =   4080
         List            =   "frmEditor_Effect.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   2535
      End
      Begin VB.HScrollBar scrlParticles 
         Height          =   255
         LargeChange     =   10
         Left            =   1920
         Max             =   5000
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.HScrollBar scrlSize 
         Height          =   255
         Left            =   5280
         Max             =   255
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Graphic"
         Height          =   3495
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   6495
         Begin VB.HScrollBar scrlYAcc 
            Height          =   255
            Left            =   1320
            Max             =   100
            Min             =   -100
            TabIndex        =   46
            Top             =   3120
            Width           =   5055
         End
         Begin VB.HScrollBar scrlXAcc 
            Height          =   255
            Left            =   1320
            Max             =   100
            Min             =   -100
            TabIndex        =   44
            Top             =   2760
            Width           =   5055
         End
         Begin VB.HScrollBar scrlYSpeed 
            Height          =   255
            Left            =   1320
            Max             =   100
            Min             =   -100
            TabIndex        =   42
            Top             =   2400
            Width           =   5055
         End
         Begin VB.HScrollBar scrlXSpeed 
            Height          =   255
            Left            =   1320
            Max             =   100
            Min             =   -100
            TabIndex        =   40
            Top             =   2040
            Width           =   5055
         End
         Begin VB.HScrollBar scrlAlpha 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   14
            Top             =   240
            Width           =   5055
         End
         Begin VB.HScrollBar scrlDecay 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   13
            Top             =   600
            Width           =   5055
         End
         Begin VB.HScrollBar scrlRed 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   12
            Top             =   960
            Width           =   5055
         End
         Begin VB.HScrollBar scrlGreen 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   11
            Top             =   1320
            Width           =   5055
         End
         Begin VB.HScrollBar scrlBlue 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   10
            Top             =   1680
            Width           =   5055
         End
         Begin VB.Label lblYAcc 
            AutoSize        =   -1  'True
            Caption         =   "YAcc: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   47
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label lblXAcc 
            AutoSize        =   -1  'True
            Caption         =   "XAcc: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   45
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label lblYSpeed 
            AutoSize        =   -1  'True
            Caption         =   "YSpeed: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   43
            Top             =   2400
            Width           =   780
         End
         Begin VB.Label lblXSpeed 
            AutoSize        =   -1  'True
            Caption         =   "XSpeed: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   41
            Top             =   2040
            Width           =   780
         End
         Begin VB.Label lblAlpha 
            AutoSize        =   -1  'True
            Caption         =   "Alpha: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblDecay 
            AutoSize        =   -1  'True
            Caption         =   "Decay: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   690
         End
         Begin VB.Label lblRed 
            AutoSize        =   -1  'True
            Caption         =   "Red: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lblGreen 
            AutoSize        =   -1  'True
            Caption         =   "Green: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   660
         End
         Begin VB.Label lblBlue 
            AutoSize        =   -1  'True
            Caption         =   "Blue: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   540
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visualize"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   5070
         TabIndex        =   48
         Top             =   960
         Width           =   915
      End
      Begin VB.Label lblDuration 
         AutoSize        =   -1  'True
         Caption         =   "Duration: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Type:"
         Height          =   255
         Left            =   3480
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblParticles 
         AutoSize        =   -1  'True
         Caption         =   "Particles: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "Size: 0"
         Height          =   180
         Left            =   3480
         TabIndex        =   24
         Top             =   600
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmEditor_Effect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Effect(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Effect(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Effect(EditorIndex).Type = cmbType.ListIndex + 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EffectEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    
    ClearEffect EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Effect(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    EffectEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EffectEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite.max = NumParticles
    scrlMultiParticle.max = MAX_MULTIPARTICLE
    scrlEffect.max = MAX_EFFECTS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Label4_Click()
    CastEffect EditorIndex, GetPlayerX(MyIndex), GetPlayerY(MyIndex)
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EffectEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optEffectType_Click(Index As Integer)
    Select Case Index
        Case 0
            fraMultiParticle.visible = False
            fraEffect.visible = True
        Case 1
            fraMultiParticle.visible = True
            fraEffect.visible = False
    End Select
    Effect(EditorIndex).isMulti = Index
End Sub

Private Sub scrlAlpha_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAlpha.Caption = "Alpha: " & scrlAlpha.Value / 100
    Effect(EditorIndex).Alpha = scrlAlpha.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAlpha_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBlue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblBlue.Caption = "Blue: " & scrlBlue.Value / 100
    Effect(EditorIndex).Blue = scrlBlue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlBlue_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDecay_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblDecay.Caption = "Decay: " & scrlDecay.Value / 100
    Effect(EditorIndex).Decay = scrlDecay.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDecay_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDuration_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblDuration.Caption = "Duration: " & scrlDuration.Value
    Effect(EditorIndex).Duration = scrlDuration.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDuration_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlEffect_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlEffect.Value > 0 Then
        lblEffect.Caption = "Effect: " & Trim$(Effect(scrlEffect.Value).Name)
    Else
        lblEffect.Caption = "Effect: None"
    End If
    
    Effect(EditorIndex).MultiParticle(scrlMultiParticle.Value) = scrlEffect.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEffect_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlGreen_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblGreen.Caption = "Green: " & scrlGreen.Value / 100
    Effect(EditorIndex).Green = scrlGreen.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlGreen_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMultiParticle_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblMultiParticle.Caption = "Multi-particle: " & scrlMultiParticle.Value
    scrlEffect.Value = Effect(EditorIndex).MultiParticle(scrlMultiParticle)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMultiParticle_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlParticles_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblParticles.Caption = "Particles: " & scrlParticles.Value
    Effect(EditorIndex).Particles = scrlParticles.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlParticles_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRed.Caption = "Red: " & scrlRed.Value / 100
    Effect(EditorIndex).Red = scrlRed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRed_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSize_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSize.Caption = "Size: " & scrlSize.Value
    Effect(EditorIndex).Size = scrlSize.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Effect(EditorIndex).Sprite = scrlSprite.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Scroll", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlXSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblXSpeed.Caption = "XSpeed: " & scrlXSpeed.Value
    Effect(EditorIndex).XSpeed = scrlXSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlXSpeed_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub scrlYSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblYSpeed.Caption = "YSpeed: " & scrlYSpeed.Value
    Effect(EditorIndex).YSpeed = scrlYSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlYSpeed_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub scrlXAcc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblXAcc.Caption = "XAcc: " & scrlXAcc.Value
    Effect(EditorIndex).XAcc = scrlXAcc.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlXAcc_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub scrlYAcc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblYAcc.Caption = "YAcc: " & scrlYAcc.Value
    Effect(EditorIndex).YAcc = scrlYAcc.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlYAcc_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Effect(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Effect(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
