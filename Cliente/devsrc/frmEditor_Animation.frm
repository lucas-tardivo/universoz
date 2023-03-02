VERSION 5.00
Begin VB.Form frmEditor_Animation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation editor"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
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
   ScaleHeight     =   8295
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmbCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   5400
      TabIndex        =   43
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8520
      TabIndex        =   20
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   7575
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      Begin VB.HScrollBar scrlBuraco 
         Height          =   255
         Left            =   2400
         Max             =   10
         TabIndex        =   41
         Top             =   960
         Width           =   4095
      End
      Begin VB.HScrollBar scrlTremor 
         Height          =   255
         Left            =   2400
         Max             =   30000
         TabIndex        =   40
         Top             =   600
         Width           =   4095
      End
      Begin VB.HScrollBar scrlXAxis 
         Height          =   255
         Left            =   1080
         Max             =   256
         Min             =   -256
         TabIndex        =   36
         Top             =   1920
         Width           =   2175
      End
      Begin VB.HScrollBar scrlYAxis 
         Height          =   255
         Left            =   4320
         Max             =   256
         Min             =   -256
         TabIndex        =   35
         Top             =   1920
         Width           =   2175
      End
      Begin VB.OptionButton optDir 
         Caption         =   "All"
         Height          =   375
         Index           =   4
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Right"
         Height          =   375
         Index           =   3
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Left"
         Height          =   375
         Index           =   2
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Down"
         Height          =   375
         Index           =   1
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Up"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   240
         Width           =   2415
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   27
         Top             =   4560
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   4560
         Width           =   3135
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   2535
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   17
         Top             =   3960
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   15
         Top             =   3360
         Width           =   3135
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Index           =   1
         Left            =   3360
         ScaleHeight     =   2475
         ScaleWidth      =   3075
         TabIndex        =   13
         Top             =   4920
         Width           =   3135
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   12
         Top             =   2760
         Width           =   3135
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   3960
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Width           =   3135
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Index           =   0
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   3075
         TabIndex        =   6
         Top             =   4920
         Width           =   3135
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label lblBuraco 
         Caption         =   "Hole: Nenhum"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblTremor 
         Caption         =   "Shake: Nenhum"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblXAxis 
         Caption         =   "X-Axis:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblYAxis 
         Caption         =   "Y-Axis:"
         Height          =   255
         Left            =   3360
         TabIndex        =   37
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   26
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Above player"
         Height          =   180
         Left            =   3360
         TabIndex        =   18
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   16
         Top             =   3720
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   14
         Top             =   3120
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   11
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   3720
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Below player"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   2280
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List"
      Height          =   8055
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
         Height          =   7665
         ItemData        =   "frmEditor_Animation.frx":0000
         Left            =   120
         List            =   "frmEditor_Animation.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DirSelect As Byte

Private Sub cmbCopy_Click()
    Dim n As Long
    n = Val(InputBox("Digite o numero da animação"))
    Animation(EditorIndex) = Animation(n)
    AnimationEditorInit
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Animation(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Animation(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    AnimationEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    
    ClearAnimation EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    AnimationEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    AnimationEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 0 To 1
        scrlSprite(i).max = NumAnimations
        scrlLoopCount(i).max = 100
        scrlFrameCount(i).max = 100
        scrlLoopTime(i).max = 1000
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    AnimationEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optDir_Click(Index As Integer)
Dim i As Long
    DirSelect = Index
    
    If Index < 4 Then
        For i = 0 To 1
            scrlSprite(i).Value = Animation(EditorIndex).Sprite(i, Index)
        Next i
        scrlXAxis.Value = Animation(EditorIndex).XAxis(Index)
        scrlYAxis.Value = Animation(EditorIndex).YAxis(Index)
    End If
End Sub

Private Sub scrlBuraco_Change()
    If scrlBuraco.Value > 0 Then
        lblBuraco.Caption = "Hole: size " & scrlBuraco.Value
    Else
        lblBuraco.Caption = "Hole: None"
    End If
    
    Animation(EditorIndex).Buraco = scrlBuraco.Value
End Sub

Private Sub scrlFrameCount_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFrameCount(Index).Caption = "Frames: " & scrlFrameCount(Index).Value
    Animation(EditorIndex).Frames(Index) = scrlFrameCount(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFrameCount_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFrameCount_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlFrameCount_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFrameCount_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopCount_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblLoopCount(Index).Caption = "Loops: " & scrlLoopCount(Index).Value
    Animation(EditorIndex).LoopCount(Index) = scrlLoopCount(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopCount_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopCount_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlLoopCount_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopCount_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopTime_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblLoopTime(Index).Caption = "Loop time: " & scrlLoopTime(Index).Value
    Animation(EditorIndex).looptime(Index) = scrlLoopTime(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopTime_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopTime_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlLoopTime_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopTime_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite(Index).Caption = "Sprite: " & scrlSprite(Index).Value
    
    If DirSelect < 4 Then
        Animation(EditorIndex).Sprite(Index, DirSelect) = scrlSprite(Index).Value
    Else
        Dim i As Long
        For i = 0 To 3
            Animation(EditorIndex).Sprite(Index, i) = scrlSprite(Index).Value
        Next i
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTremor_Change()
    If scrlTremor.Value > 0 Then
        lblTremor.Caption = "Shake: " & scrlTremor.Value & " milisecs"
    Else
        lblTremor.Caption = "Shake: None"
    End If
    Animation(EditorIndex).Tremor = scrlTremor.Value
End Sub

Private Sub scrlXAxis_Change()
    Dim i As Long
    lblXAxis.Caption = "X-Axis: " & scrlXAxis.Value
    If DirSelect < 4 Then
        Animation(EditorIndex).XAxis(DirSelect) = scrlXAxis.Value
    Else
        For i = 0 To 3
            Animation(EditorIndex).XAxis(i) = scrlXAxis.Value
        Next i
    End If
End Sub

Private Sub scrlYAxis_Change()
    Dim i As Long
    lblYAxis.Caption = "Y-Axis: " & scrlYAxis.Value
    If DirSelect < 4 Then
        Animation(EditorIndex).YAxis(DirSelect) = scrlYAxis.Value
    Else
        For i = 0 To 3
            Animation(EditorIndex).YAxis(i) = scrlYAxis.Value
        Next i
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Animation(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
