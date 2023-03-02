VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Índice de Quest"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data"
      Height          =   5535
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   6015
      Begin VB.Frame Frame6 
         Caption         =   "Talks"
         Height          =   1215
         Left            =   120
         TabIndex        =   24
         Top             =   4200
         Width           =   5775
         Begin VB.TextBox txtCompleted 
            Height          =   285
            Left            =   1320
            TabIndex        =   28
            Top             =   720
            Width           =   4335
         End
         Begin VB.TextBox txtStartTalk 
            Height          =   285
            Left            =   960
            TabIndex        =   27
            Top             =   360
            Width           =   4695
         End
         Begin VB.Label Label11 
            Caption         =   "Completed Talk:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Start Talk:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":0000
         Left            =   1200
         List            =   "frmEditor_Quest.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2400
         Width           =   4575
      End
      Begin VB.TextBox txtEvent 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox chkRepeat 
         Caption         =   "Repeat"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1200
         Width           =   4335
      End
      Begin VB.HScrollBar scrlIcon 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   5055
      End
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
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
         Height          =   480
         Left            =   5280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         Top             =   600
         Width           =   480
      End
      Begin VB.Frame Frame3 
         Caption         =   "Requisites"
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   5775
         Begin VB.TextBox txtHours 
            Height          =   285
            Left            =   2880
            TabIndex        =   18
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkNight 
            Caption         =   "Only on night"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   2295
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Only on day"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label4 
            Caption         =   "Time until do it again (Soon):"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   3135
         End
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label6 
         Caption         =   "Negative values show an variable"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label lblEventName 
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Information:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblIcon 
         Caption         =   "Icon: None"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Quest"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.ListBox lstIndex 
         Height          =   3765
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkDay_Click()
    Quest(EditorIndex).NotDay = chkDay.Value
End Sub

Private Sub chkNight_Click()
    Quest(EditorIndex).NotNight = chkNight.Value
End Sub

Private Sub chkRepeat_Click()
    Quest(EditorIndex).Repeat = chkRepeat.Value
End Sub



Private Sub cmbType_Click()
    Quest(EditorIndex).Type = cmbType.ListIndex
End Sub

Private Sub Command1_Click()
    QuestEditorOk EditorIndex
End Sub

Private Sub Command2_Click()
    Editor = 0
    Unload Me
End Sub

Private Sub lstIndex_Click()
    EditorIndex = lstIndex.ListIndex + 1
    QuestEditorInit
End Sub

Private Sub scrlData1_Change()
    Dim Prefix As String
    Dim IsNPC As Boolean
    
    IsNPC = (cmbObjectiveType.ListIndex = QuestObjectiveType.KillNPC Or cmbObjectiveType.ListIndex = QuestObjectiveType.TalkToNPC)

    If IsNPC Then
        Prefix = "NPC: "
    Else
        Prefix = "Item: "
    End If
    
    If scrlData1.Value > 0 Then
        lblData1.Caption = Prefix & scrlData1.Value
    Else
        lblData1.Caption = Prefix & "None"
    End If
End Sub


Private Sub scrlIcon_Change()
    Quest(EditorIndex).Icon = scrlIcon.Value
    lblIcon.Caption = "Icon: " & scrlIcon.Value
End Sub

Private Sub Text3_Change()
    
End Sub




Private Sub txtDesc_Change()
    Quest(EditorIndex).Desc = txtDesc.Text
End Sub

Private Sub txtEvent_Change()
    If IsNumeric(txtEvent) Then
        If txtEvent > MAX_EVENTS Then Exit Sub
        If Val(txtEvent) > 0 Then
            lblEventName.Caption = "Nome: " & Trim$(Events(txtEvent).Name)
            Quest(EditorIndex).EventNum = txtEvent.Text
        Else
            If Val(txtEvent) <> 0 Then
                lblEventName.Caption = "Nome: " & Trim$(Variables(-txtEvent))
                Quest(EditorIndex).EventNum = txtEvent.Text
            End If
        End If
    End If
End Sub

Private Sub txtHours_Change()
    If IsNumeric(txtHours) Then
        Quest(EditorIndex).Cooldown = txtHours.Text
    End If
End Sub

Private Sub txtName_Change()
    Quest(EditorIndex).Name = txtName.Text
End Sub

Private Sub LoadObjective(TaskIndex As Long)
    Dim IsTalkToNpc As Boolean
    Dim Task As QuestTaskRec
    Dim IsNPC As Boolean
    
    IsNPC = (cmbObjectiveType.ListIndex = QuestObjectiveType.KillNPC Or cmbObjectiveType.ListIndex = QuestObjectiveType.TalkToNPC)

    lblData2.visible = Not IsTalkToNpc
    txtData2.visible = Not IsTalkToNpc
    lblFinishNPC.visible = Not IsTalkToNpc
    scrlFinishNPC.visible = Not IsTalkToNpc
    
    If IsNPC Then
        scrlData1.max = MAX_NPCS
    Else
        scrlData1.max = MAX_ITEMS
    End If
    
    scrlData1.Value = Task.Data1
    txtData2.Text = Val(Task.Data2)
    txtData3.Text = Task.Data3
    scrlFinishNPC.Value = Task.FinishNPCQuest
    chkComplete.Value = Task.CompleteQuest
End Sub

