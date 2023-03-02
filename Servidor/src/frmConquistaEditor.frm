VERSION 5.00
Begin VB.Form frmConquistaEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor de conquistas"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7575
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
   ScaleHeight     =   4455
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.Frame Frame3 
         Caption         =   "Recompensas"
         Height          =   2295
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   4335
         Begin VB.TextBox txtItem 
            Height          =   285
            Left            =   3120
            TabIndex        =   20
            Top             =   1440
            Width           =   1095
         End
         Begin VB.HScrollBar scrlItem 
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   2895
         End
         Begin VB.TextBox txtQuant 
            Height          =   285
            Left            =   1320
            TabIndex        =   16
            Top             =   1920
            Width           =   1335
         End
         Begin VB.HScrollBar scrlItemIndex 
            Height          =   255
            Left            =   120
            Max             =   5
            Min             =   1
            TabIndex        =   14
            Top             =   840
            Value           =   1
            Width           =   4095
         End
         Begin VB.TextBox txtExp 
            Height          =   285
            Left            =   1320
            TabIndex        =   12
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblItemNumber 
            Caption         =   "Número:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Quantidade:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label lblItem 
            Caption         =   "Item"
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
            TabIndex        =   13
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label Label4 
            Caption         =   "Experiência:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox txtProgress 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Progressão:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conquistas"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command3 
         Caption         =   "Clonar"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Adicionar nova"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3600
         Width           =   2415
      End
      Begin VB.ListBox lstIndex 
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmConquistaEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private EditorIndex As Long

Private Sub Command1_Click()
    ReDim Preserve Conquistas(1 To UBound(Conquistas) + 1)
    PopulateList
    lstIndex.ListIndex = lstIndex.ListCount - 1
    EditorIndex = lstIndex.ListIndex + 1
    LoadConquista
End Sub

Private Sub Command2_Click()
    Dim i As Long
    For i = 1 To UBound(Conquistas)
        Dim filename As String
        filename = App.path & "\data\Conquistas.ini"
        PutVar filename, "CONQUISTA" & i, "Name", Conquistas(i).Name
        PutVar filename, "CONQUISTA" & i, "Desc", Conquistas(i).Desc
        PutVar filename, "CONQUISTA" & i, "Exp", Val(Conquistas(i).Exp)
        PutVar filename, "CONQUISTA" & i, "Progress", Val(Conquistas(i).Progress)
        
        Dim n As Long
        For n = 1 To 5
            PutVar filename, "CONQUISTA" & i, "Reward" & n, Val(Conquistas(i).Reward(n).Num)
            PutVar filename, "CONQUISTA" & i, "Value" & n, Val(Conquistas(i).Reward(n).Value)
        Next n
    Next i
    LoadConquistas
    Unload Me
End Sub

Private Sub Command3_Click()
    ReDim Preserve Conquistas(1 To UBound(Conquistas) + 1)
    Conquistas(UBound(Conquistas)) = Conquistas(lstIndex.ListIndex + 1)
    PopulateList
    lstIndex.ListIndex = lstIndex.ListCount - 1
    EditorIndex = lstIndex.ListIndex + 1
    LoadConquista
End Sub

Private Sub Form_Load()
    PopulateList
    EditorIndex = 1
    LoadConquista
    scrlItem.Max = MAX_ITEMS
End Sub
Private Sub PopulateList()
    Dim i As Long
    lstIndex.Clear
    For i = 1 To UBound(Conquistas)
        lstIndex.AddItem i & ": " & Trim$(Conquistas(i).Name)
    Next i
End Sub
Private Sub LoadConquista()
    txtNome = Trim$(Conquistas(EditorIndex).Name)
    txtDesc = Conquistas(EditorIndex).Desc
    txtProgress = Conquistas(EditorIndex).Progress
    txtExp = Conquistas(EditorIndex).Exp
    scrlItemIndex.Value = 2
    scrlItemIndex.Value = 1
End Sub

Private Sub lstIndex_Click()
    EditorIndex = lstIndex.ListIndex + 1
    LoadConquista
End Sub

Private Sub scrlItem_Change()
    lblItemNumber.Caption = "Número: " & scrlItem.Value
    Conquistas(EditorIndex).Reward(scrlItemIndex.Value).Num = scrlItem.Value
    If scrlItem.Value > 0 Then
        lblItem.Caption = "Item " & Trim$(Item(Conquistas(EditorIndex).Reward(scrlItemIndex.Value).Num).Name)
    Else
        lblItem.Caption = "Item <Nenhum>"
    End If
    txtItem = scrlItem.Value
End Sub

Private Sub scrlItemIndex_Change()
    lblItem.Caption = "Item <Nenhum>"
    If Conquistas(EditorIndex).Reward(scrlItemIndex.Value).Num > 0 Then
        scrlItem.Value = Conquistas(EditorIndex).Reward(scrlItemIndex.Value).Num
        txtQuant = Conquistas(EditorIndex).Reward(scrlItemIndex.Value).Value
    End If
End Sub

Private Sub txtDesc_Change()
    Conquistas(EditorIndex).Desc = txtDesc
End Sub

Private Sub txtExp_Change()
    Conquistas(EditorIndex).Exp = Val(txtExp)
End Sub

Private Sub txtItem_Change()
    scrlItem.Value = Val(txtItem)
End Sub

Private Sub txtNome_Change()
    Conquistas(EditorIndex).Name = txtNome
End Sub

Private Sub txtNome_LostFocus()
    lstIndex.List(EditorIndex - 1) = EditorIndex & ": " & txtNome
End Sub

Private Sub txtProgress_Change()
    Conquistas(EditorIndex).Progress = Val(txtProgress)
End Sub

Private Sub txtQuant_Change()
    Conquistas(EditorIndex).Reward(scrlItemIndex.Value).Value = Val(txtQuant)
End Sub
