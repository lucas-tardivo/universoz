VERSION 5.00
Begin VB.Form frmInv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventário"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmLoad 
      Height          =   4095
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   4575
      Begin VB.Label Label8 
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
         TabIndex        =   19
         Top             =   1920
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Equipamentos"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   4575
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   3
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1440
         Width           =   3615
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   2
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1080
         Width           =   3375
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   1
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox cmbEquip 
         Height          =   315
         Index           =   0
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label7 
         Caption         =   "Botas:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Calças:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Camisa:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Arma:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bolsa"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cmbInvNum 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   3855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Item"
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   4335
         Begin VB.ComboBox cmbItem 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txtQuant 
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Salvar"
            Height          =   255
            Left            =   2760
            TabIndex        =   2
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Numero:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Quantidade:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Slot:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub UpdateInvList()
    Dim i As Long
    
    cmbInvNum.Clear
    For i = 1 To MAX_INV
        If Player.Inv(i).Num > 0 Then
            cmbInvNum.AddItem i & ": " & Trim$(Item(Player.Inv(i).Num).Name)
        Else
            cmbInvNum.AddItem i & ": {Nenhum}"
        End If
    Next i
End Sub

Private Sub cmbEquip_Click(Index As Integer)

    Dim ItemNum As Long, CanEquip As Boolean
    ItemNum = cmbEquip(Index).ListIndex
    
    If ItemNum > 0 Then
        CanEquip = False
        If Index = 0 And Item(ItemNum).Type = ItemType.ITEM_TYPE_WEAPON Then CanEquip = True
        If Index = 1 And Item(ItemNum).Type = ItemType.ITEM_TYPE_ARMOR Then CanEquip = True
        If Index = 2 And Item(ItemNum).Type = ItemType.ITEM_TYPE_HELMET Then CanEquip = True
        If Index = 3 And Item(ItemNum).Type = ItemType.ITEM_TYPE_SHIELD Then CanEquip = True
    End If
    
    If CanEquip = True Or ItemNum = 0 Then
        Player.Equipment(Index + 1) = cmbEquip(Index).ListIndex
    Else
        cmbEquip(Index).ListIndex = Player.Equipment(Index + 1)
        MsgBox "Esse item não é compatível com esse slot!", vbCritical
    End If
End Sub

Private Sub cmbInvNum_Click()
    On Error Resume Next
    cmbItem.ListIndex = Player.Inv(cmbInvNum.ListIndex + 1).Num
    txtQuant.Text = Player.Inv(cmbInvNum.ListIndex + 1).value
End Sub

Private Sub cmbWeapon_Change()

End Sub

Private Sub cmbWeapon_Click()
    
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim LastEdit As Long
    Player.Inv(cmbInvNum.ListIndex + 1).Num = cmbItem.ListIndex
    Player.Inv(cmbInvNum.ListIndex + 1).value = Val(txtQuant.Text)
    LastEdit = cmbInvNum.ListIndex
    UpdateInvList
    cmbInvNum.ListIndex = LastEdit
End Sub

Private Sub Form_Load()
    frmInv.Visible = True
    frmInv.Enabled = False
    UpdateInvList
    Dim i As Long, n As Long
    cmbItem.AddItem "{Nenhum}"
    
    For n = 1 To Equipment.Equipment_Count - 1
        cmbEquip(n - 1).AddItem "{Nenhum}"
    Next n
    
    For i = 1 To MAX_ITEMS
        cmbItem.AddItem i & ": " & Trim$(Item(i).Name)
        For n = 1 To Equipment.Equipment_Count - 1
            cmbEquip(n - 1).AddItem i & ": " & Trim$(Item(i).Name)
        Next n
        DoEvents
    Next i
    
    For n = 1 To Equipment.Equipment_Count - 1
        cmbEquip(n - 1).ListIndex = Player.Equipment(n)
    Next n
    
    frmInv.Enabled = True
    frmLoad.Visible = False
    cmbInvNum.ListIndex = 0
End Sub

Private Sub Frame4_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
