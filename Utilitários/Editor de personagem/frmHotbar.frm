VERSION 5.00
Begin VB.Form frmHotbar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotbar"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHotbar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   4575
      Begin VB.CommandButton Command2 
         Caption         =   "Limpar"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salvar"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cmbMagia 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox cmbItem 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.OptionButton optMagia 
         Caption         =   "Magia"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Atalhos"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.ListBox lstAtalhos 
         Height          =   3570
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmHotbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If optItem.value = True Then
        Player.Hotbar(lstAtalhos.ListIndex + 1).Slot = cmbItem.ListIndex + 1
        Player.Hotbar(lstAtalhos.ListIndex + 1).sType = 1
    Else
        Player.Hotbar(lstAtalhos.ListIndex + 1).Slot = cmbMagia.ListIndex + 1
        Player.Hotbar(lstAtalhos.ListIndex + 1).sType = 2
    End If
    Dim Lastindex As Long
    Lastindex = lstAtalhos.ListIndex
    UpdateList
    lstAtalhos.ListIndex = Lastindex
End Sub

Private Sub Command2_Click()
    Player.Hotbar(lstAtalhos.ListIndex + 1).Slot = 0
    Player.Hotbar(lstAtalhos.ListIndex + 1).sType = 0
    Dim Lastindex As Long
    Lastindex = lstAtalhos.ListIndex
    UpdateList
    lstAtalhos.ListIndex = Lastindex
End Sub

Private Sub Form_Load()
    UpdateList
End Sub

Sub UpdateList()
    lstAtalhos.Clear
    
    Dim i As Long, Prefix As String
    For i = 1 To MAX_HOTBAR
        If i <= 9 Then
            Prefix = i & ": "
        End If
        If i = 10 Then Prefix = "0: "
        If i = 11 Then Prefix = "-: "
        If i = 12 Then Prefix = "=: "
        
        If Player.Hotbar(i).sType = 0 Then Prefix = Prefix & "{NADA}"
        
        If Player.Hotbar(i).sType = 1 Then
            If Player.Inv(Player.Hotbar(i).Slot).Num > 0 Then
                Prefix = Prefix & "{ITEM} " & Trim$(Item(Player.Inv(Player.Hotbar(i).Slot).Num).Name)
            Else
                Prefix = Prefix & "{ITEM} Nenhum (BUG!)"
            End If
        End If
        
        If Player.Hotbar(i).sType = 2 Then
            If Player.Spell(Player.Hotbar(i).Slot) > 0 Then
                Prefix = Prefix & "{MAGIA} " & Trim$(Spell(Player.Spell(Player.Hotbar(i).Slot)).Name)
            Else
                Prefix = Prefix & "{MAGIA} Nenhuma (BUG!)"
            End If
        End If
        
        lstAtalhos.AddItem Prefix
    Next i
    
    For i = 1 To MAX_INV
        If Player.Inv(i).Num > 0 Then
            cmbItem.AddItem i & ": " & Trim$(Item(Player.Inv(i).Num).Name)
        Else
            cmbItem.AddItem i & ": {Nenhum}"
        End If
    Next i
    
    For i = 1 To MAX_PLAYER_SPELLS
        If Player.Spell(i) > 0 Then
            cmbMagia.AddItem i & ": " & Trim$(Spell(Player.Spell(i)).Name)
        Else
            cmbMagia.AddItem i & ": {Nenhuma}"
        End If
    Next i
    
End Sub

Private Sub lstAtalhos_Click()
    Frame2.Enabled = True
    
    If Player.Hotbar(lstAtalhos.ListIndex + 1).sType = 0 Then
        optItem.value = False
        optMagia.value = False
        cmbItem.ListIndex = -1
        cmbMagia.ListIndex = -1
        Exit Sub
    End If
    
    If Player.Hotbar(lstAtalhos.ListIndex + 1).sType = 1 Then
        optItem.value = True
        cmbItem.ListIndex = Player.Hotbar(lstAtalhos.ListIndex + 1).Slot - 1
    Else
        optMagia.value = True
        cmbMagia.ListIndex = Player.Hotbar(lstAtalhos.ListIndex + 1).Slot - 1
    End If
End Sub

Private Sub optItem_Click()
    cmbItem.Enabled = True
    cmbMagia.Enabled = False
    cmbMagia.ListIndex = -1
End Sub

Private Sub optMagia_Click()
    cmbItem.Enabled = False
    cmbMagia.Enabled = True
    cmbItem.ListIndex = -1
End Sub
