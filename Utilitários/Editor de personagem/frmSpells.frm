VERSION 5.00
Begin VB.Form frmSpells 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Magias"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpells.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Modificar"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Modificar"
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox cmbMagia 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Magia:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Magias"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.ListBox lstMagias 
         Height          =   2400
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmSpells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub UpdateSpellList()
    Dim i As Long
    lstMagias.Clear
    For i = 1 To MAX_PLAYER_SPELLS
        If Player.Spell(i) > 0 Then
            lstMagias.AddItem i & ": " & Trim$(Spell(Player.Spell(i)).Name)
        Else
            lstMagias.AddItem i & ": {Nenhuma}"
        End If
    Next i
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Player.Spell(lstMagias.ListIndex + 1) = cmbMagia.ListIndex
    UpdateSpellList
End Sub

Private Sub Form_Load()
    Dim i As Long
    cmbMagia.AddItem "{Nenhuma}"
    For i = 1 To MAX_SPELLS
        If Trim$(Spell(i).Name) <> "" Then
            cmbMagia.AddItem i & ": " & Trim$(Spell(i).Name)
        Else
            cmbMagia.AddItem i & ": {Não editada}"
        End If
    Next i
    UpdateSpellList
End Sub

Private Sub lstMagias_Click()
    cmbMagia.ListIndex = Player.Spell(lstMagias.ListIndex + 1)
End Sub
