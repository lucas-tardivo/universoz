VERSION 5.00
Begin VB.Form frmSwitches 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Switches e variaveis"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSwitches.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Conteúdo"
      Height          =   1095
      Left            =   5400
      TabIndex        =   8
      Top             =   8520
      Width           =   5175
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salvar"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Valor:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Variaveis"
      Height          =   8295
      Left            =   5400
      TabIndex        =   6
      Top             =   120
      Width           =   5175
      Begin VB.ListBox lstVariables 
         Height          =   7860
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Conteúdo"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   8520
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "Salvar"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmbValue 
         Height          =   315
         ItemData        =   "frmSwitches.frx":2982
         Left            =   720
         List            =   "frmSwitches.frx":298C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Valor:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Switches"
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ListBox lstSwitches 
         Height          =   7860
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmSwitches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim lastindex As Long
    Player.Switches(lstSwitches.ListIndex + 1) = cmbValue.ListIndex
    lastindex = lstSwitches.ListIndex
    UpdateSwitches
    lstSwitches.ListIndex = lastindex
End Sub

Private Sub Command2_Click()
    Dim lastindex As Long
    Player.Variables(lstVariables.ListIndex + 1) = txtValue.Text
    lastindex = lstVariables.ListIndex
    UpdateVariables
    lstVariables.ListIndex = lastindex
End Sub

Private Sub Form_Load()
    UpdateSwitches
    UpdateVariables
End Sub

Sub UpdateSwitches()
    lstSwitches.Clear
    
    Dim i As Long, Switch As String
    For i = 1 To MAX_SWITCHES
        If Switches(i) <> "" Then
            Switch = "Falso"
            If Player.Switches(i) = 1 Then Switch = "Verdadeiro"
            lstSwitches.AddItem i & ": " & Switches(i) & " (" & Switch & ")"
        Else
            lstSwitches.AddItem i & ": {Não editada}"
        End If
    Next i
    
End Sub

Sub UpdateVariables()
    lstVariables.Clear
    
    Dim i As Long
    For i = 1 To MAX_VARIABLES
        If Variables(i) <> "" Then
            lstVariables.AddItem i & ": " & Variables(i) & " (" & Player.Variables(i) & ")"
        Else
            lstVariables.AddItem i & ": {Não editada}"
        End If
    Next i
    
End Sub

Private Sub lstSwitches_Click()
    cmbValue.ListIndex = Player.Switches(lstSwitches.ListIndex + 1)
End Sub

Private Sub lstVariables_Click()
    txtValue.Text = Player.Variables(lstVariables.ListIndex + 1)
End Sub
