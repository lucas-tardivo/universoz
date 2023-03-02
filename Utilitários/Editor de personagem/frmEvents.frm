VERSION 5.00
Begin VB.Form frmEvents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eventos abertos"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEvents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Eventos"
      Height          =   8295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5175
      Begin VB.ListBox lstEvents 
         Height          =   7860
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Conteúdo"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   8520
      Width           =   5175
      Begin VB.ComboBox cmbValue 
         Height          =   315
         ItemData        =   "frmEvents.frx":2982
         Left            =   720
         List            =   "frmEvents.frx":298C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salvar"
         Height          =   255
         Left            =   3480
         TabIndex        =   1
         Top             =   720
         Width           =   1575
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
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Lastindex As Long
    Player.EventOpen(lstEvents.ListIndex + 1) = cmbValue.ListIndex
    Lastindex = lstEvents.ListIndex
    UpdateEventList
    lstEvents.ListIndex = Lastindex
End Sub

Private Sub Form_Load()
    UpdateEventList
End Sub

Sub UpdateEventList()
    lstEvents.Clear
    
    Dim i As Long, EventState As String
    
    For i = 1 To MAX_EVENTS
        EventState = "Fechado"
        If Player.EventOpen(i) = 1 Then EventState = "Aberto"
        
        If Trim$(Events(i).Name) <> "" Then
            lstEvents.AddItem i & ": " & Trim$(Events(i).Name) & " (" & EventState & ")"
        Else
            lstEvents.AddItem i & ": {Não editado}"
        End If
    Next i
End Sub

Private Sub lstEvents_Click()
    cmbValue.ListIndex = Player.EventOpen(lstEvents.ListIndex + 1)
End Sub
