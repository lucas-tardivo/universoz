VERSION 5.00
Begin VB.Form frmVIP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VIP"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVIP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtDias 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox cmbVIP 
      Height          =   315
      ItemData        =   "frmVIP.frx":2982
      Left            =   1200
      List            =   "frmVIP.frx":298C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Início:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Dias:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de vip:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmVIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Player.VIPData = txtData.Text
Player.VIP = cmbVIP.ListIndex
Player.VIPDias = txtDias.Text

    Dim PlayerDir As String
    PlayerDir = "Nenhum"
    If Player.VIP = 1 Then PlayerDir = "Comum"
    
    With frmEditor
    .lblVIP.Caption = "Plano VIP: " & PlayerDir
    .lblDias.Caption = "Dias: " & Player.VIPDias
    .lblInicio.Caption = "Inicio: " & Player.VIPData
    If Player.VIP = 1 Then
        .lblRestantes.Caption = "Dias restantes: " & (Player.VIPDias - DateDiff("d", Player.VIPData, Date))
    Else
        .lblRestantes.Caption = "Dias restantes: "
    End If
    End With
    
Unload Me
End Sub

Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()

If Player.VIPData = "" Then
    txtData.Text = Date
Else
    txtData.Text = Player.VIPData
End If

txtDias.Text = Player.VIPDias
cmbVIP.ListIndex = Player.VIP
End Sub
