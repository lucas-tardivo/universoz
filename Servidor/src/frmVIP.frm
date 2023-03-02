VERSION 5.00
Begin VB.Form frmVIP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add VIP"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   315
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkGlobal 
      Caption         =   "Avisar no global"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox cmbVIP 
      Height          =   315
      ItemData        =   "frmVIP.frx":0000
      Left            =   3120
      List            =   "frmVIP.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtDias 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "30"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtVipData 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtIndex 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Type:"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Days:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Init date:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Index:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmVIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Index As Long
    Index = txtIndex
    
    If IsPlaying(Index) Then
        If Player(Index).VIP = 0 Then Player(Index).VIP = cmbVIP.ListIndex + 1
        Player(Index).VIPData = txtVipData.Text
        Player(Index).VIPDias = Val(txtDias.Text)
        SavePlayer Index
        SendPlayerData Index
        'If chkGlobal.Value = 0 Then
        '    PlayerMsg Index, "Parabéns! Você recebeu " & Player(Index).VIPDias & " dias VIP!", yellow
        'Else
        '    GlobalMsg "Parabéns " & GetPlayerName(Index) & "! Você recebeu " & Player(Index).VIPDias & " dias VIP!", yellow
        'End If
    End If
    
    frmVIP.Visible = False
End Sub

Private Sub Form_Load()
    txtVipData.Text = Date
    cmbVIP.ListIndex = 0
End Sub
