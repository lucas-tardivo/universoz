VERSION 5.00
Begin VB.Form frmSuporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suporte"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7230
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
   ScaleHeight     =   3135
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstPlayers 
      Height          =   2205
      Left            =   4680
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox txtMsg 
      Height          =   285
      Left            =   120
      MaxLength       =   255
      TabIndex        =   1
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox txtChat 
      Height          =   2655
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Jogadores:"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmSuporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim i As Long
    
    If GetPlayerAccess(MyIndex) > 0 Then
        Me.Width = 7305
        lstPlayers.Clear
        For i = 1 To 20
            lstPlayers.AddItem "<Slot para chat>"
            Load txtChat(i)
        Next i
    Else
        Me.Width = 4650
        txtChat(0).visible = True
    End If
    
End Sub

Private Sub lstPlayers_Click()
    Dim i As Long
    For i = 1 To 20
        txtChat(i).visible = (i = lstPlayers.ListIndex + 1)
    Next i
    If SupportNames(lstPlayers.ListIndex + 1) <> vbNullString Then
        lstPlayers.List(lstPlayers.ListIndex) = SupportNames(lstPlayers.ListIndex + 1)
    End If
End Sub

Private Sub txtMsg_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txtMsg) > 0 Then
            If GetPlayerAccess(MyIndex) = 0 Then
                SendSupportMsg txtMsg
                txtChat(0).Text = txtChat(0).Text & vbNewLine & GetPlayerName(MyIndex) & ": " & txtMsg
                txtChat(0).SelStart = Len(txtChat(0).Text)
            Else
                SendSupportMsg txtMsg, SupportNames(lstPlayers.ListIndex + 1)
                txtChat(lstPlayers.ListIndex + 1).Text = txtChat(lstPlayers.ListIndex + 1).Text & vbNewLine & GetPlayerName(MyIndex) & ": " & txtMsg
                txtChat(lstPlayers.ListIndex + 1).SelStart = Len(txtChat(lstPlayers.ListIndex + 1).Text)
            End If
            
            txtMsg = vbNullString
        End If
    End If
End Sub
