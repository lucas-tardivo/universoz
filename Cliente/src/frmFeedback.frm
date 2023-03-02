VERSION 5.00
Begin VB.Form frmFeedback 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nos dê a sua opinião!"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtFeedback 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5175
   End
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmFeedback.frx":0000
      Left            =   600
      List            =   "frmFeedback.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   $"frmFeedback.frx":0045
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmFeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim buffer As clsBuffer
    If Len(txtFeedback) > 3 Then
        If cmbType.ListIndex >= 0 Then
            Set buffer = New clsBuffer
                buffer.WriteLong CFeedback
                buffer.WriteLong cmbType.ListIndex
                buffer.WriteString txtFeedback.Text
                SendData buffer.ToArray
            Set buffer = Nothing
            AddText "Agradecemos muito por sua mensagem! Mantenha-nos sempre informados para que possamos trazer a melhor experiência de jogo para você e seus amigos!", White
            Unload frmFeedback
        Else
            MsgBox "Por favor selecione um tipo de feedback"
        End If
    Else
        MsgBox "Por favor digite uma mensagem"
    End If
End Sub
