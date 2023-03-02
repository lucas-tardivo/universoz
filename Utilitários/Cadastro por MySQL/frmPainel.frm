VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPainel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GoPlay Games - Web Manager"
   ClientHeight    =   4935
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   12120
   Icon            =   "frmPainel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock socket 
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picManual 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   2880
      ScaleHeight     =   915
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "Register"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtPass 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtLogin 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Login:"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Timer tmrCadastro 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   12135
   End
   Begin VB.Menu mnuCommands 
      Caption         =   "Comandos"
      Begin VB.Menu mnuManual 
         Caption         =   "Add manual"
      End
      Begin VB.Menu btnConnect 
         Caption         =   "Conectar ao servidor"
      End
      Begin VB.Menu btnSincronize 
         Caption         =   "Sincronizar Ranks"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmPainel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnConnect_Click()
    If ConnectToServer Then
        AddLog "Conexão efetuada com sucesso!"
    Else
        AddLog "Falha na conexão!"
    End If
End Sub

Private Sub btnSincronize_Click()
    Dim sFilename As String
    Dim SQL As String
    sFilename = Dir(App.Path & "\data\accounts\")

    'Utilizado para sincronizar ranking no site
    AddLog "Obtendo informações de contas..."
    Do While sFilename > ""
    
        LoadPlayer sFilename
        If sFilename <> "charlist.bin" Then
        If Trim$(Player.Login) <> vbNullString Then
            SQL = "UPDATE users SET level = " & Player.Level & ", stars = " & Player.TopStars & " WHERE login = '" & Trim$(Player.Login) & "'"
            On Error Resume Next
            gConexao.Execute SQL
        End If
        End If
      
      
      sFilename = Dir()
    
    Loop
    AddLog "Todas as contas foram sincronizadas!"
End Sub

Private Sub Command1_Click()
    Call AddAccount(txtLogin, txtPass)
    Call AddLog("Conta " & txtLogin & " registrada!")
    picManual.Visible = False
End Sub

Private Sub Form_Load()
On Error GoTo Errorhandler
Call AddLog("Conectando...")
Call AddLog("Logando em " & GetVar(App.Path & "\WebManager.ini", "CONFIG", "Server") & " with user " & GetVar(App.Path & "\WebManager.ini", "CONFIG", "User") & " in database " & GetVar(App.Path & "\WebManager.ini", "CONFIG", "database"))

    If ConnectDatabase Then
        Call AddLog("Logado com sucesso!")
        Call AddLog("Os próximos registros serão adicionados no banco")
        tmrCadastro.Enabled = True
    Else
        Call AddLog("Não foi possível conectar com o MySQL")
    End If
    
    AddLog "Iniciando componente de comunicação com o servidor..."
    If ConnectToServer Then
        AddLog "Componente iniciado com sucesso!"
    Else
        AddLog "Não foi possível efetuar conexão com o servidor!"
    End If
    Exit Sub
Errorhandler:
    Select Case Err.Number
        Case -2147467259
        Call AddLog("A problem with driver occurs! Please install MySQL ODBC 5.1 Driver 86x")
        Exit Sub
    End Select

    Call AddLog("Unknown connection error: [" & Err.Number & "] " & Err.Description)
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuManual_Click()
    picManual.Visible = Not picManual.Visible
End Sub

Private Sub socket_Close()
    socket.Close
End Sub

Private Sub socket_DataArrival(ByVal bytesTotal As Long)
    Dim packet As String
    socket.GetData packet, vbString, bytesTotal
    HandleData packet
End Sub

Private Sub tmrCadastro_Timer()
    Dim pRsAdd As ADODB.Recordset, pRsRemove As ADODB.Recordset
    
    On Error GoTo errhandler
    
    SQL = "SELECT id,login,senha FROM " & GetVar(App.Path & "\WebManager.ini", "CONFIG", "table")
    Set pRsAdd = gConexao.Execute(SQL)
    
    Do While Not pRsAdd.EOF
        If Not FileExist(App.Path & "\data\accounts\" & pRsAdd.Fields("login").Value & ".bin") Then
            Call AddAccount(pRsAdd.Fields("login").Value, pRsAdd.Fields("senha").Value)
            Call AddLog("Conta " & pRsAdd.Fields("login").Value & " registrada!")
        Else
            Call AddLog("Conta " & pRsAdd.Fields("login").Value & " já existente!")
        End If
        SQL = "DELETE FROM " & GetVar(App.Path & "\WebManager.ini", "CONFIG", "table") & " WHERE id=" & pRsAdd.Fields("id").Value & ";"
        Set pRsRemove = gConexao.Execute(SQL)
        pRsAdd.MoveNext
        DoEvents
    Loop
    
    Set pRsAdd = Nothing
    
    If frmPainel.socket.State = sckConnected Then
        SQL = "SELECT bolsa.id, T1.login, T2.realitem FROM bolsa INNER JOIN users T1 ON T1.id = bolsa.conta INNER JOIN shop T2 ON T2.id = bolsa.item AND bolsa.recebido = 0"
        Set pRsAdd = gConexao.Execute(SQL)
        
        Do While Not pRsAdd.EOF
            
            AddLog "[SHOP] " & pRsAdd.Fields("login").Value & " está sendo tratado com o item " & pRsAdd.Fields("realitem").Value & " requisição enviada para o servidor!"
            addItemLog "[WEBSERVER] " & pRsAdd.Fields("login").Value & " está sendo tratado com o item " & pRsAdd.Fields("realitem").Value & " requisição enviada para o servidor!"
            SQL = "UPDATE bolsa SET recebido=1 WHERE ID = " & pRsAdd.Fields("ID").Value & ";"
            Set pRsRemove = gConexao.Execute(SQL)
            
            SendData "additem;" & pRsAdd.Fields("login").Value & ";" & pRsAdd.Fields("realitem").Value & ";" & pRsAdd.Fields("id").Value
            
            pRsAdd.MoveNext
            DoEvents
        Loop
    End If
    
    Exit Sub
errhandler:
    Shell App.Path & "\WebManager.exe"
    Unload Me
    gConexao.Close
    If Not ConnectDatabase Then AddLog "Falha na tentativa ao reconectar com o banco de dados!"
End Sub

Private Sub tmrRank_Timer()

End Sub

