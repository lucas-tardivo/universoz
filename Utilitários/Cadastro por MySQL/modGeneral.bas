Attribute VB_Name = "modGeneral"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public gConexao As ADODB.Connection
Public SQL As String

Sub AddLog(ByVal Text As String)
    If frmPainel.txtLog <> "" Then
        frmPainel.txtLog.Text = frmPainel.txtLog.Text & vbNewLine & Text
    Else
        frmPainel.txtLog.Text = Text
    End If
    frmPainel.txtLog.SelStart = Len(frmPainel.txtLog.Text) - 1
End Sub

Public Function ConnectDatabase() As Boolean
On Error GoTo Errorhandler
    Set gConexao = Nothing
    Set gConexao = New ADODB.Connection
    gConexao.ConnectionTimeout = 60
    gConexao.CommandTimeout = 400
    gConexao.CursorLocation = adUseClient
    gConexao.Open "DRIVER={MySQL ODBC 5.1 Driver};" _
        & "user=" & GetVar(App.Path & "\WebManager.ini", "CONFIG", "User") _
        & ";password=" & GetVar(App.Path & "\WebManager.ini", "CONFIG", "Password") _
        & ";database=" & GetVar(App.Path & "\WebManager.ini", "CONFIG", "database") _
        & ";server=" & GetVar(App.Path & "\WebManager.ini", "CONFIG", "Server") _
        & ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        
    If gConexao.State = 1 Then
        ConnectDatabase = True
    Else
        ConnectDatabase = False
    End If
    Exit Function
Errorhandler:
    Select Case Err.Number
        Case -2147467259
        ConnectDatabase = False
        Call AddLog("Conexão com MySQL mal sucedida, verifique a instalação do driver corretamente e as credenciais de conexão " & vbNewLine & "[" & Err.Description & "]")
        Exit Function
    End Select

    Call AddLog("Erro desconhecido: [" & Err.Number & "] " & Err.Description)
    ConnectDatabase = False
End Function

Sub HandleData(ByVal Data As String)
    Dim Parse() As String
    Dim User As String
    Dim Rank As String
    Dim Value As Long
    Dim ItemNum As Long
    Dim ID As Long
    Dim SQL As String
    If Len(Data) > 0 Then
        Parse = Split(Data, ";")
        
        If Parse(0) = "updaterank" Then
            User = Parse(1)
            Rank = Parse(2)
            Value = Val(Parse(3))
            If gConexao.State = 1 Then
                SQL = "UPDATE users SET " & Rank & " = " & Value & " WHERE login = '" & User & "'"
                gConexao.Execute SQL
                AddLog "Usuário " & User & " atualizado no ranking " & Rank & " com novo valor " & Value
            Else
                AddLog "Pedido de atualização de ranking negado por falta de conexão com o banco!"
            End If
        End If
        
        If Parse(0) = "close" Then
            End
        End If
        
        If Parse(0) = "confirm" Then
            User = Parse(1)
            ItemNum = Val(Parse(2))
            ID = Val(Parse(3))
            'Recebimento feito por jogador online
            AddLog "[SHOP] Recebimento de item de " & User & " confirmada pelo servidor!"
            addItemLog "[WEBSERVER] Confirmado o recebimento do item " & ItemNum & " de " & User
            
            If gConexao.State = 1 Then
                SQL = "UPDATE bolsa SET recebido = 2 WHERE id = '" & ID & "'"
                gConexao.Execute SQL
                AddLog "Usuário " & User & " teve seu item atualizado no shop!"
            Else
                AddLog "Pedido de atualização de item no shop negado por falta de conexão com o banco!"
                addItemLog "[WEBSERVER] FALHA Pedido de " & User & " item " & ItemNum & " foi entregue porem nao foi possivel atualizar o site!"
            End If
        End If
        
        If Parse(0) = "notonline" Then
            User = Parse(1)
            ItemNum = Val(Parse(2))
            ID = Val(Parse(3))
            GiveInvItem User, ItemNum, ID
        End If
    End If
End Sub
Function ConnectToServer() As Boolean
    frmPainel.socket.Close
    frmPainel.socket.RemoteHost = "localhost"
    frmPainel.socket.RemotePort = 7499
    frmPainel.socket.Connect
    
    Dim Tick As Long
    Tick = GetTickCount + 1000
    
    Do While frmPainel.socket.State <> sckConnected And Tick > GetTickCount
        DoEvents
    Loop
    
    If frmPainel.socket.State = sckConnected Then
        ConnectToServer = True
    Else
        ConnectToServer = False
    End If
End Function
Sub SendData(ByVal Data As String)
    If frmPainel.socket.State = sckConnected Then
        frmPainel.socket.SendData Data
    End If
End Sub
