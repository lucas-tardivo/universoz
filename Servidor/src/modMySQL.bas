Attribute VB_Name = "modMySQL"
Public gConexaoDatabase As ADODB.Connection

Public Function ConnectDatabase() As Boolean
On Error GoTo Errorhandler
    Dim filename As String
    filename = App.path & "\mysql.ini"
    Dim CONFIG As String
    CONFIG = GetVar(filename, "GENERAL", "SQL")

    Set gConexaoDatabase = New ADODB.Connection
    gConexaoDatabase.ConnectionTimeout = 1000
    gConexaoDatabase.CommandTimeout = 1000
    gConexaoDatabase.CursorLocation = adUseClient
    gConexaoDatabase.Open "DRIVER={MySQL ODBC 5.1 Driver};" _
        & "user=" & GetVar(filename, "SQL" & CONFIG, "user") _
        & ";password=" & GetVar(filename, "SQL" & CONFIG, "pass") _
        & ";database=" & GetVar(filename, "SQL" & CONFIG, "database") _
        & ";server=" & GetVar(filename, "SQL" & CONFIG, "host") _
        & ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        
    If gConexaoDatabase.State = 1 Then
        ConnectDatabase = True
    Else
        ConnectDatabase = False
    End If
    Exit Function
Errorhandler:
    Select Case Err.Number
        Case -2147467259
        ConnectDatabase = False
        Call SetStatus("Conexão com MySQL mal sucedida, verifique a instalação do driver corretamente e as credenciais de conexão " & vbNewLine & "[" & Err.Description & "]")
        Exit Function
    End Select

    Call SetStatus("Erro desconhecido: [" & Err.Number & "] " & Err.Description)
    ConnectDatabase = False
End Function

Public Sub Query(ByRef RecordSet As ADODB.RecordSet, SQL As String)
    On Error GoTo Retry
    If gConexaoDatabase.State = 1 Then
        Set RecordSet = gConexaoDatabase.Execute(SQL)
    Else
        If ConnectSQL Then
            Set RecordSet = gConexaoDatabase.Execute(SQL)
        Else
            MsgBox "Erro de conexão com o banco de dados"
            Call DestroyServer
        End If
    End If
    Exit Sub
Retry:
    MsgBox "Erro ao conectar com o banco de dados: " & Err.Description
        'Query RecordSet, SQL
    Call DestroyServer
End Sub

Public Sub Update(SQL As String)
    gConexaoDatabase.Execute SQL
End Sub

Sub WebUpdate(Data As String)
    If frmServer.sckWeb.State = sckConnected Then
        frmServer.sckWeb.SendData Data
    End If
End Sub

Sub UpdateWebRank(ByVal Login As String, ByVal Rank As String, ByVal Value As Long)
    WebUpdate "updaterank;" & Login & ";" & Rank & ";" & Value
End Sub

Sub WebConfirmReception(ByVal Login As String, ByVal ItemNum As Long, ByVal CompraID As Long)
    WebUpdate "confirm;" & Login & ";" & ItemNum & ";" & CompraID
End Sub

Sub WebPlayerNotOnline(ByVal Login As String, ByVal ItemNum As Long, ByVal CompraID As Long)
    WebUpdate "notonline;" & Login & ";" & ItemNum & ";" & CompraID
End Sub

Sub CloseWebManager()
    WebUpdate "close"
End Sub

Sub HandleWebData(ByVal Data As String)
    Dim Parse() As String
    Dim User As String
    Dim Rank As String
    Dim Value As Long
    Dim SQL As String
    If Len(Data) > 0 Then
        Parse = Split(Data, ";")
        
        If Parse(0) = "additem" Then
            Dim i As Long
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If Trim$(LCase(Player(i).Login)) = Trim$(LCase(Parse(1))) Then
                        GiveInvItem i, Val(Parse(2)), 1, True
                        PlayerMsg i, "Parabéns! Você recebeu um " & Trim$(Item(Val(Parse(2))).Name) & " da loja!" & "!", brightgreen
                        WebConfirmReception Parse(1), Val(Parse(2)), Val(Parse(3))
                        addItemLog "(SERVER) Jogador " & GetPlayerName(i) & " recebeu o item " & Trim$(Item(Val(Parse(2))).Name)
                        Exit Sub
                    End If
                End If
            Next i
            WebPlayerNotOnline Parse(1), Val(Parse(2)), Val(Parse(3))
        End If
        
    End If
End Sub
