Attribute VB_Name = "modDatabase"
Public Player As PlayerRec

Sub LoadPlayer(ByVal name As String)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\accounts\" & Trim(name)
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Player
    Close #f
End Sub

Sub LoadMyPlayer(ByVal name As String, ByRef MyPlayer As PlayerRec)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\accounts\" & Trim(name) & ".bin"
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , MyPlayer
    Close #f
End Sub

Sub AddAccount(ByVal name As String, ByVal Password As String)
    Player.Login = name
    Player.Password = Password
    Player.name = ""

    Call SavePlayer
End Sub

Sub SavePlayer()
    Dim filename As String
    Dim f As Long

    filename = App.Path & "\data\accounts\" & Trim$(Player.Login) & ".bin"
    
    f = FreeFile
    
    Open filename For Binary As #f
    Put #f, , Player
    Close #f
End Sub

Sub SaveMyPlayer(ByRef MyPlayer As PlayerRec)
    Dim filename As String
    Dim f As Long

    filename = App.Path & "\data\accounts\" & Trim$(MyPlayer.Login) & ".bin"
    
    f = FreeFile
    
    Open filename For Binary As #f
    Put #f, , MyPlayer
    Close #f
End Sub

Public Sub addItemLog(ByVal Mensagem As String)
Dim filename As String
    filename = App.Path & "\data\logs\shopitem.txt"
    Open filename For Append As #1
        Print #1, "[" & Now & "] " & Mensagem
        Print #1, ""
    Close #1
End Sub

Sub GiveInvItem(User As String, ItemNum As Long, ID As Long)
            'Jogador offline, entregar o item
            AddLog "[SHOP] Jogador " & User & " offline, entregando item!"
            addItemLog "[WEBSERVER] " & User & " está offline, item " & ItemNum & " será entregado pelo WebServer"
            
            Dim MyPlayer As PlayerRec
            LoadMyPlayer User, MyPlayer
            
            Dim i As Long
            Dim InvFull As Boolean
            InvFull = True
            For i = 1 To MAX_INV
                If MyPlayer.Inv(i).Num = 0 Then
                    MyPlayer.Inv(i).Num = ItemNum
                    MyPlayer.Inv(i).Value = 1
                    InvFull = False
                    Exit For
                End If
            Next i
            
            SaveMyPlayer MyPlayer
            
            If InvFull Then
                If gConexao.State = 1 Then
                    SQL = "UPDATE bolsa SET recebido = 3 WHERE id = '" & ID & "'"
                    gConexao.Execute SQL
                    AddLog "Usuário " & User & " esta com o inventário cheio!"
                Else
                    AddLog "Pedido de atualização de item no shop negado por falta de conexão com o banco!"
                End If
                addItemLog "[WEBSERVER] FALHA Pedido de " & User & " item " & ItemNum & " não foi entregue pelo WebServer pelo inventario estar cheio!"
            Else
                If gConexao.State = 1 Then
                    SQL = "UPDATE bolsa SET recebido = 2 WHERE ID = " & ID
                    gConexao.Execute SQL
                    AddLog "Usuário " & User & " teve seu item atualizado no shop!"
                    addItemLog "[WEBSERVER] Pedido de " & User & " item " & ItemNum & " foi entregue pelo WebServer"
                Else
                    AddLog "Pedido de atualização de item no shop negado por falta de conexão com o banco!"
                    addItemLog "[WEBSERVER] FALHA Pedido de " & User & " item " & ItemNum & " foi entregue pelo WebServer e não foi atualizado no site"
                End If
            End If
End Sub
