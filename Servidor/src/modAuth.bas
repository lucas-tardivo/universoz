Attribute VB_Name = "modAuth"
' Text API
Declare Function GeneralWinDirApi Lib "kernel32" _
        Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
        ByVal nSize As Long) As Long
        
Public gConexao As ADODB.Connection
Public PrivateMode As Boolean
        
Function WinDir() As String
    Const FIX_LENGTH% = 4096
    Dim Length As Integer
    Dim Buffer As String * FIX_LENGTH

    Length = GeneralWinDirApi(Buffer, FIX_LENGTH - 1)
    WinDir = Left$(Buffer, Length)
End Function

Public Function ConnectSQL() As Boolean
On Error GoTo Errorhandler
    Set gConexao = New ADODB.Connection
    gConexao.ConnectionTimeout = 60
    gConexao.CommandTimeout = 400
    gConexao.CursorLocation = adUseClient
    gConexao.Open "DRIVER={MySQL ODBC 5.1 Driver};" _
        & "user=goplaygame_1" _
        & ";password=" & Chr(102) & Chr(97) & Chr(50) & Chr(51) & Chr(49) & Chr(55) & Chr(49) & Chr(49) _
        & ";database=" _
        & ";server=" _
        & ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        
    If gConexao.State = 1 Then
        ConnectSQL = True
    Else
        ConnectSQL = False
        Call SetStatus("The connection doesnt have success, check if you have internet connection and GoPlay servers disponibility. The server is starting on private mode")
    End If
    Exit Function
Errorhandler:
    Select Case Err.Number
        Case -2147467259
        ConnectSQL = False
        Call SetStatus("The connection doesnt have success, check if MySQL ODBC 5.1 86x is installed correctly and if you have internet connection. The server is starting on private mode")
        Exit Function
    End Select

    Call SetStatus("Unrecognized error: [" & Err.Number & "] " & Err.Description)
    ConnectSQL = False
End Function

Public Sub DoRedims()
    ReDim Player(1 To MAX_PLAYERS) As PlayerRec
    ReDim Bank(1 To MAX_PLAYERS) As BankRec
    ReDim TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
End Sub
