Attribute VB_Name = "modDatabase"
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long


Sub LoadPlayer(ByVal Name As String)
    Dim filename As String
    Dim f As Long
    Call ClearPlayer
    filename = App.Path & "\data\accounts\" & Trim(Name) & ".bin"
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Player
    Close #f
End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim f As Long

    For i = 1 To MAX_SPELLS
        filename = App.Path & "\data\spells\spells" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Spell(i)
        Close #f
    Next

End Sub

Sub ClearPlayer()
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(Player), LenB(Player))
    Player.Login = vbNullString
    Player.Password = vbNullString
    Player.Name = vbNullString
    Player.Class = 1
End Sub

Sub SavePlayer(ByVal Name As String)
    Dim filename As String
    Dim f As Long

    filename = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
    
    f = FreeFile
    
    Player.Name = Player.Name & " "
    
    Open filename For Binary As #f
    Put #f, , Player
    Close #f
End Sub





Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim f As Long

    For i = 1 To MAX_ITEMS
        filename = App.Path & "\data\Items\Item" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Item(i)
        Close #f
        'frmEditor.Caption = "Carregando itens " & Int(i / MAX_ITEMS * 100) & "%..."
        DoEvents
    Next

End Sub

Sub LoadSwitches()
Dim i As Long, filename As String
    filename = App.Path & "\data\switches.ini"
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = GetVar(filename, "Switches", "Switch" & CStr(i) & "Name")
        DoEvents
    Next
End Sub

Sub LoadVariables()
Dim i As Long, filename As String
    filename = App.Path & "\data\variables.ini"
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = GetVar(filename, "Variables", "Variable" & CStr(i) & "Name")
        DoEvents
    Next
End Sub

Public Sub LoadEvents()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Call LoadEvent(i)
        DoEvents
    Next i
End Sub

Public Sub LoadEvent(ByVal Index As Long)
    Dim f As Long, SCount As Long, s As Long, DCount As Long, D As Long
    Dim filename As String
    filename = App.Path & "\data\events\event" & Index & ".dat"
    If FileExist(filename, True) Then
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , Events(Index).Name
        Close #f
    End If
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub
