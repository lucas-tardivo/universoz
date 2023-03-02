Attribute VB_Name = "modEspacial"
Public Const UZ As Boolean = False

Public Const VirgoMap As Long = 40

Public MAX_PLANETS As Long
Public MAX_PLAYER_PLANETS As Long
Public Planets() As PlanetRec
Public PlanetMoons() As MoonDataRec
Public PlayerPlanet() As PlayerPlanetRec
Public PlayerPlanetMoons() As MoonDataRec
Public MapSaibamans(1 To MAX_MAPS) As MapSaibamansRec

Type MoonDataRec
    Tick As Long
    Position As Double
    Local As Byte
End Type

Type MoonRec
    Size As Long
    ColorR As Long
    ColorG As Long
    ColorB As Long
    Pic As Long
    Speed As Long
End Type

Type TileConfig
    X As Long
    Y As Long
    Tileset As Long
    Layer As Long
End Type

Type PlanetConfig
    Tile(1 To 3) As TileConfig
End Type

Type PlanetRec
    Name As String * NAME_LENGTH
    Map As Long
    Owner As String * NAME_LENGTH
    State As Byte
    PointsToConquest As Long
    WaveDuration As Long
    WaveCooldown As Long
    
    EspeciariaAmarela As Byte
    EspeciariaVermelha As Byte
    EspeciariaAzul As Byte
    
    Level As Long
    Especie As Byte
    Habitantes As Long
    Gravidade As Long
    Atmosfera As Long
    Preco As Long
    Pic As Long
    X As Long
    Y As Long
    Size As Long
    ColorR As Byte
    ColorG As Byte
    ColorB As Byte
    TileConfig As PlanetConfig
    MoonData As MoonRec
    TimeToExplode As Long
    Type As Byte
End Type

Type ConstructionRec
    X As Long
    Y As Long
    ResourceNum As Long
End Type

Type SaibamanRec
    Working As Byte
    X As Long
    Y As Long
    TaskInit As String * 255
    Remaining As Long
End Type

Type MapSaibamansRec
    TotalSaibamans As Byte
    Saibaman(1 To 5) As SaibamanRec
End Type

Type PlayerPlanetRec
    PlanetData As PlanetRec
    LastLogin As String * 255
End Type

Function VIAGEMMAP() As Long
    If GetPlayerMap(MyIndex) = 1 Or GetPlayerMap(MyIndex) = 53 Or GetPlayerMap(MyIndex) = 54 Then
        VIAGEMMAP = GetPlayerMap(MyIndex)
    End If
End Function

Function InLevel(Level As Long) As Boolean
    Select Case GetPlayerMap(MyIndex)
        Case 1: If Level <= 25 Then InLevel = True
        Case 53: If Level > 25 And Level <= 50 Then InLevel = True
        Case 54: If Level > 50 Then InLevel = True
    End Select
End Function
