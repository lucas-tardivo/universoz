Attribute VB_Name = "modPlayerData"
' String constants
Public Const CHARACTER_VERSION As Byte = 1
Public Const NAME_LENGTH As Byte = 20
Public Const MUSIC_LENGTH As Byte = 40
Public Const ACCOUNT_LENGTH As Byte = 12

Public Const MAX_INV As Long = 35
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_HOTBAR As Long = 12
Public Const MAX_SWITCHES As Long = 1000
Public Const MAX_VARIABLES As Long = 1000
Public Const MAX_EVENTS As Long = 1000
Public Const MAX_QUESTS As Byte = 255

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1    'Força
    Endurance       'Constituição
    Intelligence    'KI
    agility         'Destreza
    Willpower       'Técnica
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    helmet
    shield
    
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

Type QuestStateRec
    State As Byte
    Date As String * 30
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
End Type

Public Type HotbarRec
    slot As Long
    sType As Byte
End Type

Public Type DailyMissionRec
    LastDate As String * 255
    MissionIndex As Long
    MissionObjective As Long
    MissionActual As Long
    Completed As Byte
    DailyBonus As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
End Type

Public Type PlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Long
    Exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    stat(1 To Stats.Stat_Count - 1) As Long
    Points As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    
    Switches(0 To MAX_SWITCHES) As Byte
    Variables(0 To MAX_VARIABLES) As Long
    
    EventOpen(1 To MAX_EVENTS) As Byte
    
    EsoBonus As Long
    EsoTime As Long
    EsoNum As Long
    
    VIP As Byte
    VIPDias As Long
    VIPData As String
    
    Hair As Byte
    PDL As Long
    IsDead As Byte
    
    QuestState(1 To MAX_QUESTS) As QuestStateRec
    
    Titulo As Long
    statPoints(1 To Stats.Stat_Count - 1) As Long
    NaveEspacial As Long
    GravityHours As Long
    GravityValue As Long
    GravityInit As String * 255
    InTutorial As Byte
    RealSprite As Long
    TopStars As Integer
    Guild As Long
    
    Daily As DailyMissionRec
    VIPExp As Long
    LastLogin As String * 255
    PlanetNum As Long
    Version As Byte
    Conquistas(1 To 80) As Byte
    ConquistaProgress(1 To 80) As Long
    NumServices As Long
    PlayerHouseNum As Long
    IsGod As Byte
    GodLevel As Long
    GodExp As Long
End Type
