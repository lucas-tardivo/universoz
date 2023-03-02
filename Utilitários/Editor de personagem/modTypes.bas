Attribute VB_Name = "modTypes"
'For Clear functions
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 12



Public Const MAX_ITEMS As Long = 5000
Public Const MAX_SPELLS As Long = 255

Public Player As PlayerRec

Public Item(1 To MAX_ITEMS) As ItemRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public Events(1 To MAX_EVENTS) As EventWrapperRec

Public Enum EventType
    Evt_Message = 0
    Evt_Menu
    Evt_Quit
    Evt_OpenShop
    Evt_OpenBank
    Evt_GiveItem
    Evt_ChangeLevel
    Evt_PlayAnimation
    Evt_Warp
    Evt_GOTO
    Evt_Switch
    Evt_Variable
    Evt_AddText
    Evt_Chatbubble
    Evt_Branch
    Evt_ChangeSkill
    Evt_ChangeSprite
    Evt_ChangePK
    Evt_SpawnNPC
    Evt_ChangeClass
    Evt_ChangeSex
    Evt_ChangeExp
    Evt_SpecialEffect
    Evt_PlaySound
    Evt_PlayBGM
    Evt_StopSound
    Evt_FadeoutBGM
    Evt_SetAccess
    Evt_CustomScript
    Evt_OpenEvent
    'EventType_Count should be below everything else
    EventType_Count
End Enum

Public Enum ItemType
    ITEM_TYPE_NONE = 0
    ITEM_TYPE_WEAPON
    ITEM_TYPE_ARMOR
    ITEM_TYPE_HELMET
    ITEM_TYPE_SHIELD
    ITEM_TYPE_CONSUME
    ITEM_TYPE_CURRENCY
    ITEM_TYPE_SPELL
    ITEM_TYPE_SCOUTER
    ITEM_TYPE_ESOTERICA
    ITEM_TYPE_DRAGONBALL
End Enum


Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Pic As Long

    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Long
    Rarity As Byte
    speed As Long
    Handed As Byte
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Long
    Animation As Long
    Paperdoll As Long
    
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    Stackable As Byte
    Effect As Long
    
    Projectile As Long
    Range As Byte
    Rotation As Integer
    Ammo As Long
    
    EsotericaTime As Long
    EsotericaBonus As Long
    
    CantDrop As Byte
    Dragonball As Byte
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    Effect As Long
    
    Add_Stat(1 To Stats.Stat_Count - 1) As Long
    SpriteTrans As Long
    TransAnim As Long
    PDLBonus As Long
    
    SpellLinearAnim(1 To 3) As Long
    CastPlayerAnim As Byte
    HairChange As Byte
    
    TransVital(1 To Vitals.Vital_Count - 1) As Long
    
    Projectile As Long
    RotateSpeed As Byte
End Type

Private Type SubEventRec
    Type As EventType
    HasText As Boolean
    Text() As String * 250
    HasData As Boolean
    Data() As Long
End Type

Private Type EventWrapperRec
    Name As String * NAME_LENGTH
    chkSwitch As Byte
    chkVariable As Byte
    chkHasItem As Byte
    
    SwitchIndex As Long
    SwitchCompare As Byte
    VariableIndex As Long
    VariableCompare As Byte
    VariableCondition As Long
    HasItemIndex As Long
    
    HasSubEvents As Boolean
    SubEvents() As SubEventRec
    
    Trigger As Byte
    WalkThrought As Byte
End Type
