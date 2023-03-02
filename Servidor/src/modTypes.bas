Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map() As MapRec
Public MapCache() As Cache
Public PlayersOnMap() As Long
Public ResourceCache() As ResourceCacheRec
Public Player() As PlayerRec
Public Bank() As BankRec
Public TempPlayer() As TempPlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem() As MapItemRec
Public MapNpc() As MapDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Options As OptionsRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public MapBlocks() As MapBlockRec
Public Events(1 To MAX_EVENTS) As EventWrapperRec
Public Effect(1 To MAX_EFFECTS) As EffectRec
Public House() As HouseRec
Public Transporte() As TransporteRec
Public Quest(1 To MAX_QUESTS) As QuestRec
Public Provação() As ProvRec
Public Wish() As WishRec
Public NPCBase() As NPCBaseRec
Public AoEEffect(1 To MAX_AOEEFFECTS) As AoeEffectRec
Public Guild(1 To MAX_GUILDS) As GuildRec
Public DailyMission() As DailyRec
Public ArenaChallenge As ArenaChallengeRec
Public Conquistas() As ConquistaRec

Type ConquistaRec
    Name As String
    Desc As String
    Exp As Long
    Progress As Long
    Reward(1 To 5) As PlayerInvRec
End Type

Type ArenaChallengeRec
    Active As Byte
    Aposta As Long
    Players(1 To 6) As String
    PlayerAccept(1 To 6) As Boolean
    TotalPlayers As Byte
    MatchType As Byte
    LastCall As Long
    Tick As Long
    Count As Byte
    CountTick As Long
End Type

Type DailyRec
    Description As String
    NumberFactory As Double
    Subtype() As String
End Type

Type GuildMemberRec
    Name As String * NAME_LENGTH
    Rank As Byte
    Level As Long
    Donations As Long
    GuildExp As Long
End Type

Type GuildRec
    Name As String * NAME_LENGTH
    IconColor(1 To 25) As Byte
    Level As Long
    Exp As Long
    Member(1 To 10) As GuildMemberRec
    MOTD As String * 255
    TNL As Long
    Red As Long
    Blue As Long
    Yellow As Long
    Gold As Long
    UpBlock As Byte
End Type

Type AoeEffectRec
    X As Long
    Y As Long
    Map As Long
    Caster As String
    SpellNum As Long
    Tick As Long
    Duration As Long
    CastTick As Long
    CasterType As Byte
End Type

Type NPCBaseRec
    Damage As Long
    HP As Long
    Acc As Long
    Esq As Long
    Exp As Long
End Type

Type WishRec
    Phrase As String
    Type As Byte
    Event As Long
    Item As Long
    ItemVal As Long
End Type

Type HouseRec
    Proprietario As String
    DataDeInicio As String
    Dias As Long
End Type

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Website As String
    Logs As Byte
    EventChance As Long
    Language As String
    AuthID As String
    Password As String
    ExpFactor As Double
    DropFactor As Double
    ResourceFactor As Double
    GoldFactor As Double
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Type Cache
    Data() As Byte
End Type

Public Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    Target As Long
    tType As Byte
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    TargetType As Byte
    Target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
    'InEventWith
    CurrentEvent As Long
    inDevSuite As Byte
    'trans
    Trans As Long
    'fly
    Fly As Byte
    'afk
    LastMove As Long
    AlertMSG As Byte
    'visual
    HairChange As Byte
    'combat
    ImpactedBy As Long
    ImpactedTick As Long
    'fish
    NextFish As Long
    MatchIndex As Long
    PlanetCaptured As Long
    Speed As Long
    Nave As Long
    RespawnTick As Long
    GuildInvite As Long
    GuildInviteIndex As Long
    Instance As Long
    Confirmation As Byte
    ConfirmationVar As Long
    PlanetService As Long
    OnlineServices As Long
    DamageAmount As Long
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    data1 As Long
    data2 As Long
    data3 As Long
    Data4 As String
    DirBlock As Byte
End Type

Public Type TileRec2
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    data1 As Long
    data2 As Long
    data3 As Long
    DirBlock As Byte
End Type

Type MapRec
    Name As String * NAME_LENGTH
    Music As String * NAME_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    BossNpc As Long
    Fog As Long
    FogSpeed As Long
    FogOpacity As Long
    Panorama As Long
    SunRays As Long
    Fly As Byte
    BGS As String * MUSIC_LENGTH
    Weather As Long
    WeatherIntensity As Long
    Red As Long
    Green As Long
    Blue As Long
    Alpha As Long
    NpcSpawnType(1 To MAX_MAP_NPCS) As Long
    Ambiente As Byte
    FogDir As Byte
End Type

Type MapRec2
    Name As String * NAME_LENGTH
    Music As String * MUSIC_LENGTH
    BGS As String * MUSIC_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    Weather As Long
    WeatherIntensity As Long
    
    Fog As Long
    FogSpeed As Long
    FogOpacity As Long
    
    Red As Long
    Green As Long
    Blue As Long
    Alpha As Long
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    NpcSpawnType(1 To MAX_MAP_NPCS) As Long
    Panorama As Long
    
    Fly As Byte
    
    Ambiente As Byte
    FogDir As Byte
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    stat(1 To Stats.Stat_Count - 1) As Long
    MaleSprite() As Long
    FemaleSprite() As Long
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
End Type

Private Type LuckySlotRec
    ItemNum As Integer
    Quant As Long
    Chance As Byte
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Pic As Long

    Type As Byte
    data1 As Long
    data2 As Long
    data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Long
    Rarity As Byte
    Speed As Long
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
    
    LuckySlot(1 To 40) As LuckySlotRec
End Type

Type MapItemRec
    Num As Long
    Value As Long
    X As Byte
    Y As Byte
    ' ownership + despawn
    PlayerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
End Type

Type DropRec
    Chance As Long
    Value As Long
    Num As Long
End Type

Type IARec
    Data(1 To 5) As Long
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    Drop(1 To 10) As DropRec
    stat(1 To Stats.Stat_Count - 1) As Long
    HP As Long
    Exp As Long
    Animation As Long
    Damage As Long
    Level As Long
    Speed As Long
    Event As Long
    Effect As Long
    attackspeed As Long
    ND As Integer
    
    Fly As Byte
    FlyTick As Long
    Shadow As Byte
    
    Ranged As Byte
    ArrowAnim As Long
    ArrowDamage As Long
    ArrowAnimation As Long
    
    GFXPack As Byte
    
    'IA
    IA(1 To NPCIA.Count - 1) As IARec
    IsPlanetable As Byte
    Evolution As Long
    ECostGold As Long
    ECostRed As Long
    ECostBlue As Long
    ECostYellow As Long
    TimeToEvolute As Long
    MinLevel As Long
End Type

Public Type MapNpcRec
    Num As Long
    Target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    X As Byte
    Y As Byte
    Dir As Byte
    WalkingTick As Long
    Spawned As Byte
    Level As Long
    Points As Long
    PDL As Long
End Type

Type TempMapNpcRec
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' combat
    ImpactedBy As Long
    ImpactedTick As Long
    LastStorm As Long
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    costitem(1 To 5) As Long
    costvalue(1 To 5) As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
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
    Upgrade As Long
    Requisite As Long
    Dir As Byte
    Vital As Long
    Item As Long
    Price As Long
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
    
    Impact As Byte
    
    AoEDuration As Long
    AoETick As Long
    LinearRange As Byte
End Type

Type MapDataRec
    Npc() As MapNpcRec
    TempNpc() As TempMapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    X As Long
    Y As Long
    cur_health As Long
    ResourceNum As Long
End Type

Type ResourceCacheRec
    Resource_Count As Long
    ExtractorCount As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    WalkThrough As Boolean
    Animation As Long
    Effect As Long
    
    IsPlanetable As Byte
    Evolution As Long
    ECostGold As Long
    ECostRed As Long
    ECostBlue As Long
    ECostYellow As Long
    TimeToEvolute As Long
    NucleoLevel As Long
    ResourceLevel As Long
    MinLevel As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    XAxis(0 To 3) As Long
    YAxis(0 To 3) As Long
    
    Sprite(0 To 1, 0 To 3) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
    
    Tremor As Long
    Buraco As Byte
End Type

Public Type Vector
    X As Long
    Y As Long
End Type

Public Type MapBlockRec
    Blocks() As Long
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

Private Type EffectRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    isMulti As Byte
    MultiParticle(1 To MAX_MULTIPARTICLE) As Long
    Type As Long
    Sprite As Long
    Particles As Long
    Size As Single
    Alpha As Single
    Decay As Single
    Red As Single
    Green As Single
    Blue As Single
    data1 As Long
    data2 As Long
    data3 As Long
    Duration As Single
    XSpeed As Single
    YSpeed As Single
    XAcc As Single
    YAcc As Single
End Type

Private Type TransporteRec
    Nome As String
    Map As Long
    AlterMap As Long
    LoadMap As Long
    LoadX As Long
    LoadY As Long
    TravelMap As Long
    DestinyMap As Long
    DestinyX As Long
    DestinyY As Long
    AlterDestinyMap As Long
    AlterDestinyX As Long
    AlterDestinyY As Long
    Tick As Long
    IntervalTravel As Long
    IntervalWait As Long
    State As Byte
    Embarque As Long
    Sound As String
    Passaporte As Long
End Type

Type QuestRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Icon As Long
    Repeat As Byte
    EventNum As Long
    NotDay As Byte
    NotNight As Byte
    Cooldown As Integer
    Type As Byte
End Type

Type WaveRec
    Enemy(1 To MAX_MAP_NPCS) As MapNpcRec
    WaveTimer As Long
End Type

Type ProvRec
    Map As Long
    X As Byte
    Y As Byte
    Cost As Long
    MinLevel As Byte
    RewardItem As Long
    TradeItem As Long
    RewardXP As Long
    Wave() As WaveRec
    ActualTick As Long
    ActualWave As Byte
    ProvaçãoIndex As Long
End Type

Public Function IsTransporteEmpty() As Boolean
    Dim rv As Long

    On Error Resume Next

    rv = UBound(Transporte)
    IsTransporteEmpty = (Err.Number = 0 Or Err.Number = 9)
End Function

Public Function IsConquistaEmpty() As Boolean
    Dim rv As Long

    On Error Resume Next

    rv = UBound(Conquistas)
    IsConquistaEmpty = (Err.Number = 0 Or Err.Number = 9)
End Function
