Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public Bank As BankRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public Events(1 To MAX_EVENTS) As EventWrapperRec
Public Effect(1 To MAX_EFFECTS) As EffectRec
Public Quest(1 To MAX_QUESTS) As QuestRec

' client-side stuff
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public TempMapNpc(1 To MAX_MAP_NPCS) As TempMapNpcRec
Public Autotile() As AutotileRec
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public MenuButton(1 To MAX_MENUBUTTONS) As OLD_ButtonRec
Public Party As PartyRec
Public GUIWindow() As GUIWindowRec
Public NewGUIWindow() As NewGUIWindowRec
Public Buttons(1 To MAX_BUTTONS) As ButtonRec
Public Chat(1 To 20) As ChatRec
Public MapSounds() As MapSoundRec
Public WeatherParticle(1 To MAX_WEATHER_PARTICLES) As WeatherParticleRec
Public Splash(1 To MAX_WEATHER_PARTICLES) As SplashRec
Public ProjectileList() As ProjectileRec
Public CurrentEvent As SubEventRec
Public EffectData() As Effect   'List of all the active effects
Public Transporte As TransporteRec

Public AmbientActor(1 To 30) As AmbienteRec
Public Cloud(1 To 10) As CloudRec
Public Buracos(1 To 10) As BuracoRec

Public BossMsg As BossMsgRec

Public Guild(1 To MAX_GUILDS) As GuildRec

Public Conquistas() As ConquistaRec

Type GuildMemberRec
    name As String * NAME_LENGTH
    Rank As Byte
    Level As Long
    Donations As Long
    GuildExp As Long
End Type

Type GuildRec
    name As String * NAME_LENGTH
    IconColor(1 To 25) As Byte
    Level As Long
    EXP As Long
    Member(1 To 10) As GuildMemberRec
    MOTD As String * 255
    TNL As Long
    Red As Long
    Blue As Long
    Yellow As Long
    Gold As Long
    UpBlock As Byte
End Type

Private Type BossMsgRec
    Message As String
    Created As Long
    color As Long
End Type

Public Options As OptionsRec

Private Type ChatRec
    Text As String
    colour As Long
End Type

Type ServerRec
    name As String
    IP As String
    Ping As Long
End Type

' Type recs
Private Type OptionsRec
    Game_Name As String
    savePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
    MenuMusic As String
    Music As Byte
    Sound As Byte
    Debug As Byte
    volume As Byte
    FPS As Byte
    Ambiente As Byte
    Tela As Byte
    Clima As Byte
    Neblina As Byte
    Window As Boolean
    Servers() As ServerRec
    Language As String
    PickMenu As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    num As Long
    value As Long
End Type

Type ConquistaRec
    name As String
    Desc As String
    EXP As Long
    Progress As Long
    Reward(1 To 5) As PlayerInvRec
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type SpellAnim
    SpellNum As Long
    Timer As Long
    FramePointer As Long
End Type

Type QuestStateRec
    State As Byte
    Date As String * 30
End Type

Public Type DiaryMissionRec
    LastDate As String
    MissionIndex As Long
    MissionObjective As Long
    MissionActual As Long
    Completed As Byte
End Type

Private Type PlayerRec
    ' General
    name As String
    Class As Long
    Sprite As Long
    Level As Long
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Long
    StatPoints(1 To Stats.Stat_Count - 1) As Long
    POINTS As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    ' Misc
    EventOpen(1 To MAX_EVENTS) As Byte
    Trans As Long
    TransAnimTick As Long
    EffectAnimTick As Long
    PDL As Long
    RawPDL As Long
    EsoBonus As Long
    EsoTime As Long
    EsoNum As Long
    EsoEffectTick As Long
    VIP As Byte
    ' Visual
    Hair As Byte
    IsDead As Byte
    ' Quest
    QuestState(1 To MAX_QUESTS) As QuestStateRec
    Titulo As Long
    InTutorial As Byte
    Guild As Long
    Diary As DiaryMissionRec
    VIPExp As Long
    Instance As Long
    Conquistas(1 To 80) As Byte
    ConquistaProgress(1 To 80) As Long
    NumServices As Long
    IsGod As Byte
    GodLevel As Long
    GodExp As Long
End Type

Private Type TempPlayerRec
' Client use only
    XOffSet As Double
    YOffSet As Double
    moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    StartFlash As Long
    AttackAnim As Byte
    Fly As Byte
    StunDuration As Long
    FlyBalance As Single
    HairChange As Byte
    'Spell Cast
    SpellBuffer As Long
    SpellBufferTimer As Long
    SpellBufferNum As Long
    HairAnim As Byte
    HairTrans As Byte
    HairAnimTick As Long
    KamehamehaLast As Long
    SpiritBombLast As Long
    MoveLast As Long
    MoveLastType As Byte
    AFK As Byte
    LastMove As Long
    speed As Long
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
    name As String * NAME_LENGTH
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

Private Type MapRec2
    name As String * NAME_LENGTH
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
    name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Long
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Private Type LuckySlotRec
    ItemNum As Integer
    Quant As Long
    Chance As Byte
End Type

Private Type ItemRec
    name As String * NAME_LENGTH
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
    
    LuckySlot(1 To 40) As LuckySlotRec
End Type

Private Type MapItemRec
    PlayerName As String
    num As Long
    value As Long
    Frame As Byte
    X As Byte
    Y As Byte
    Gravity As Long
    YOffSet As Long
    XOffSet As Long
    YOnSet As Long
    PlaySound As Boolean
    BalanceDir As Long
    BalanceValue As Long
    BalanceTick As Long
End Type

Type DropRec
    Chance As Long
    value As Long
    num As Long
End Type

Type IARec
    data(1 To 5) As Long
End Type

Private Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    Drop(1 To 10) As DropRec
    Stat(1 To Stats.Stat_Count - 1) As Long
    HP As Long
    EXP As Long
    Animation As Long
    Damage As Long
    Level As Long
    speed As Long
    Event As Long
    Effect As Long
    AttackSpeed As Long
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
    IA(1 To NPCIA.count - 1) As IARec
    IsPlanetable As Byte
    Evolution As Long
    ECostGold As Long
    ECostRed As Long
    ECostBlue As Long
    ECostYellow As Long
    TimeToEvolute As Long
    MinLevel As Long
End Type

Type MapNpcRec
    num As Long
    Target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    FlyOffSet As Long
    FlyOffSetTick As Long
    FlyOffsetDir As Byte
    MaxHP As Long
    PDL As Long
End Type

Private Type TempMapNpcRec
' Client use only
    XOffSet As Long
    YOffSet As Long
    moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    StartFlash As Long
    StunDuration As Long
    StunTick As Long
    AttackData1 As Long
    SpawnDelay As Byte
    AttackType As Byte
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem(1 To 5) As Long
    CostValue(1 To 5) As Long
End Type

Private Type ShopRec
    name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    name As String * NAME_LENGTH
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
    
    'Spell Linear Anim
    SpellLinearAnim(1 To 3) As Long
    
    'Cast anim
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

Public Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    name As String * NAME_LENGTH
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

Private Type ActionMsgRec
    Message As String
    Created As Long
    Type As Long
    color As Long
    Scroll As Long
    X As Long
    Y As Long
    Timer As Long
    Alpha As Byte
End Type

Private Type BloodRec
    Sprite As Long
    Timer As Long
    X As Long
    Y As Long
    Alpha As Byte
End Type

Private Type AnimationRec
    name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    XAxis(0 To 3) As Long
    YAxis(0 To 3) As Long
    
    Sprite(0 To 1, 0 To 3) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
    
    Tremor As Long
    Buraco As Byte
End Type

Private Type AnimInstanceRec
    Animation As Long
    X As Long
    Y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    Timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    frameIndex(0 To 1) As Long
    
    Dir As Byte
    
    IsLinear As Byte
    
    LockToNPC As Byte
    CastAnim As Byte
    ReturnAnim As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type OLD_ButtonRec
    Filename As String
    State As Byte
End Type

Public Type ButtonRec
    State As Byte
    X As Long
    Y As Long
    Width As Long
    Height As Long
    visible As Boolean
    PicNum As Long
End Type

Public Type GUIWindowRec
    X As Long
    Y As Long
    Width As Long
    Height As Long
    visible As Boolean
End Type

Public Type NewGUIWindowRec
    X As Long
    Y As Long
    Width As Long
    Height As Long
    visible As Boolean
    value As String
End Type

Public Type MapSoundRec
    X As Long
    Y As Long
    SoundHandle As Long
    InUse As Boolean
    channel As Long
End Type

Public Type WeatherParticleRec
    Type As Long
    X As Long
    Y As Long
    Velocity As Long
    InUse As Long
    Size As Long
End Type

Type SplashRec
    Tick As Long
    X As Long
    Y As Long
End Type

'Auto tiles "/
Public Type PointRec
    X As Long
    Y As Long
End Type

Public Type QuarterTileRec
    QuarterTile(1 To 4) As PointRec
    RenderState As Byte
    srcX(1 To 4) As Long
    srcY(1 To 4) As Long
End Type

Public Type AutotileRec
    Layer(1 To MapLayer.Layer_Count - 1) As QuarterTileRec
End Type

Public Type ChatBubbleRec
    Msg As String
    colour As Long
    Target As Long
    TargetType As Byte
    Timer As Long
    active As Boolean
    Alpha As Byte
End Type

Public Type ProjectileRec
    X As Long
    Y As Long
    tx As Long
    ty As Long
    RotateSpeed As Byte
    Rotate As Single
    Graphic As Long
End Type

Public Type SubEventRec
    Type As EventType
    HasText As Boolean
    Text() As String
    HasData As Boolean
    data() As Long
End Type

Private Type EventWrapperRec
    name As String
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
    name As String * NAME_LENGTH
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

Private Type Effect
    X As Single                 'Location of effect
    Y As Single
    GoToX As Single             'Location to move to
    GoToY As Single
    KillWhenAtTarget As Boolean     'If the effect is at its target (GoToX/Y), then Progression is set to 0
    KillWhenTargetLost As Boolean   'Kill the effect if the target is lost (sets progression = 0)
    Gfx As Byte                 'Particle texture used
    Used As Boolean 'If the effect is in use
    Alpha As Single
    Decay As Single
    Red As Single
    Green As Single
    Blue As Single
    XSpeed As Single
    YSpeed As Single
    XAcc As Single
    YAcc As Single
    EffectNum As Byte           'What number of effect that is used
    Modifier As Integer         'Misc variable (depends on the effect)
    FloatSize As Long           'The size of the particles
    Particles() As clsParticle  'Information on each particle
    Progression As Single       'Progression state, best to design where 0 = effect ends
    PartVertex() As TLVERTEX    'Used to point render particles ' Cant use in .NET maybe change
    PreviousFrame As Long       'Tick time of the last frame
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
    BindType As Byte
    BindIndex As Long       'Setting this value will bind the effect to move towards the character
    BindSpeed As Single         'How fast the effect moves towards the character
End Type

Private Type AmbienteRec
    X As Long
    Y As Long
    Dir As Long
    Used As Boolean
    speed As Byte
    Animation As Byte
    Size As Byte
    AnimTick As Long
End Type

Private Type CloudRec
    X As Long
    Y As Long
    Anim As Byte
    speed As Byte
    Use As Byte
    Alpha As Byte
    SizeX As Byte
    SizeY As Byte
End Type

Type BuracoRec
    X As Long
    Y As Long
    Size As Long
    InUse As Boolean
    Alpha As Byte
    IntervalTick As Long
    Map As Long
End Type

Type TransporteRec
    Tipo As Byte
    Tick As Long
    X As Double
    Y As Double
    Map As Long
    Anim As Byte
End Type

Type QuestRec
    name As String * NAME_LENGTH
    Desc As String * 255
    Icon As Long
    Repeat As Byte
    EventNum As Long
    NotDay As Byte
    NotNight As Byte
    Cooldown As Integer
    Type As Byte
End Type

Public Function IsConquistaEmpty() As Boolean
    Dim rv As Long

    On Error Resume Next

    rv = UBound(Conquistas)
    IsConquistaEmpty = (Err.Number = 0 Or Err.Number = 9)
End Function
