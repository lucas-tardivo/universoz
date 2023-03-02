Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the client's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SPlayerXYMap
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SMapKey
    SEditMap
    SShopEditor
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SLeft
    SResourceCache
    SResourceEditor
    SUpdateResource
    SSendPing
    SActionMsg
    SPlayerEXP
    SBlood
    SAnimationEditor
    SUpdateAnimation
    SAnimation
    SMapNpcVitals
    SCooldown
    SClearSpellBuffer
    SSayMsg
    SOpenShop
    SResetShopAction
    SStunned
    SMapWornEq
    SBank
    STrade
    SCloseTrade
    STradeUpdate
    STradeStatus
    STarget
    SHotbar
    SHighIndex
    SSound
    STradeRequest
    SPartyInvite
    SPartyUpdate
    SPartyVitals
    SPlayBGM
    SPlaySound
    SFadeoutBGM
    SStopSound
    SSwitchesAndVariables
    SChatBubble
    SSpecialEffect
    SFlash
    SEventData
    SEventEditor
    SEventUpdate
    SEffectEditor
    SUpdateEffect
    SEffect
    SMapReport
    SCreateProjectile
    SSendNews
    SNewsEditor
    SFly
    SSpecialAction
    SSpellBuffer
    SShenlong
    STransporte
    SMapNpcDataXY
    'Quests
    SUpdateQuest
    SQuestEditor
    SPlayerQuests
    SPlayerQuest
    'Info
    SPlayerInfo
    SOpenRefine
    SPlanets
    SMatchData
    SUpdateGuild
    SSaibamans
    SConquistas
    SSupport
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CDelAccount
    CLogin
    CAddChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CSaveItem
    CRequestEditNpc
    CSaveNpc
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell
    CSetAccess
    CWhosOnline
    CSetMotd
    CSpells
    CCast
    CQuit
    CSwapInvSlots
    CRequestEditResource
    CSaveResource
    CCheckPing
    CUnequip
    CRequestPlayerData
    CRequestItems
    CRequestNPCS
    CRequestResources
    CSpawnItem
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations
    CRequestSpells
    CRequestShops
    CForgetSpell
    CCloseShop
    CBuyItem
    CSellItem
    CChangeBankSlots
    CDepositItem
    CWithdrawItem
    CCloseBank
    CTradeRequest
    CAcceptTrade
    CDeclineTrade
    CTradeItem
    CUntradeItem
    CHotbarChange
    CHotbarUse
    CSwapSpellSlots
    CAcceptTradeRequest
    CDeclineTradeRequest
    CPartyRequest
    CAcceptParty
    CDeclineParty
    CPartyLeave
    CSwitchesAndVariables
    CRequestSwitchesAndVariables
    CSaveEventData
    CRequestEditEvents
    CRequestEventData
    CRequestEventsData
    CChooseEventOption
    CRequestEditEffect
    CSaveEffect
    CRequestEffects
    CTarget
    CEditNews
    CRequestEditNews
    CRequestNews
    CDevSuite
    COnDeath
    'quests
    CRequestEditQuest
    CSaveQuest
    CQuestInfo
    CUpgrade
    CSellPlanet
    CEnterGravity
    CCompleteTutorial
    CFeedback
    CCreateGuild
    CGuildAction
    CChallengeArena
    CAntiHackData
    CPlanetChange
    CConfirmation
    CSellEsp
    CSupport
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Mask2
    Fringe
    Fringe2
    MaskAnim
    FringeAnim
    Mask2Anim
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

' Sound entities
Public Enum SoundEntity
    seAnimation = 1
    seItem
    seNpc
    seResource
    seSpell
    seEffect
    ' Make sure SoundEntity_Count is below everything else
    SoundEntity_Count
End Enum

' Chat Log Enumerations
Public Enum ChatLog
    ChatGlobal = 0
    ChatMap
    ChatEmote
    ChatPlayer
    ChatSystem
End Enum

'***********************************************************
' These values below MUST match client side ones
'***********************************************************

Public Enum Colors
    Black = 0
    Blue
    Green
    Cyan
    Red
    Magenta
    Brown
    Grey
    DarkGrey
    brightblue
    brightgreen
    BrightCyan
    brightred
    Pink
    Yellow
    White
    DarkBrown
    Orange
End Enum

Public Enum AnswerType
    NO = 0
    YES
End Enum

Public Enum GenderType
     SEX_MALE = 0
     SEX_FEMALE
End Enum

Public Enum MapMoral
    MAP_MORAL_NONE = 0
    MAP_MORAL_SAFE
    MAP_MORAL_PRISON
    MAP_MORAL_OWNER
End Enum

Public Enum TileType
    TILE_TYPE_WALKABLE = 0
    tile_type_blocked
    tile_type_warp
    TILE_TYPE_ITEM
    TILE_TYPE_NPCAVOID
    TILE_TYPE_RESOURCE
    TILE_TYPE_NPCSPAWN
    TILE_TYPE_SHOP
    TILE_TYPE_BANK
    TILE_TYPE_HEAL
    TILE_TYPE_TRAP
    TILE_TYPE_SLIDE
    TILE_TYPE_SOUND
    TILE_TYPE_EVENT
    TILE_TYPE_ARENA
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
    ITEM_TYPE_TITULO
    ITEM_TYPE_EXTRATOR
    ITEM_TYPE_NAVE
    ITEM_TYPE_COMBUSTIVEL
    ITEM_TYPE_BAU
    ITEM_TYPE_VIP
    ITEM_TYPE_RADAR
    ITEM_TYPE_CAPTURE
    ITEM_TYPE_PLANETCHANGE
End Enum

Public Enum DirType
    DIR_UP = 0
    DIR_DOWN
    DIR_LEFT
    DIR_RIGHT
    DIR_UP_LEFT
    DIR_UP_RIGHT
    DIR_DOWN_LEFT
    DIR_DOWN_RIGHT
End Enum

Public Enum MovementType
    MOVING_WALKING = 1
    MOVING_RUNNING
End Enum

Public Enum AccessType
    ADMIN_MONITOR = 1
    ADMIN_MAPPER
    ADMIN_DEVELOPER
    ADMIN_CREATOR
End Enum

Public Enum NpcBehaviour
    NPC_BEHAVIOUR_ATTACKONSIGHT = 0
    NPC_BEHAVIOUR_ATTACKWHENATTACKED
    NPC_BEHAVIOUR_FRIENDLY
    NPC_BEHAVIOUR_SHOPKEEPER
    NPC_BEHAVIOUR_GUARD
    NPC_BEHAVIOUR_TREINO
    NPC_BEHAVIOUR_TREINOHOUSE
End Enum

Public Enum SpellType
    SPELL_TYPE_DAMAGEHP = 0
    SPELL_TYPE_DAMAGEMP
    SPELL_TYPE_HEALHP
    SPELL_TYPE_HEALMP
    SPELL_TYPE_WARP
    SPELL_TYPE_TRANS
    SPELL_TYPE_LINEAR
    SPELL_TYPE_VOAR
    SPELL_TYPE_SHUNPPO
End Enum

Public Enum TargetType
    TARGET_TYPE_NONE = 0
    TARGET_TYPE_PLAYER
    TARGET_TYPE_NPC
End Enum

Public Enum ActionMsgType
    ACTIONMSG_STATIC = 0
    ACTIONMSG_SCROLL
    ACTIONMSG_SCREEN
End Enum

Public Enum SpecialEffectType
    EFFECT_TYPE_FADEIN = 1
    EFFECT_TYPE_FADEOUT
    EFFECT_TYPE_FLASH
    EFFECT_TYPE_FOG
    EFFECT_TYPE_WEATHER
    EFFECT_TYPE_TINT
End Enum

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
    Evt_Quest
    'EventType_Count should be below everything else
    EventType_Count
End Enum

Public Enum ComparisonOperator
    GEQUAL = 0
    LEQUAL
    GREATER
    LESS
    EQUAL
    NOTEQUAL
End Enum

Public Enum AutomaticEvents
    None = 0
    PalavraMagica
    Aventura
    
    Count
End Enum

Public Enum NPCIA
    Spawn = 1
    Shunppo
    Stun
    Storm
    NOTDEFINED4
    NOTDEFINED5
    NOTDEFINED6
    NOTDEFINED7
    NOTDEFINED8
    NOTDEFINED9
    Count
End Enum

Public Enum GuildRank
    Member = 0
    Capitao
    Major
    Mestre
End Enum

Public Enum Daily
    DestruaPlanetas = 1
    DestruaPlanetasHalf
    GetRed
    GetBlue
    GetYellow
    GetAny
    KillElite
    KillFeralElite
    KillInsectElite
    KillHumanElite
    KillCreatures
    KillFeralCreatures
    KillInsectCreatures
    KillHumanCreatures
    GetGlobes
    Destroy
    DestroyArmy
    DestroyArt
    DestroyMayor
    DestroyGold
    DestroyMinor
End Enum

Public Enum ConfirmType
    NewPlanet = 1
    DestroyItem
End Enum

Public Enum PlayerInfoType
    AFK = 1
    Fish
    Gravidade
    GravityOk
    ProvacaoInit
    OpenGuild
    GuildInvite
    PlayerDaily
    OpenArena
    ArenaChallenging
    AntiHackData
    FabricaData
    ExercitoData
    Confirmation
    ConquistasInfo
    ConquistaInfo
    OpenTroca
    ServiceFeedback
    
    Count
End Enum

