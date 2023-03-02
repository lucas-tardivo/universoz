Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the server's packet enumeration

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

Public HandleDataSub(SMSG_COUNT) As Long

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Endurance
    Intelligence
    Agility
    Willpower
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
    Helmet
    Shield
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

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

Public Enum GUIType
    GUI_CHAT = 1
    GUI_HOTBAR
    GUI_MENU
    GUI_BARS
    GUI_INVENTORY
    GUI_SPELLS
    GUI_CHARACTER
    GUI_OPTIONS
    GUI_PARTY
    GUI_DESCRIPTION
    GUI_MAINMENU
    GUI_SHOP
    GUI_BANK
    GUI_TRADE
    GUI_CURRENCY
    GUI_DIALOGUE
    GUI_EVENTCHAT
    GUI_NEWS
    GUI_DEATH
    GUI_QUESTS
    GUI_CONQUISTAS
    Gui_Count
End Enum

Public Enum MenuType
    MENU_MAIN = 1
    MENU_LOGIN
    MENU_REGISTER
    MENU_CREDITS
    MENU_CLASS
    MENU_NEWCHAR
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
    
    ComparisonOperator_Count
End Enum

Public Enum NewGui
    Loading = 1
    Esfera
    Fader
    Nuvens
    Login
    Logo
    Flyingbanner
    Flyingbanner2
    Banner
    Personagem
    Estrela
    buttonlogin
    FaderBlack
    NewChar
    Cabelo
    Criar
    Pele
    Nuvens2
    Nuvens3
    
    NewGui_Count
End Enum

Public Enum NewGUIWindows
    TEXTLOGIN = 1
    TEXTPASSWORD
    LOGINBUTTON
    TEXTCHARNAME
    HAIRBUTTON
    COLORBUTTON
    CREATEBUTTON
    
    NewGui_Count
End Enum

Public Enum NPCIA
    Spawn = 1
    Shunppo
    Stun
    NOTDEFINED3
    NOTDEFINED4
    NOTDEFINED5
    NOTDEFINED6
    NOTDEFINED7
    NOTDEFINED8
    NOTDEFINED9
    count
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
    
    count
End Enum
