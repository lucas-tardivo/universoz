Attribute VB_Name = "modConstants"

Option Explicit

' API
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' path constants
Public Const MAX_LINES As Long = 500 ' Used for frmServer.txtText
'- Pathfinding Constant -
'1 is the old method, faster but not smart at all
'2 is the new method, smart but can slow the server down if maps are huge and alot of npcs have targets.
Public Const PathfindingType As Long = 1
Public Const Decorations As Boolean = True

' server-side stuff
Public Const ITEM_SPAWN_TIME As Long = 30000 ' 30 seconds
Public Const ITEM_DESPAWN_TIME As Long = 90000 ' 1:30 seconds
Public Const MAX_DOTS As Long = 30

' ********************************************
' Default starting location [Server Only]
Public START_MAP As Long
Public START_X As Byte
Public START_Y As Byte
' Default respawn location [Server Only]
Public RESPAWN_MAP As Long
Public RESPAWN_X As Byte
Public RESPAWN_Y As Byte
' ********************************************************
' * The values below must match with the client's values *
' ********************************************************
' General constants
Public MAX_PLAYERS As Long
Public Const MAX_ITEMS As Long = 500
Public Const MAX_NPCS As Long = 255
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_MAP_ITEMS As Long = 255
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_SHOPS As Long = 50
Public Const MAX_SPELLS As Long = 255
Public Const MAX_TRADES As Long = 30
Public Const MAX_RESOURCES As Long = 100
Public MAX_LEVELS As Long
Public MAX_STAT_LEVELS As Long
Public Const MAX_BANK As Long = 99
Public Const MAX_PARTYS As Long = 35
Public Const MAX_PARTY_MEMBERS As Long = 4
Public Const MAX_EFFECTS As Byte = 255
Public Const MAX_MULTIPARTICLE As Byte = 5
Public Const MAX_AOEEFFECTS As Byte = 30
Public Const MAX_GUILDS As Byte = 200

' text color constants
Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = brightblue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = brightgreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = brightblue
Public Const WhoColor As Byte = brightblue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = brightblue


' Map constants
Public MAX_MAPS As Long
Public Const MAX_MAPX As Byte = 24
Public Const MAX_MAPY As Byte = 18
Public Const ARENA_MAP As Byte = 20

' server configs
'Game
Public Const SECURELEVEL As Byte = 30

' Moeda
Public MoedaZ As Long
Public Const EspV As Long = 80
Public Const EspAz As Long = 81
Public Const EspAm As Long = 82
Public Const TesouroItem As Long = 202

'Animations
Public Const PlayerAttackAnim As Byte = 1
Public Const PlayerLevelUpAnim As Byte = 34
Public Const StatLevelUpAnim As Byte = 35
Public Const DeathEffect As Long = 6

'AFK
Public Const AFKTime As Long = 60000

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

' Pesc
Public Const MarginFish As Integer = 1000
Public Const EventGlobalInterval As Long = 3600000
Public Const SalaDoTempo As Long = 52
Public Const GravityMap As Long = 3
