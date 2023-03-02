Attribute VB_Name = "modGlobals"
Option Explicit

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Long

' Text vars
Public vbQuote As String

' Maximum classes
Public Max_Classes As Long

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines(ChatSystem) As Long

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Long

' high indexing
Public Player_HighIndex As Long

' lock the CPS?
Public CPSUnlock As Boolean

' Packet Tracker
Public PacketsIn As Long
Public PacketsOut As Long

' Server Online Time
Public ServerSeconds As Long
Public ServerMinutes As Long
Public ServerHours As Long

' Houses
Public TotalHouses As Long

'Shenlong
Public ShenlongTick As Long
Public ShenlongActive As Byte
Public ShenlongMap As Long
Public ShenlongX As Long
Public ShenlongY As Long
Public ShenlongOwner As String

' Exp
Public Experience() As Currency
Public StatExperience() As Currency
Public LevelUpBonus As Long
Public ExpToPDL As Double

'Provações
Public ProvaçãoCount As Byte

'XP
Public PDLBase() As Long

Public DragonballInfo(1 To 7) As Long
Public Relogio As Long
Public EspAmount(1 To 3) As Long

Public PlanetInService() As Boolean

'Invasões
Public EventGlobalTick As Long
Public EventGlobalType As Long
