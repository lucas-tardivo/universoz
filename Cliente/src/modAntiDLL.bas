Attribute VB_Name = "modAntiDLL"
Option Explicit
 
Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
 
Private Declare Function EnumProcesses Lib "PSAPI.DLL" ( _
   lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
 
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" ( _
    ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
 
Private Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" ( _
    ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
 
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400

Public Const ProccessList As String = "@ntdll.dll@kernel32.dll@KERNELBASE.dll@dx8vb.dll@msvcrt.dll@ADVAPI32.dll@sechost.dll@RPCRT4.dll@SspiCli.dll@CRYPTBASE.dll@ole32.dll@GDI32.dll@USER32.dll@LPK.dll@USP10.dll@MSACM32.dll@WINMM.dll@OLEAUT32.dll@d3dxof.dll@MSVBVM60.DLL@IMM32.DLL@MSCTF.dll@uxtheme.dll@SXS.DLL@CRYPTSP.dll@rsaenh.dll@dwmapi.dll@comctl32.DLL@SHLWAPI.dll@CLBCatQ.DLL@MSWINSCK.OCX@WSOCK32.dll@WS2_32.dll@NSI.dll@d3d8.dll@VERSION.dll@d3d8thk.dll@aticfx32.dll@atiu9pag.dll@atiumdag.dll@atiumdva.dll@POWRPROF.dll@SETUPAPI.dll@CFGMGR32.dll@DEVOBJ.dll@gdiplus.DLL@USERENV.dll@profapi.dll@WindowsCodecs.dll@mswsock.dll@wshtcpip.dll@NLAapi.dll@napinsp.dll@pnrpnsp.dll@DNSAPI.dll@winrnr.dll@IPHLPAPI.DLL@WINNSI.DLL@fwpuclnt.dll@rasadhlp.dll@fmod.dll@MMDevAPI.DLL@PROPSYS.dll@wdmaud.drv@ksuser.dll@AVRT.dll@AUDIOSES.DLL@msacm32.drv@midimap.dll@dsound.dll@PSAPI.dll"
Public Const ExceptionList As String = "@apphelp.dll@PSAPI.DLL@kernel32.dll"
Public NewProccessList As String
Public ProcessArray() As String
Public ProcessOccur() As Byte
Public AntiHackStarted As Boolean

Public Sub StartAntiHack()
    'NewProccessList = "@samcli.dll@sfc.dll@sfc_os.DLL@api-ms-win-downlevel-ole32-l1-1-0.dll@api-ms-win-downlevel-shlwapi-l1-1-0.dll@api-ms-win-downlevel-user32-l1-1-0.dl@WINSPOOL.DRV@MPR.dll@normaliz.DLL@api-ms-win-downlevel-normaliz-l1-1-0.dll@AcLayers.DLL@api-ms-win-downlevel-version-l1-1-0.dll@api-ms-win-downlevel-advapi32-l1-1-0.dll@api-ms-win-downlevel-shlwapi-l2-1-0.dll@Secur32.dll@nvd3dum.dll@api-ms-win-downlevel-advapi32-l2-1-0.dll@gbpinj.dll@igdumdx32.dll@igdumd32.dll@ntdll.dll@kernel32.dll@KERNELBASE.dll@MSVBVM60.DLL@USER32.dll@GDI32.dll@LPK.dll@USP10.dll@msvcrt.dll@ADVAPI32.dll@sechost.dll@RPCRT4.dll@ole32.dll@OLEAUT32.dll@apphelp.dll@SspiCli.dll@SHLWAPI.dll@UxTheme.dll@shdocvw.dll@AcLayers.DLL@MpcSafeDll.dll@nvspcap.dll@nvapi.dll@nvspcap.dll@WINHTTP.dll@webio.dll@nvapi.dll@resampledmo.dll@WinCRT.dll@Secur32.dll@msctfime.ime@WS2HELP.dll@IMAGEHLP.dll@COMRes.dll@MSVFW32.dll@xpsp2res.dll@hnetcfg.dll@wmvcore.dll@DRMClien.DLL@WMASF.DLL@wmidx.dll@quartz.dll@l3codeca.acm@ffdshow.ax@DINPUT.dll@"
    If App.LogMode = 0 Then Exit Sub
    frmAntiHack.Show
    'If Not FileExist(App.Path & "\universez.exe", True) Then
    '    FileCopy App.Path & "\" & App.EXEName & ".exe", App.Path & "\universez.exe"
    'End If
    'Dim WaitTick As Long
    'WaitTick = GetTickCount
    'Do While Not FileExist(App.Path & "\universez.exe", True) And Not WaitTick + 10000 < GetTickCount
    '    DoEvents
    'Loop
    'If Not FileExist(App.Path & "\universez.exe", True) Then
    '    MsgBox "Um arquivo " & App.Path & "\universez.exe está faltando para executar o jogo!", vbExclamation
    '    DestroyGame
    'Else
        'Shell App.Path & "\universez.exe silence", vbHide
        'DoEvents
        'WaitTick = GetTickCount
        'Do While Not WaitTick + 20000 < GetTickCount
        '    DoEvents
        'Loop
        
        If CreateLibrary Then
            ProcessArray = Split(ExceptionList & NewProccessList, "@")
            AntiHackStarted = True
            'Shell "TASKKILL /F /IM universez.exe"
            Unload frmAntiHack
        Else
            MsgBox "Houve uma falha na ativação do sistema de defesa, por favor reinicie o jogo", vbExclamation
            DestroyGame
        End If
    
    
End Sub
Public Sub AddDLL(DllName As String)
    NewProccessList = NewProccessList & "@" & DllName
End Sub

Public Sub ClearProcesses()
    ReDim ProcessOccur(1 To UBound(ProcessArray)) As Byte
End Sub

Public Sub CheckDLL(sName As String, Optional Relatorio As Boolean = False)
    Dim i As Long
    If LCase(sName) = LCase(App.EXEName & ".exe") Then Exit Sub
    If InStr(1, sName, "hack", vbTextCompare) > 0 Or InStr(1, sName, "speed", vbTextCompare) > 0 Or InStr(1, sName, "wpe", vbTextCompare) > 0 Or InStr(1, sName, "spy", vbTextCompare) > 0 Then
        If Relatorio Then HandleDLL "(ANTIHACK)" & sName
        MsgBox "Foram encontradas entradas proibidas no seu jogo!", vbCritical, "Anti-hack"
        DestroyGame
    End If
    
    For i = 1 To UBound(ProcessArray)
        If Len(Trim$(ProcessArray(i))) = Len(Trim$(sName)) Then
            If LCase(Trim$(ProcessArray(i))) = LCase(Trim$(sName)) Then
                If ProcessOccur(i) = 0 Then
                    ProcessOccur(i) = 1
                Else
                    'MsgBox "Seu jogo será fechado pois foram constadas mais de uma ocorrencia de uma DLL conectadas ao jogo (Nome da dll: " & sName & "!)"
                    'DestroyGame
                    'HandleDLL "Mais de uma ocorrencia:" & sName
                End If
                Exit Sub
            End If
        End If
    Next i
    'MsgBox "Seu jogo será fechado pois foram constadas DLLs ilegais conectadas ao jogo (Nome da dll: " & sName & "!)"
    'DestroyGame
    If Relatorio Then HandleDLL sName
End Sub

Public Sub VerificarAntiHack(Optional Relatorio As Boolean = False)
    If App.LogMode = 0 Then Exit Sub
    If Not AntiHackStarted Then Exit Sub
    Const MAX_PATH As Long = 260
    Dim lProcesses() As Long, lModules() As Long, N As Long, lRet As Long, hProcess As Long
    Dim sName As String
    Dim sProcess As String
    sProcess = App.EXEName & ".exe"
    
    sProcess = UCase$(sProcess)
    
    ReDim lProcesses(1023) As Long
    ClearProcesses
    If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
        For N = 0 To (lRet \ 4) - 1
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(N))
            If hProcess Then
                ReDim lModules(1023)
                If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
                    sName = String$(MAX_PATH, vbNullChar)
                    GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
                    sName = Left$(sName, InStr(sName, vbNullChar) - 1)
                    If Len(sName) = Len(sProcess) Then
                        If sProcess = UCase$(sName) Then
                            
                            Dim i As Long
                            For i = 1 To UBound(lModules)
                                sName = String$(MAX_PATH, vbNullChar)
                                GetModuleBaseName hProcess, lModules(i), sName, MAX_PATH
                                sName = Left$(sName, InStr(sName, vbNullChar) - 1)
                                If Len(sName) = Len(sProcess) Then
                                    If sProcess = UCase$(sName) Then
                                        Exit For
                                    End If
                                End If
                                CheckDLL sName, Relatorio
                            Next i
                            
                        End If
                    End If
                End If
            End If
            CloseHandle hProcess
        Next N
    End If
End Sub

Public Function CreateLibrary() As Boolean
    Const MAX_PATH As Long = 260
    Dim lProcesses() As Long, lModules() As Long, N As Long, lRet As Long, hProcess As Long
    Dim sName As String
    Dim sProcess As String
    sProcess = App.EXEName & ".exe"
    
    sProcess = UCase$(sProcess)
    
    CreateLibrary = False
    
    ReDim lProcesses(1023) As Long
    If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
        For N = 0 To (lRet \ 4) - 1
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(N))
            If hProcess Then
                ReDim lModules(1023)
                If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
                    sName = String$(MAX_PATH, vbNullChar)
                    GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
                    sName = Left$(sName, InStr(sName, vbNullChar) - 1)
                    If Len(sName) = Len(sProcess) Then
                        If sProcess = UCase$(sName) Then
                            
                            Dim i As Long
                            For i = 1 To UBound(lModules)
                                sName = String$(MAX_PATH, vbNullChar)
                                GetModuleBaseName hProcess, lModules(i), sName, MAX_PATH
                                sName = Left$(sName, InStr(sName, vbNullChar) - 1)
                                If Len(sName) = Len(sProcess) Then
                                    If sProcess = UCase$(sName) Then
                                        Exit For
                                    End If
                                End If
                                AddDLL sName
                                CreateLibrary = True
                            Next i
                            
                        End If
                    End If
                End If
            End If
            CloseHandle hProcess
        Next N
    End If
End Function
