@echo off
IF EXIST "%WinDir%\SysWOW64" (
	SET Sistema=%WinDir%\SysWOW64
) ELSE (
	SET Sistema=%WinDir%\System32
)

ECHO ** Universo Z Arquivos Necessarios **

ECHO Registrando DLL do DirectX 7
ECHO Executando Regsvr32 em %Sistema% dx7vb.dll
%Sistema%\regsvr32 /s %Sistema%\dx7vb.dll
IF errorlevel 3 (
	ECHO "Falha ao registrar a DLL dx7vb.dll"
) ELSE (
	ECHO Completo!
)


ECHO Registrando DLL do DirectX 8
ECHO Executando Regsvr32 em %Sistema% dx8vb.dll
%Sistema%\regsvr32 /s %Sistema%\dx8vb.dll
IF errorlevel 3 (
	ECHO "Falha ao registrar a DLL dx8vb.dll"
) ELSE (
	ECHO Completo!
)

ECHO Registrando DLL dos recursos do sistema
ECHO Executando Regsvr32 em %Sistema% mscomctl.ocx
%Sistema%\regsvr32 /s %Sistema%\mscomctl.ocx
IF errorlevel 3 (
	ECHO "Falha ao registrar a DLL mscomctl.ocx"
) ELSE (
	ECHO Completo!
)

ECHO Registrando DLL do componente de conexao
ECHO Executando Regsvr32 em %Sistema% MSWINSCK.OCX
%Sistema%\regsvr32 /s %Sistema%\MSWINSCK.OCX
IF errorlevel 3 (
	ECHO "Falha ao registrar a DLL MSWINSCK.OCX"
) ELSE (
	ECHO Completo!
)

ECHO Registrando DLL do componente de texto
ECHO Executando Regsvr32 em %Sistema% RICHTX32.OCX
%Sistema%\regsvr32 /s %Sistema%\RICHTX32.OCX
IF errorlevel 3 (
	ECHO "Falha ao registrar a DLL RICHTX32.OCX"
) ELSE (
	ECHO Completo!
)

ECHO Registrando DLL do componente de abas
ECHO Executando Regsvr32 em %Sistema% TABCTL32.OCX
%Sistema%\regsvr32 /s %Sistema%\TABCTL32.OCX
IF errorlevel 3 (
	ECHO "Falha ao registrar a DLL TABCTL32.OCX"
) ELSE (
	ECHO Completo!
)

ECHO Registrando DLL do componente de abas
ECHO Executando Regsvr32 em %Sistema% MSINET.OCX
%Sistema%\regsvr32 /s %Sistema%\MSINET.OCX
IF errorlevel 3 (
	ECHO "Falha ao registrar a DLL MSINET.OCX"
) ELSE (
	ECHO Completo!
)

ECHO ** FIM DE PROCESSO **
ECHO Equipe GoPlay Games

pause