Attribute VB_Name = "ModINI"
'::::::::::::::::::::::::::::::
':::::Weron Onix Engine :::::::
':::::::::::By:::::::::::::::::
':::::::::::::Equipe SpyRP:::::
':::::::: Knooz_Admin::::::::::
'::::::::::::::::::::::::::::::

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Public SOffsetX As Integer
Public SOffsetY As Integer
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
 Dim retlen As String
 Dim Ret As String
 Ret = String$(255, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
 Ret = Left$(Ret, retlen)
 ReadINI = Ret
 'Escrever
 'Call WriteINI("Geral", "Tempo", txttempo.Text, App.Path & "\show.ini")
 'Call WriteINI("Geral", "Ajuda", Text1.Text, App.Path & "\show.ini")
 'Call WriteINI("Geral", "Atualiza", Text2.Text, App.Path & "\show.ini")
 'Call WriteINI("Server", "Status", Text5.Text, App.Path & "\show.ini")
 '
 'ler
 'valortempo = ReadINI("Geral", "Tempo", App.Path & "\show.ini")
 'valorajuda = ReadINI("Geral", "Ajuda", App.Path & "\show.ini")
 'atualizaperguntas = ReadINI("Geral", "Atualiza", App.Path & "\show.ini")
End Function

Public Sub WriteINI(Secao As String, Entrada As String, Texto As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
  'texto= valor que vem depois do igual
  WritePrivateProfileString Secao, Entrada, Texto, Arquivo
End Sub

Sub SpecialPutVar(File As String, _
   Header As String, _
   Var As String, _
   value As String)

    ' Igual ao de baixo, exceto que fica tudo 0 e valores em branco. (usado para configurações)
    Call WritePrivateProfileString(Header, Var, value, File)
End Sub
Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, Var As String, value As String)

        Call WritePrivateProfileString(Header, Var, value, File)

End Sub
Sub MovePicture(pb As Object, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        pb.Left = pb.Left + x - SOffsetX
        pb.Top = pb.Top + Y - SOffsetY
    End If
End Sub

Function FileExist(ByVal Filename As String) As Boolean
    If Dir(Filename) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function
Function Rands(ByVal Low As Long, ByVal High As Long)
Randomize
Do Until Rands >= Low
    Rands = Int(Rnd * High)
    High = High + 1
    DoEvents
Loop
End Function
'Generates Randsom numbers
Public Function RandsomNo(Max As Long, Optional Last As Integer) As Long
Dim a, b
If Val(Last) < 1 Then Last = 100


If Max < 1 Then
RandsomNo = 0
Exit Function
End If

a = Rnd(Last)
b = Mid(a, InStr(1, a, ".", vbTextCompare) + 1, Len(Str(Max)) - 1)


If b > Max Then
b = b - ((b \ Max) * Max)
End If

If b < 1 Then b = 0
RandsomNo = b
End Function
Public Function Rand(ByVal Low As Long, _
                     ByVal High As Long) As Long
  Rand = Int((High - Low + 1) * Rnd) + Low
End Function
Function LoadTextFile(Filename As String) As String
Dim f As Long
f = FreeFile
Open Filename For Input As #f
    LoadTextFile = Input$(LOF(f), f)
Close #f
End Function
Sub WriteTextFile(Filename As String, Text As String)
Dim f As Long
f = FreeFile
Open Filename For Output As #f
    Print #f, Text
Close #f
End Sub
