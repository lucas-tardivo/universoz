Attribute VB_Name = "modAutoUpdater"
' file host
Private UpdateURL As String

' stores the variables for the version downloaders
Private VersionCount As Long
Public Sub Update()
Dim CurVersion As Long
Dim filename As String
Dim i As Long
    UpdateURL = GetVar(App.Path & "\data files\config.ini", "UPDATER", "URL")

    If UpdateURL <> "" Then
    ' get the file which contains the info of updated files
    DownloadFile UpdateURL & "/masterversion.ini", App.Path & "\masterversion.ini"
    
    ' read the version count
    VersionCount = Val(GetVar(App.Path & "\masterversion.ini", "MASTER", "Atualizações"))
    
    ' check if we've got a current client version saved
    If FileExist(App.Path & "\data files\config.ini", True) Then
        CurVersion = Val(GetVar(App.Path & "\data files\config.ini", "CLIENTE", "versão"))
    Else
        CurVersion = 0
    End If
    
    ' are we up to date?
    If CurVersion < VersionCount Then
        frmUpdater.Show
        ' make sure it's not 0!
        If CurVersion = 0 Then CurVersion = 1
        ' loop around, download and unrar each update
        For i = CurVersion To VersionCount
            ' let them know!
            AddProgress printf("Baixando atualização %d.", Val(i))
            filename = "universoz" & i & ".rar"
            ' set the download going through inet
            DownloadFile UpdateURL & "/" & filename, App.Path & "\" & filename
            ' us the unrar.dll to extract data
            RARExecute OP_EXTRACT, filename
            ' kill the temp update file
            Kill App.Path & "\" & filename
            ' update the current version
            PutVar App.Path & "\data files\config.ini", "CLIENTE", "versão", str(i)
            ' let them know!
            AddProgress printf("Versão %d instalada.", Val(i))
        Next
        Unload frmUpdater
    End If
    End If
End Sub

Private Sub AddProgress(ByVal Progress As String)
    frmUpdater.Label2.Caption = Progress
End Sub

Private Sub DownloadFile(ByVal URL As String, ByVal filename As String)
    Dim fileBytes() As Byte
    Dim fileNum As Integer
    
    On Error GoTo DownloadError
    
    ' download data to byte array
    fileBytes() = frmUpdater.inetDownload.OpenURL(URL, icByteArray)
    
    fileNum = FreeFile
    Open filename For Binary Access Write As #fileNum
        ' dump the byte array as binary
        Put #fileNum, , fileBytes()
    Close #fileNum
    
    Exit Sub
    
DownloadError:
    MsgBox Err.Description
End Sub
