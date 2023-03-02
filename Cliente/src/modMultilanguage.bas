Attribute VB_Name = "modMultilanguage"
Private Messages() As String
Private Translates() As String

Public Sub LoadLanguage()
    On Error GoTo ErrHandler
    If Options.Language <> "pt" Then
        Dim i As Long, lineCount As Long
        
        'Load message base
        lineCount = 1
        Do While GetVar(App.Path & "\data files\lang\pt.ini", "LANGUAGE", "Msg" & lineCount) <> ""
            lineCount = lineCount + 1
        Loop
        lineCount = lineCount - 1
        
        ReDim Messages(1 To lineCount) As String
        For i = 1 To lineCount
            Messages(i) = GetVar(App.Path & "\data files\lang\pt.ini", "LANGUAGE", "Msg" & i)
        Next i
        
        'Load translation
        lineCount = 1
        Do While GetVar(App.Path & "\data files\lang\" & Options.Language & ".ini", "LANGUAGE", "Msg" & lineCount) <> ""
            lineCount = lineCount + 1
        Loop
        lineCount = lineCount - 1
        
        ReDim Translates(1 To lineCount) As String
        For i = 1 To lineCount
            Translates(i) = GetVar(App.Path & "\data files\lang\" & Options.Language & ".ini", "LANGUAGE", "Msg" & i)
        Next i
        
    End If
    Exit Sub
ErrHandler:
        MsgBox "Language loading error!"
        Call DestroyGame
End Sub

Public Function printf(sTemplate As String, Optional sArgs As String)
    Dim sRet As String, lPoint As Long, lArg As Long, aArg() As String
    sRet = sTemplate
    
    If Options.Language <> "pt" Then
        sRet = Translate(sRet)
    End If
    
    If sArgs <> "" Then
        aArg = Split(sArgs, ",")
        lPoint = InStr(1, sRet, "%")
        While lPoint > 0
            Select Case Mid(sRet, lPoint + 1, 1)
                Case "%"
                    sRet = Left(sRet, lPoint - 1) & Mid(sRet, lPoint + 1)
                Case "d"
                    sRet = Left(sRet, lPoint - 1) & CStr(CLng(aArg(lArg))) & Mid(sRet, lPoint + 2)
                    lArg = lArg + 1
                Case "s"
                    sRet = Left(sRet, lPoint - 1) & aArg(lArg) & Mid(sRet, lPoint + 2)
                    lArg = lArg + 1
            End Select
            lPoint = InStr(lPoint + 1, sRet, "%")
        Wend
        sRet = Replace(sRet, "\n", vbLf)
        sRet = Replace(sRet, "\r", vbCr)
    End If
    
    printf = sRet
End Function

Function Translate(ByVal Msg As String) As String
    Dim i As Long
    Translate = Msg
    For i = 1 To UBound(Messages)
        If Msg = Messages(i) Then
            Translate = Translates(i)
            Exit Function
        End If
    Next i
    Call SetStatus("Error! Unrecognized translation: " & Msg)
End Function
