Attribute VB_Name = "modGFXPack"
Sub HandleGFXPack(MapNpcNum As Long)
    Select Case Npc(MapNpc(MapNpcNum).num).GFXPack
        Case 1 'Vegeta
            Call DrawVegeta(MapNpcNum)
    End Select
End Sub

Sub DrawVegeta(ByVal MapNpcNum As Long)
    Dim Anim As Byte, i As Long, X As Long, Y As Long, Sprite As Long, spritetop As Long
    Dim rec As RECT
    Dim AttackSpeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug >= 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).num = 0 Then Exit Sub ' no npc set
    
    Sprite = Npc(MapNpc(MapNpcNum).num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    AttackSpeed = Npc(MapNpc(MapNpcNum).num).AttackSpeed
    
    If AttackSpeed < 100 Then AttackSpeed = 100

    ' Reset frame
    Anim = 0
    
    ' Check for attacking animation
    If TempMapNpc(MapNpcNum).AttackTimer + (AttackSpeed / 2) > GetTickCount Then
        If TempMapNpc(MapNpcNum).Attacking = 1 Then
            If TempMapNpc(MapNpcNum).AttackType = 0 Then
                If TempMapNpc(MapNpcNum).AttackTimer + (AttackSpeed / 8) < GetTickCount Then
                    Anim = 5 + (2 * TempMapNpc(MapNpcNum).AttackData1)
                Else
                    Anim = 4 + (2 * TempMapNpc(MapNpcNum).AttackData1)
                End If
            Else
                If TempMapNpc(MapNpcNum).AttackTimer + (AttackSpeed / 8) < GetTickCount Then
                    Anim = 9
                Else
                    Anim = 8
                End If
            End If
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (TempMapNpc(MapNpcNum).YOffSet > 8) Then Anim = TempMapNpc(MapNpcNum).Step
            Case DIR_DOWN
                If (TempMapNpc(MapNpcNum).YOffSet < -8) Then Anim = TempMapNpc(MapNpcNum).Step
            Case DIR_LEFT
                If (TempMapNpc(MapNpcNum).XOffSet > 8) Then Anim = TempMapNpc(MapNpcNum).Step
            Case DIR_RIGHT
                If (TempMapNpc(MapNpcNum).XOffSet < -8) Then Anim = TempMapNpc(MapNpcNum).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With TempMapNpc(MapNpcNum)
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (Tex_Character(Sprite).Height / 4) * spritetop
        .Bottom = .Top + Tex_Character(Sprite).Height / 4
        .Left = Anim * (Tex_Character(Sprite).Width / 10)
        .Right = .Left + (Tex_Character(Sprite).Width / 10)
    End With

    ' Calculate the X
    If VXFRAME = False Then
        X = MapNpc(MapNpcNum).X * PIC_X + TempMapNpc(MapNpcNum).XOffSet - ((Tex_Character(Sprite).Width / 10 - 32) / 2)
    Else
        X = MapNpc(MapNpcNum).X * PIC_X + TempMapNpc(MapNpcNum).XOffSet - ((Tex_Character(Sprite).Width / 3 - 32) / 2)
    End If
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + TempMapNpc(MapNpcNum).YOffSet - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + TempMapNpc(MapNpcNum).YOffSet
    End If
    
    ' render player shadow
    If Npc(MapNpc(MapNpcNum).num).Shadow = 1 Then RenderTexture Tex_Shadow, ConvertMapX(X) - (MapNpc(MapNpcNum).FlyOffSet / 2), ConvertMapY(Y + 18), 0, 0, (Tex_Character(Sprite).Width / 4) + MapNpc(MapNpcNum).FlyOffSet, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    
    If Npc(MapNpc(MapNpcNum).num).Fly = 1 Then
        Y = Y + MapNpc(MapNpcNum).FlyOffSet
    End If
    
    ' render the actual sprite
    If GetTickCount > TempMapNpc(MapNpcNum).StartFlash Then
        If TempMapNpc(MapNpcNum).StunDuration > 0 Then
            If GetTickCount < TempMapNpc(MapNpcNum).StunTick + (TempMapNpc(MapNpcNum).StunDuration * 1000) Then
                Call DrawSprite(Sprite, X, Y, rec, True)
                Exit Sub
            End If
        End If
        Call DrawSprite(Sprite, X, Y, rec)
        TempMapNpc(MapNpcNum).StartFlash = 0
    Else
        Call DrawSprite(Sprite, X, Y, rec, True)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpc", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
