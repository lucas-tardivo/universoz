VERSION 5.00
Begin VB.Form frmPaperdoll 
   Caption         =   "Paperdoll"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CheckBox chkPlayer 
      Caption         =   "Player"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.TextBox txtHair 
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   4455
   End
   Begin VB.TextBox txtHair 
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   4455
   End
   Begin VB.TextBox txtHair 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   4455
   End
   Begin VB.TextBox txtHair 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   4455
   End
   Begin VB.TextBox txtHair 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   4455
   End
   Begin VB.TextBox txtHair 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   4455
   End
   Begin VB.ListBox lstAdjust 
      Height          =   2205
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4455
   End
   Begin VB.HScrollBar scrlY 
      Height          =   255
      Left            =   840
      Max             =   32
      Min             =   -32
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.HScrollBar scrlX 
      Height          =   255
      Left            =   840
      Max             =   48
      Min             =   -48
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmPaperdoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i As Long
    Call PutVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "GlobalX", Val(GlobalRepositionX))
    Call PutVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "GlobalY", Val(GlobalRepositionY))
    For i = 1 To 6
        Call PutVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "LineX" & Val(i), Val(HairRepositionX(i)))
        Call PutVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "LineY" & Val(i), Val(HairRepositionY(i)))
    Next i
    For i = 1 To 84
        Call PutVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "HairX" & Val(i), Val(PositionRepositionX(i)))
        Call PutVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "HairY" & Val(i), Val(PositionRepositionY(i)))
    Next i
End Sub

Private Sub Command2_Click()
    Dim i As Long
    GlobalRepositionX = GetVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "GlobalX")
    GlobalRepositionY = GetVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "GlobalY")
    For i = 1 To 6
        HairRepositionX(i) = GetVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "LineX" & Val(i))
        HairRepositionY(i) = GetVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "LineY" & Val(i))
    Next i
    For i = 1 To 84
        PositionRepositionX(i) = GetVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "HairX" & Val(i))
        PositionRepositionX(i) = GetVar(App.Path & "\data files\graphics\adm\hairrepos.ini", "HAIR", "HairY" & Val(i))
    Next i
End Sub

Private Sub Form_Load()
    lstAdjust.Clear
    lstAdjust.AddItem "Global"
    Dim i As Long
    For i = 1 To 6
        lstAdjust.AddItem "Line " & i
        txtHair(i - 1) = GetVar(App.Path & "\data files\graphics\adm\hairpositions.ini", "HAIR", Val(i))
    Next i
    For i = 1 To 84
        lstAdjust.AddItem "Position " & i
    Next i
    
End Sub

Private Sub HScroll1_Change()
    
End Sub

Private Sub lstAdjust_Click()
    If lstAdjust.ListIndex = 0 Then
        scrlX.Value = GlobalRepositionX
        scrlY.Value = GlobalRepositionY
    Else
        If lstAdjust.ListIndex <= 6 Then
            scrlX.Value = HairRepositionX(lstAdjust.ListIndex)
            scrlY.Value = HairRepositionY(lstAdjust.ListIndex)
        Else
            scrlX.Value = PositionRepositionX(lstAdjust.ListIndex - 6)
            scrlY.Value = PositionRepositionY(lstAdjust.ListIndex - 6)
        End If
    End If
End Sub

Private Sub lstAdjust_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 65 Then
        scrlX.Value = scrlX.Value - 1
    End If
    If KeyCode = 68 Then
        scrlX.Value = scrlX.Value + 1
    End If
    If KeyCode = 83 Then
        scrlY.Value = scrlY.Value + 1
    End If
    If KeyCode = 87 Then
        scrlY.Value = scrlY.Value - 1
    End If
    If KeyCode = 13 Then lstAdjust.ListIndex = lstAdjust.ListIndex + 1
End Sub

Private Sub scrlX_Change()
    On Error Resume Next
    If lstAdjust.ListIndex = 0 Then
        GlobalRepositionX = scrlX.Value
    Else
        If lstAdjust.ListIndex <= 6 Then
            HairRepositionX(lstAdjust.ListIndex) = scrlX.Value
        Else
            PositionRepositionX(lstAdjust.ListIndex - 6) = scrlX.Value
        End If
    End If
End Sub

Private Sub scrlY_Change()
    On Error Resume Next
    If lstAdjust.ListIndex = 0 Then
        GlobalRepositionY = scrlY.Value
    Else
        If lstAdjust.ListIndex <= 6 Then
            HairRepositionY(lstAdjust.ListIndex) = scrlY.Value
        Else
            PositionRepositionY(lstAdjust.ListIndex - 6) = scrlY.Value
        End If
    End If
End Sub

Private Sub txtHair_Change(Index As Integer)
    Call PutVar(App.Path & "\data files\graphics\adm\hairpositions.ini", "HAIR", Val(Index + 1), txtHair(Index).Text)
End Sub
