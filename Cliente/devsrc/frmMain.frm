VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   630
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picAdminBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11970
      TabIndex        =   0
      Top             =   8640
      Width           =   12000
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   9480
         TabIndex        =   6
         Top             =   50
         Width           =   1335
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   9360
         X2              =   9360
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Label lblNames 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hide Names"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8040
         TabIndex        =   5
         Top             =   50
         Width           =   1335
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   7920
         X2              =   7920
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Label lblCollision 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Disable collisions"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   4
         Top             =   45
         Width           =   1935
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   5880
         X2              =   5880
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Show Info"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   45
         Width           =   1575
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   4200
         X2              =   4200
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Label lblGUI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hide GUI"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   45
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   2520
         X2              =   2520
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Label lblAdminPanel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Open Admin"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   45
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long

Private Sub Form_DblClick()
    HandleDoubleClick
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseDown Button
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleMouseUp Button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Cancel = True
    logoutGame
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    HandleMouseMove CLng(X), CLng(Y), Button
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Label1_Click()
    picAdminBack.visible = False
End Sub

Private Sub lblCollision_Click()
    If Val(lblCollision.Tag) = 0 Then
        lblCollision.Tag = 1
        lblCollision.Caption = "Activate collisions"
        CollisionDisabled = True
    Else
        lblCollision.Tag = 0
        lblCollision.Caption = "Disable collisions"
        CollisionDisabled = False
    End If
End Sub

Private Sub lblAdminPanel_Click()
    If Val(lblAdminPanel.Tag) = 0 Then
        lblAdminPanel.Tag = 1
        lblAdminPanel.Caption = "Close Admin"
        frmAdminPanel.visible = True
    Else
        lblAdminPanel.Tag = 0
        lblAdminPanel.Caption = "Open Admin"
        frmAdminPanel.visible = False
    End If
End Sub

Private Sub lblGUI_Click()
    If Val(lblGUI.Tag) = 0 Then
        lblGUI.Tag = 1
        lblGUI.Caption = "Show GUI"
        hideGUI = True
    Else
        lblGUI.Tag = 0
        lblGUI.Caption = "Hide GUI"
        hideGUI = False
    End If
End Sub

Private Sub lblInfo_Click()
    If Val(lblInfo.Tag) = 0 Then
        lblInfo.Tag = 1
        lblInfo.Caption = "Hide Info"
        BFPS = True
    Else
        lblInfo.Tag = 0
        lblInfo.Caption = "Show Info"
        BFPS = False
    End If
End Sub

Private Sub lblNames_Click()
    If Val(lblNames.Tag) = 0 Then
        lblNames.Tag = 1
        lblNames.Caption = "Show Names"
        hideNames = True
    Else
        lblNames.Tag = 0
        lblNames.Caption = "Hide Names"
        hideNames = False
    End If
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If KeyCode = vbKeyF1 And Shift = 1 Then
        picAdminBack.visible = True
        Exit Sub
    End If
    
    HandleKeyUp KeyCode

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
