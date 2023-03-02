VERSION 5.00
Begin VB.Form frmMapReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Report"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdWarp 
      Caption         =   "Teleport"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   2
      Top             =   4200
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4395
      Begin VB.ListBox lstMaps 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3480
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4155
      End
   End
End
Attribute VB_Name = "frmMapReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdWarp_Click()
    Call WarpTo(lstMaps.ListIndex + 1)
End Sub

Private Sub lstMaps_DblClick()
    Call WarpTo(lstMaps.ListIndex + 1)
End Sub
