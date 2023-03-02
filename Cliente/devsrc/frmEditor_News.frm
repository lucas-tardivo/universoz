VERSION 5.00
Begin VB.Form frmEditor_News 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "News Editor"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Welcome message"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtNews 
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmEditor_News"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload frmEditor_News
End Sub

Private Sub cmdSave_Click()
    SendEditNews txtNews.Text
    Unload frmEditor_News
End Sub
