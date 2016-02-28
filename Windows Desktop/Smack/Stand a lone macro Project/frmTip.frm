VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   960
   ClientLeft      =   3990
   ClientTop       =   4380
   ClientWidth     =   4050
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDontShow 
      Caption         =   "&Don't show this again"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   200
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't show this again"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblTip 
      BackStyle       =   0  'Transparent
      Caption         =   "Press Ctrl+Esc to Terminate playback"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Label1_Click()

If chkDontShow.Value = 1 Then chkDontShow.Value = 0: Exit Sub

If chkDontShow.Value = 0 Then chkDontShow.Value = 1

End Sub



Private Sub OKButton_Click()

If chkDontShow.Value = 1 Then

   SaveSetting App.EXEName, "k", "d", "one"

End If

Unload Me

End Sub
