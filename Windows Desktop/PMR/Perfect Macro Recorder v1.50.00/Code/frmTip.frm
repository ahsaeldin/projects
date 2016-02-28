VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmTip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1920
   ClientLeft      =   3990
   ClientTop       =   4380
   ClientWidth     =   4050
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3855
      Begin VB.Label lblTip 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
   End
   Begin JwldButn2b.JeweledButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&OK"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BackColor       =   12632256
   End
   Begin VB.CheckBox chkDontShow 
      Caption         =   "&Don't show this again"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't show this again"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
'cprin 19July2006: added next code block
'to ensure that jeweled buttons will remain
'as i want.
''''''''''''''''''''''''''''''''''''''''''''
OKButton.Caption = "&OK"
End Sub

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
