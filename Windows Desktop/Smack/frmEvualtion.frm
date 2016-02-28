VERSION 5.00
Begin VB.Form frmEvualtion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smack 1.06 Demo version"
   ClientHeight    =   4665
   ClientLeft      =   3000
   ClientTop       =   2280
   ClientWidth     =   7065
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmEvualtion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "help"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Order Now!"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "If you purchase Smack"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   6855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEvualtion.frx":08CA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2340
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   6795
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   2415
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEvualtion.frx":0A28
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   4200
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmEvualtion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
ShellExecute Me.hWnd, vbNullString, "http://www.cpringold.com/buy/ordersmack.html", vbNullString, "", SW_SHOWNORMAL
  Command4_Click
End Sub

Private Sub Command2_Click()
Shell "hh " & AppPath & "smart.chm", vbMaximizedFocus
End Sub

Private Sub Command3_Click()

End

End Sub

Private Sub Command4_Click()
 If Command$ = "hide" Then
      
      Load frmJMacro
   
   Else
   
       frmJMacro.Show
   
   End If

Tags.OkEvual = True

Unload Me

End Sub

