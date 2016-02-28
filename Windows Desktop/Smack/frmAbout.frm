VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4035
   ClientLeft      =   4980
   ClientTop       =   3210
   ClientWidth     =   5385
   ClipControls    =   0   'False
   FillColor       =   &H00FF0000&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":08CA
   ScaleHeight     =   2785.029
   ScaleMode       =   0  'User
   ScaleWidth      =   5056.793
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Order "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox Store1 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   1440
      Picture         =   "frmAbout.frx":B6298
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "O&K"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1260
   End
   Begin VB.Label lblLicenseTo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "This Product is  Licensed."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Smack 1.06"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   3435
   End
   Begin VB.Label Label6 
      Caption         =   "use gold rect"
      Height          =   855
      Left            =   6600
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   3840
      Left            =   120
      Top             =   120
      Width           =   5205
   End
   Begin VB.Label Label2 
      Caption         =   "companyy logo"
      Height          =   495
      Left            =   7680
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "support@cpringold.com"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmAbout.frx":B65A2
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":B68AC
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   4958.192
      Y1              =   2319.132
      Y2              =   2319.132
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Save your time with Smack by Record your Keyboard and Mouse actions for playback later."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   4485
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Smack 1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   600
      TabIndex        =   3
      Top             =   5760
      Width           =   2115
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8040
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CprinGold Software"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1680
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.cpringold.com"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmAbout.frx":B695C
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  
  Unload Me

End Sub

Private Sub Form_Load()
    
    Me.caption = "About " & App.Title
    
    lblVersion.caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    lblTitle.caption = App.Title
    
    
End Sub




Private Sub Label3_Click()

ShellExecute Me.hWnd, vbNullString, "http://www.cpringold.com", vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub Label5_Click()

ShellExecute Me.hWnd, vbNullString, "mailto:support@cpringold.com", vbNullString, "C:\", SW_SHOWNORMAL

End Sub



