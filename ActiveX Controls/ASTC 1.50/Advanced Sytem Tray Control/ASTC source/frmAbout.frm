VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Smack 1.0"
   ClientHeight    =   5640
   ClientLeft      =   4980
   ClientTop       =   3210
   ClientWidth     =   5445
   ClipControls    =   0   'False
   FillColor       =   &H00FF0000&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":08CA
   ScaleHeight     =   3892.829
   ScaleMode       =   0  'User
   ScaleWidth      =   5113.137
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5040
      Width           =   1260
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3600
      Picture         =   "frmAbout.frx":B6298
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   1440
      Width           =   480
   End
   Begin VB.PictureBox Store1 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   1440
      Picture         =   "frmAbout.frx":B6EDA
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
      Index           =   0
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You will need to purchase  a developers licence to stop this message appearing while your application runs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   2880
      Width           =   4380
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLicenseTo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "unregistered version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0E0FF&
      BorderWidth     =   2
      FillColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   240
      Top             =   240
      Width           =   4935
   End
   Begin VB.Image Image5 
      Height          =   1080
      Left            =   1200
      Picture         =   "frmAbout.frx":B71E4
      Top             =   240
      Width           =   2805
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   1440
      Picture         =   "frmAbout.frx":C10C6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2445
   End
   Begin VB.Image Image4 
      Height          =   1215
      Left            =   240
      Picture         =   "frmAbout.frx":CD0B4
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3420
   End
   Begin VB.Image Image3 
      Height          =   1215
      Left            =   2520
      Picture         =   "frmAbout.frx":DF432
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2625
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ASTC 1.0"
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
      Left            =   1200
      TabIndex        =   10
      Top             =   1560
      Width           =   2475
   End
   Begin VB.Label Label6 
      Caption         =   "use gold rect"
      Height          =   855
      Left            =   6600
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   5400
      Left            =   120
      Top             =   120
      Width           =   5205
   End
   Begin VB.Label Label2 
      Caption         =   "companyy logo"
      Height          =   495
      Left            =   7680
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":F17B0
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   4958.193
      Y1              =   3395.872
      Y2              =   3395.872
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced System Tray Icon control "
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
      Height          =   450
      Left            =   600
      TabIndex        =   0
      Top             =   2160
      Width           =   3525
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
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.cpringold.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1560
      MouseIcon       =   "frmAbout.frx":F1860
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdOK_Click(index As Integer)
 Unload Me
End Sub

Private Sub Command1_Click()

ShellExecute Me.hwnd, vbNullString, "http://www.cpringold.com/products/astc.html", vbNullString, "C:\", SW_SHOWNORMAL


End Sub

Private Sub Form_Load()
    
    Me.Caption = "About " & App.Title & " 1.0"
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    lblTitle.Caption = App.Title
    
End Sub

Private Sub Label5_Click()

ShellExecute Me.hwnd, vbNullString, "mailto:support@cpringold.com", vbNullString, "C:\", SW_SHOWNORMAL

End Sub



Private Sub Label3_Click()

ShellExecute Me.hwnd, vbNullString, "http://cpringold.com", vbNullString, "C:\", SW_SHOWNORMAL


End Sub
