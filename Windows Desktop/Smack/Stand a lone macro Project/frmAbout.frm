VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4905
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
   ScaleHeight     =   3385.519
   ScaleMode       =   0  'User
   ScaleWidth      =   5056.793
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Buy Now"
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
      TabIndex        =   4
      Top             =   4320
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   4320
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   360
      Picture         =   "frmAbout.frx":B65A2
      Top             =   1440
      Width           =   4680
   End
   Begin VB.Label lblLicenseTo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " This is an unregistered version for evaluation use only."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Executable Macro File"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1125
      Left            =   1800
      MouseIcon       =   "frmAbout.frx":B8522
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":B8674
      Top             =   2280
      Width           =   1785
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
      Left            =   1200
      TabIndex        =   9
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "use gold rect"
      Height          =   855
      Left            =   6600
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   4680
      Left            =   120
      Top             =   120
      Width           =   5205
   End
   Begin VB.Label Label2 
      Caption         =   "companyy logo"
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":B91F3
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   4958.192
      Y1              =   2898.915
      Y2              =   2898.915
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
      TabIndex        =   2
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
      TabIndex        =   3
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
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOK_Click()
  
 End
  
End Sub

Private Sub Command1_Click()
ShellExecute Me.Hwnd, vbNullString, "http://www.cpringold.com/buy/ordersmack.html", vbNullString, "", SW_SHOWNORMAL

End Sub

Private Sub Form_Load()
    
    Me.Caption = App.Title
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    lblTitle.Caption = App.Title
    
    
End Sub

Private Sub Image2_Click()

ShellExecute Me.Hwnd, vbNullString, "http://www.cpringold.com/buy/ordersmack.html", vbNullString, "", SW_SHOWNORMAL
End Sub

Private Sub Label3_Click()

ShellExecute Me.Hwnd, vbNullString, "http://www.cpringold.com", vbNullString, "", SW_SHOWNORMAL

End Sub

Private Sub Label5_Click()

ShellExecute Me.Hwnd, vbNullString, "mailto:support@cpringold.com", vbNullString, "", SW_SHOWNORMAL

End Sub



