VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About "
   ClientHeight    =   3480
   ClientLeft      =   3420
   ClientTop       =   2595
   ClientWidth     =   5205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2401.958
   ScaleMode       =   0  'User
   ScaleWidth      =   4887.763
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1260
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "http://cpringold.atspace.com"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   2253.724
      X2              =   225.372
      Y1              =   1822.175
      Y2              =   1822.175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Smack Executable Macro File"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Smack 1.0"
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
      TabIndex        =   7
      Top             =   240
      Width           =   2475
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You will need to purchase licence to stop this message."
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
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLicenseTo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "from appearing"
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
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "use gold rect"
      Height          =   855
      Left            =   6600
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   3240
      Left            =   120
      Top             =   120
      Width           =   4965
   End
   Begin VB.Label Label2 
      Caption         =   "companyy logo"
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   4732.819
      Y1              =   1905.002
      Y2              =   1905.002
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
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOK_Click(Index As Integer)
Load Form1
End Sub

Private Sub Label5_Click()
ShellExecute Me.Hwnd, vbNullString, "http://cpringold.atspace.com", vbNullString, "C:\", 1
End Sub
