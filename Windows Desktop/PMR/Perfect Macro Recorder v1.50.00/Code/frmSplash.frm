VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   150
   ClientLeft      =   4875
   ClientTop       =   3795
   ClientWidth     =   90
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151.24
   ScaleMode       =   0  'User
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   850
      Left            =   600
      Top             =   5280
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   4560
   End
   Begin VB.Image Image5 
      Height          =   885
      Left            =   3360
      Top             =   840
      Width           =   1065
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   360
      Top             =   3960
      Width           =   3960
   End
   Begin VB.Image Image3 
      Height          =   1095
      Left            =   120
      Top             =   3720
      Width           =   2130
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CprinGold Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   5880
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   1080
      Top             =   3960
      Width           =   2805
   End
   Begin VB.Image imgLogo 
      Height          =   705
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   675
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3000
      TabIndex        =   3
      Top             =   5880
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   3495
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CprinGold Software"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   1200
      TabIndex        =   2
      Top             =   5040
      Width           =   3120
   End
   Begin VB.Label lblLicenseTo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "This Product is  Licensed for Single User Only."
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
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Smack 1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   3720
      Width           =   765
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub Form_Load()

Dim strUserName As String

On Error Resume Next

Set Tags = New CTags

lblCompanyProduct.ForeColor = RGB(255, 215, 0)

If App.PrevInstance Then MsgBox SoftwareName & " is already running": End

strUserName = String(100, Chr$(0))

GetUserName strUserName, 100
    
  'cprin 13July2006: comment next line to disable the Splash at the StartUp
  '  If Command$ = "hide" Then
       
       Me.Visible = False
       
       Timer1.Enabled = False
       
       Unload frmSplash
       
       Exit Sub
  '  End If

lblLicenseTo.Caption = lblLicenseTo & " " & Trim(strUserName)

End Sub

Private Sub Form_Unload(cancel As Integer)
  
  Dim IsFirstTime As String
  
  Dim Res As Integer

  Res = GetINIProfileInt("chkMinStartUp", 0)

  If Command$ = "hide" Or Res Then
      
      Load frmJMacro
   
   Else
      
      IsFirstTime = GetSetting(App.EXEName, "srtew", "k", 0)
          
      If IsFirstTime = 0 Then
          '//cprin 19July2006: added next line to ensure if this is the first time that Perfect Macro Recorder Running
          '//we will use to enable and disable the balloons tooltips
          IsBalloonTooltipsEnabled = True
          SaveSetting App.EXEName, "srtew", "k", 1
          Load frmJMacro
          frmJMacro.tmrTransperent.Enabled = True
          Exit Sub
      End If
      
      frmJMacro.Show
   
   End If
   
End Sub

Private Sub Timer1_Timer()

On Error Resume Next

Me.Hide

Timer2.Enabled = True

End Sub

Private Sub Timer2_Timer()

Timer2.Enabled = False

Unload Me

End Sub
