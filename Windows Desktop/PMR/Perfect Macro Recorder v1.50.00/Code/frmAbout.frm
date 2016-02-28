VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2265
   ClientLeft      =   4980
   ClientTop       =   3210
   ClientWidth     =   6825
   ClipControls    =   0   'False
   FillColor       =   &H00FF0000&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1563.343
   ScaleMode       =   0  'User
   ScaleWidth      =   6409.026
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6615
      Begin VB.Image Image1 
         Height          =   720
         Left            =   240
         Picture         =   "frmAbout.frx":6852
         Top             =   240
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   5760
         Picture         =   "frmAbout.frx":D0A4
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Perfection Tools Software Support Email  Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         MouseIcon       =   "frmAbout.frx":138F6
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   " © Copyright 2007, Perfection Tools Software. All rights reserved."
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Perfection Tools Software Homepage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2160
         MouseIcon       =   "frmAbout.frx":13A48
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Perfect Macro Recorder 1.50"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   2835
      End
   End
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox Store1 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   1440
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   540
   End
   Begin JwldButn2b.JeweledButton cmdOK 
      Height          =   375
      Left            =   2798
      TabIndex        =   16
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "O&K"
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
      Left            =   3720
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   765
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
      Left            =   840
      TabIndex        =   9
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label6 
      Caption         =   "use gold rect"
      Height          =   855
      Left            =   6240
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "companyy logo"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "support@cpringold.com"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   6240
      Width           =   4215
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
      Left            =   480
      TabIndex        =   0
      Top             =   7200
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
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
  
  Unload Me

End Sub

Private Sub Form_Load()
    
    'cprin 26July2006: added next code block
    'to ensure that jeweled buttons will remain
    'as i want.
    ''''''''''''''''''''''''''''''''''''''''''''
    cmdOk.Caption = "O&K"
    
    Me.Caption = "About " & App.Title
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    lblTitle.Caption = App.Title
    
  ' Label8.Caption = "© Copyright 2007 - " & Year(Date) & ", Perfection Tools Software. All rights reserved."
    
End Sub




Private Sub Label3_Click()

ShellExecute Me.hWnd, vbNullString, SiteName, vbNullString, "", SW_SHOWNORMAL

End Sub

Private Sub Label5_Click()

ShellExecute Me.hWnd, vbNullString, "mailto:" & Email, vbNullString, "", SW_SHOWNORMAL

End Sub

Private Sub Label9_Click()
    
    ShellExecute Me.hWnd, vbNullString, "mailto:" & Email, vbNullString, "", SW_SHOWNORMAL

End Sub
