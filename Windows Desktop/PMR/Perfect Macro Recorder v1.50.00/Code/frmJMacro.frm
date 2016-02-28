VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Object = "{30E5838E-8E26-11DA-A1C2-807A50C10000}#1.0#0"; "acomte445.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmJMacro 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfect Macro Recorder 1.50"
   ClientHeight    =   3210
   ClientLeft      =   5055
   ClientTop       =   3195
   ClientWidth     =   6810
   Icon            =   "frmJMacro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrTerminatePlayback 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7320
      Top             =   720
   End
   Begin VB.TextBox txtRepeat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   21
      ToolTipText     =   "Set the Repeat count of playback."
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox PIco 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   8880
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin JwldButn2b.JeweledButton cmdHelp 
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Help"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin ASTC.Balloon ExeBalloon 
      Left            =   9000
      Top             =   2040
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon BackAllBalloon 
      Left            =   9000
      Top             =   1560
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon BackupMacro 
      Left            =   9000
      Top             =   1080
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon ImportBalloon 
      Left            =   9000
      Top             =   600
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon OptionBalllon 
      Left            =   9000
      Top             =   120
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon MyMacrosBalloon 
      Left            =   8400
      Top             =   2400
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.SubClass SubClassCombo 
      Left            =   10080
      Top             =   120
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin ASTC.Balloon ComboBalloon 
      Left            =   8400
      Top             =   1920
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon ContinueBallon 
      Left            =   8400
      Top             =   1560
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon PlayBackBalloon 
      Left            =   8400
      Top             =   1080
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon EditBalloon 
      Left            =   8400
      Top             =   600
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon RecordBalloon 
      Left            =   8400
      Top             =   120
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   6375
         Begin JwldButn2b.JeweledButton cmdStandaLone 
            Height          =   495
            Left            =   1680
            TabIndex        =   6
            ToolTipText     =   "Click to convert selected macro in the combo box to an EXE-file"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "E&xport to exe file"
            PictureSize     =   0
            OriginalPicSizeW=   0
            OriginalPicSizeH=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            BackColor       =   12632256
            BorderColor_Outer=   8421504
         End
         Begin JwldButn2b.JeweledButton cmdoptions 
            Height          =   495
            Left            =   3240
            TabIndex        =   10
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&Options"
            PictureSize     =   0
            OriginalPicSizeW=   0
            OriginalPicSizeH=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            BackColor       =   12632256
            BorderColor_Outer=   8421504
         End
         Begin JwldButn2b.JeweledButton cmdmanage 
            Height          =   495
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Manage the macros you has created"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&My Macros"
            PictureSize     =   0
            OriginalPicSizeW=   0
            OriginalPicSizeH=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            BackColor       =   12632256
            BorderColor_Outer=   8421504
         End
         Begin JwldButn2b.JeweledButton cmdImport 
            Height          =   495
            Left            =   3240
            TabIndex        =   7
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&Import"
            PictureSize     =   0
            OriginalPicSizeW=   0
            OriginalPicSizeH=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            BackColor       =   12632256
            BorderColor_Outer=   8421504
         End
         Begin JwldButn2b.JeweledButton cmdExport 
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&Backup Macro"
            PictureSize     =   0
            OriginalPicSizeW=   0
            OriginalPicSizeH=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            BackColor       =   12632256
            BorderColor_Outer=   8421504
         End
         Begin JwldButn2b.JeweledButton cmdBackUpAll 
            Height          =   495
            Left            =   1680
            TabIndex        =   9
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "Backup &All"
            PictureSize     =   0
            OriginalPicSizeW=   0
            OriginalPicSizeH=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            BackColor       =   12632256
            BorderColor_Outer=   8421504
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   6375
         Begin VB.ComboBox CobMacros 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   1935
         End
         Begin JwldButn2b.JeweledButton cmdRecord 
            Default         =   -1  'True
            Height          =   495
            Left            =   120
            TabIndex        =   1
            ToolTipText     =   "Click to Start Recording a new macro"
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&Record"
            PictureSize     =   0
            OriginalPicSizeW=   0
            OriginalPicSizeH=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
         Begin JwldButn2b.JeweledButton cmdEditMacro 
            Height          =   495
            Left            =   4800
            TabIndex        =   4
            ToolTipText     =   "Edit the selected macro in the combo box"
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&Edit Macro"
            PictureSize     =   0
            OriginalPicSizeW=   0
            OriginalPicSizeH=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
         Begin JwldButn2b.JeweledButton cmdContinue 
            Height          =   495
            Left            =   3240
            TabIndex        =   3
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "Contin&ue"
            PictureWidth    =   48
            PictureHeight   =   48
            PictureSize     =   2
            OriginalPicSizeW=   0
            OriginalPicSizeH=   0
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
         Begin JwldButn2b.JeweledButton cmdPlayBack 
            Height          =   495
            Left            =   1680
            TabIndex        =   2
            ToolTipText     =   "Click to play back the selected macro in the Combo Box"
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "&Play back"
            PictureSize     =   0
            OriginalPicSizeW=   0
            OriginalPicSizeH=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
      End
   End
   Begin ASTC.TrayIcon TrayIcon1 
      Left            =   9480
      Top             =   120
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Text            =   "fade in first time only"
      Top             =   10800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer tmrTransperent 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   12240
      Top             =   120
   End
   Begin VB.CommandButton cmdAliasRecord 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Record New Macro"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Record your Mouse  and keyboards events"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdLanguage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Language"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJMacro.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJMacro.frx":1264
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJMacro.frx":1F3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrCheckShortKeys 
      Interval        =   1
      Left            =   11880
      Top             =   120
   End
   Begin JwldButn2b.JeweledButton cmdreviews 
      Height          =   495
      Left            =   9120
      TabIndex        =   19
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Exit"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BorderColor_Outer=   8421504
   End
   Begin JwldButn2b.JeweledButton cmdGetNow 
      Height          =   495
      Left            =   9360
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&GetNow!"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BorderColor_Outer=   8421504
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Repeat Playback Count"
      Height          =   255
      Left            =   7200
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      Visible         =   0   'False
      X1              =   4080
      X2              =   4320
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Menu JMacroMenu 
      Caption         =   "Perfect Macro Recorder"
      Visible         =   0   'False
      Begin VB.Menu mnu_jmacro 
         Caption         =   "Hide Perfect Macro Recorder"
      End
      Begin VB.Menu a 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuRecord 
         Caption         =   "Record New Macro"
      End
      Begin VB.Menu mnuPlayback 
         Caption         =   "&Run Macro"
         Enabled         =   0   'False
         Begin VB.Menu menuSavedMacros 
            Caption         =   "Available Macros"
            Index           =   0
         End
      End
      Begin VB.Menu uu 
         Caption         =   "-"
      End
      Begin VB.Menu mnumanage 
         Caption         =   "My Macros"
      End
      Begin VB.Menu mnuDisableAllShortcutKeys 
         Caption         =   "Disable All Shortcut Keys"
      End
      Begin VB.Menu mnuoptions 
         Caption         =   "Perfect Macro Recorder Options"
      End
      Begin VB.Menu af 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelp2 
         Caption         =   "Perfect Macro Recorder Help"
      End
      Begin VB.Menu mnuRegistrationKey 
         Caption         =   "Enter Registration Key..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuweb 
         Caption         =   "Perfection Tools on the web"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu vismnusmack 
      Caption         =   "M&ain"
      Visible         =   0   'False
      Begin VB.Menu vismnuHide 
         Caption         =   "H&ide Perfect Macro Recorder"
      End
      Begin VB.Menu vismnudash1 
         Caption         =   "-"
      End
      Begin VB.Menu vismnuRecord 
         Caption         =   "Record New Macro"
      End
      Begin VB.Menu vismnuplayback 
         Caption         =   "Pla&yBack Macro"
         Enabled         =   0   'False
         Begin VB.Menu vismnuAvailbemacros 
            Caption         =   "Available Macros"
            Index           =   0
         End
      End
      Begin VB.Menu vismnudash2 
         Caption         =   "-"
      End
      Begin VB.Menu vismnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu vismnuControl 
      Caption         =   "&Control"
      Visible         =   0   'False
      Begin VB.Menu vismnumanage 
         Caption         =   "&Manage Macros"
      End
      Begin VB.Menu vismnuLanguage 
         Caption         =   "Language"
         Visible         =   0   'False
      End
      Begin VB.Menu visImport 
         Caption         =   "Import macro"
      End
      Begin VB.Menu visExport 
         Caption         =   "Backup macro"
      End
      Begin VB.Menu visEXE 
         Caption         =   "Export to EXE file"
      End
      Begin VB.Menu vismnuoptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu vismnuhelp 
      Caption         =   "He&lp"
      Visible         =   0   'False
      Begin VB.Menu vismnuhelpt 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu vismnuCprinGoldonline 
         Caption         =   "Perfection Tools online"
      End
      Begin VB.Menu vismnuSendReviews 
         Caption         =   "Send Reviews"
         Visible         =   0   'False
      End
      Begin VB.Menu vismnuabout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmJMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TickCountNow As Variant

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Sub cmdAliasRecord_Click()
        Rem because we can't rename cmdRecord
        Rem we make Alias Button
        cmdRecord_Click
End Sub

Private Sub cmdBackUpAll_Click()
    
'//pt 20Aug2007:
Dim PaidFor As Boolean
'//pt 20Aug2007:
SECUREBEGIN_B
'//pt 20Aug2007:
PaidFor = True

    Dim ShowFolderResult As Boolean
    Dim SelectedFolderPath As String

    '//cprin 26Jul2006: added next line to display frmBackUpAll modelly
    frmBackUpAll.Show 1
     
    '//cprin 26Jul2006: Only if User Click Ok button in frmBackUpAll
    If Tags.BuckupAllType <> "" Then
       '//cprin 26Jul2006: User Choosed to Bundle
       If Tags.BuckupAllType = "bundle" Then
           '//cprin 27Jul2006: Bandle all macros files in one large file
           BandleAllMacros
       '//cprin 26Jul2006: User Choosed to Separate
       ElseIf Tags.BuckupAllType = "separate" Then
           '//cprin 26Jul2006: Display ShowFolder Dialog to get the folder from the user in which to save files to
           ShowFolderResult = ShowSelectFolder(hWnd, "Select folder to save BackUp files to.", SelectedFolderPath)
           '//cprin 26Jul2006: User Press Ok button in ShowFolder Dialog
           If ShowFolderResult Then
              '//cprin 26Jul2006: Backup All macros of PMR into a selected folder by the User
              BackupAllMacros (SelectedFolderPath)
           End If
       End If
    End If
    
'//pt 20Aug2007:
SECUREEND_B

'//pt 20Aug2007:
If PaidFor = False Then
   
   MsgBox "Sorry, this function is only available to registered users."

End If
    
End Sub

Private Sub cmdBackUpAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
    
    '//cprin 19July2006: added next line to dispaly a balloon tooltip whenever mousemove called
    BackAllBalloon.ShowBalloon cmdBackUpAll.hWnd, "Click to backup all the macros you had created.", SoftwareName, Info, 2000, 5000
End Sub

Private Sub cmdContinue_Click()

Dim MyMacroName As String
Dim FrmJMacroState As Boolean

Rem the Story for the only next line of code is very Important
Rem while i test the program , i found that if click the playback button from
Rem the FrmJmacro From,the playback repeat itself due to sending a click from
Rem Space button .......etc
Rem i solve this problem by minimize FrmJmacro when play back get started
Rem and we make this in line 20 and line 30
'========================================================
Rem The Brief:
Rem we minimize the FrmJmacro if she was displayed in Line 20
Rem Then we restore it to it's orginial state in Line 30
'========================================================
20  If Me.WindowState <> 1 Then Me.WindowState = 1: FrmJMacroState = True

    If Not IsPlayingMode Then
    'if no playbacking running
        'If GetAsyncKeyState(18) And GetAsyncKeyState(vbKeyInsert) Then
        'element 0 = macropath
        'element 1 = playlog
        'element 2 = there is a paused macro or not
        'PlayBackPauseForValerie_Dion(0) = ""
        'PlayBackPauseForValerie_Dion(1) = 0
            'element 3 = there is a paused macro or not
            If PlayBackPauseForValerie_Dion(2) = 1 Then
                PlayBackPauseForValerie_Dion(2) = 0
                'Debug.Print "alt + pause"
                MyMacroName = PlayBackPauseForValerie_Dion(0)
                '//cprin 28July2006: added next line to make an animation in systray whenever you start Play back
                If TrayIcon1.AnimateState = False Then TrayIcon1.Animate ImageList1, 100
                Call OpenMacro(MyMacroName)
                StartPlay True
                WatchForCancel
                ''TrayIcon1.ShowBalloon "PlayBack Over", App.Title, Info, 3000
            End If
        'End If
    End If
    
    '//cprin 28July2006: added next line Stop the animation
    If TrayIcon1.AnimateState = True Then TrayIcon1.StopAnimateing (Me.Icon)
    
             '/////////////////////////
         Dim key12 As String
         Dim key13 As String
         Dim dkey1 As String * 15
         Dim DefChar7 As String
         Dim DefChar9 As String
         DefChar9 = Str(vbKeyPause)
         DefChar7 = Str(vbKeyF11)
         GetINIProfileString Resume_Play_back, "0,1,0#" & DefChar9 & "*", dkey1, Len(dkey1)
         key12 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         
         GetINIProfileString Stop_Play_back, "0,1,0#" & DefChar7 & "*", dkey1, Len(dkey1)
         key13 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         '////////////////////////////////////////////
         
    If PlayBackPauseForValerie_Dion(2) = 1 Then
       TrayIcon1.ShowBalloon "Press " & DecodeKey(key12) & " to resume.", "Replaying " & Chr(34) & CobMacros.List(CobMacros.ListIndex) & Chr(34) & " Macro Paused", Info, 3000
    ElseIf PlayBackPauseForValerie_Dion(2) = 0 Then
        TrayIcon1.ShowBalloon "Replaying " & CobMacros.List(CobMacros.ListIndex) & " Macro Finish", App.Title, Info, 3000
    End If
        

30  If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE



End Sub

Private Sub cmdContinue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
    
    '//cprin 19July2006: added next line to dispaly a balloon tooltip whenever mousemove called
    ContinueBallon.ShowBalloon cmdContinue.hWnd, "Click to resume play backing the macro" & Chr(13) & Chr(10), SoftwareName, Info, 2000, 5000

End Sub

Private Sub cmdEditMacro_Click()
   frmEdit.Show
End Sub

Private Sub cmdEditMacro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
    
    '//cprin 19July2006: added next line to dispaly a balloon tooltip whenever mousemove called
    EditBalloon.ShowBalloon cmdEditMacro.hWnd, "Click to edit the macro selected in the Combo Box.", SoftwareName, Info, 2000, 5000

End Sub

Private Sub cmdExport_Click()
    'cprin 13July2006: added next line in support of adding cmdExport Button
    visExport_Click
End Sub

Private Sub cmdExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
    
    '//cprin 19July2006: added next line to dispaly a balloon tooltip whenever mousemove called
    BackupMacro.ShowBalloon cmdExport.hWnd, "Click to backup selected macro in the combo box.", SoftwareName, Info, 2000, 5000
End Sub

Private Sub cmdHelp_Click()

  PopupMenu mnuHelp, 0, cmdHelp.Left, cmdHelp.Top + cmdHelp.Height, mnuHelp2

End Sub

Private Sub cmdImport_Click()
     'cprin 13July2006: added next line in support of adding cmdImport Button
      visImport_Click
End Sub

Private Sub cmdImport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
     
    
    '//cprin 19July2006: added next line to dispaly a balloon tooltip whenever mousemove called
    ImportBalloon.ShowBalloon cmdImport.hWnd, _
    "Click to import macros to " & SoftwareName & " that" & Chr(13) & Chr(10) & "you had backup before using Backup Macro or Backup All buttons.", SoftwareName, Info, 2000, 5000
End Sub

Private Sub cmdLanguage_Click()

'frmLanguage.Show 1

End Sub

Private Sub cmdmanage_Click()
    
  frmManage.Show , frmJMacro
  
End Sub

Private Sub cmdmanage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
    
    '//cprin 19July2006: added next line to dispaly a balloon tooltip whenever mousemove called
    MyMacrosBalloon.ShowBalloon cmdmanage.hWnd, "Here you can found and manage" & Chr(13) & Chr(10) & "all the macros you has created.", SoftwareName, Info, 2000, 5000
End Sub

Private Sub cmdoptions_Click()
     
     Load frmOptions2
     
     frmOptions2.Show , frmJMacro
  
End Sub

Private Sub cmdoptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
    
    '//cprin 19July2006: added next line to dispaly a balloon tooltip whenever mousemove called
    OptionBalllon.ShowBalloon cmdoptions.hWnd, "Click to Change" & Chr(13) & Chr(10) & SoftwareName & " Options.", SoftwareName, Info, 2000, 5000
End Sub

Private Sub cmdPlayBack_Click()
    '//cprin 11Aug2006: changed next line to fix a bug reported by marcos
    '//where we get overflow error if repeat count is so big
    'Dim i As Long
    Dim i As Variant
    Dim OF As OFSTRUCT
    Dim TipState As String
    '//cprin 11Aug2006: changed next line to fix a bug reported by marcos
    '//where we get overflow error if repeat count is so big
    'Dim RepeatCount As Long
    Dim RepeatCount As Variant
    
    Dim FrmJMacroState As Boolean
    '//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
    Dim IsAlreadyChoosed As Integer
    Dim SelectedMacroName As String
    
    '//cprin 28July2006: disable the button again since there are no paused macro
    cmdContinue.Enabled = False
    
    TipState = GetSetting(App.EXEName, "k", "d", "")
    
             '/////////////////////////
         Dim key12 As String
         Dim key13 As String
         Dim key14 As String
         Dim dkey1 As String * 15
         Dim DefChar7 As String
         Dim DefChar9 As String
         Dim DefChar4 As String
         
         DefChar4 = Str(vbKeyPause)
         DefChar9 = Str(vbKeyPause)
         DefChar7 = Str(vbKeyF11)
         GetINIProfileString Resume_Play_back, "0,1,0#" & DefChar9 & "*", dkey1, Len(dkey1)
         key12 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         
         GetINIProfileString Stop_Play_back, "0,1,0#" & DefChar7 & "*", dkey1, Len(dkey1)
         key13 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         
         GetINIProfileString Pause_Play_back, "0,0,0#" & DefChar4 & "*", dkey1, Len(dkey1)
         key14 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         '////////////////////////////////////////////
         
         
    If Not TipState = "one" Then
          
           frmTip.lblTip.Caption = "Press " & DecodeKey(key13) & " to stop Play back." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Press " & DecodeKey(key14) & " to Pause Play back." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Press " & DecodeKey(key12) & " to Resume Play back."
             
           frmTip.Show 1
    End If
    
    Dim res22 As Integer
    Dim IsForEver As Boolean
    res22 = GetINIProfileInt("repeat", 1)
     
    Select Case res22
           
           Case 1
                
                tmrTerminatePlayback.Enabled = False '//just to make sure
                txtRepeat = GetINIProfileInt("repeatcount", 1)
           
           Case 2
                            
                Debug.Print 1, time
                TickCountNow = GetTickCount
                
                tmrTerminatePlayback.Enabled = True
                 
                IsForEver = True
                 
           Case 3
                 
                tmrTerminatePlayback.Enabled = False '//just to make sure
                IsForEver = True
    
    End Select
    
    If txtRepeat <> "" Then RepeatCount = txtRepeat Else RepeatCount = 1

Rem the Story for the only next line of code is very Important
Rem while i test the program , i found that if click the playback button from
Rem the FrmJmacro From,the playback repeat itself due to sending a click from
Rem Space button .......etc
Rem i solve this problem by minimize FrmJmacro when play back get started
Rem and we make this in line 20 and line 30
'========================================================
Rem The Brief:
Rem we minimize the FrmJmacro if she was displayed in Line 20
Rem Then we restore it to it's orginial state in Line 30
'========================================================
20  If Me.WindowState <> 1 Then Me.WindowState = 1: FrmJMacroState = True
    
    '//cprin 11Aug2006: changed next line to fix a bug reported by marcos
    '//where we get overflow error if repeat count is so big
    'For i = 1 To RepeatCount
    Do While CDbl(i) < CDbl(RepeatCount) '//the loop will be infinite if you didn't convert both RepeatCount and i from varient to Double
        
          SelectedMacroName = AppPath & "Data\" & CobMacros.List(CobMacros.ListIndex) & ".pmf"
          
          '//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
          'IsAlreadyChoosed = OpenFile(SelectedMacroName, OF, OF_READ)
          
          '//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
          'CloseHandle IsAlreadyChoosed
            
          '//cprin 29Jul2006: changed next line because if file name is very large, OpenFile return -1
          'If IsAlreadyChoosed <> -1 Then
          '//cprin 9Feb2007: changed next line because PathFileExists fails due to Craig=>operationivyflag@gmail.com report
          'If PathFileExists(SelectedMacroName) Then
          If FileExists(SelectedMacroName) Then
         
            '//cprin 16July2006: added next line to make an animation in systray whenever you start Play back
            If TrayIcon1.AnimateState = False Then TrayIcon1.Animate ImageList1, 100
            
            Call OpenMacro(SelectedMacroName, CobMacros.List(CobMacros.ListIndex))
              
            StartPlay
         
            WatchForCancel
            
         Else

             '/cprin 17July2006: changed next MsgBox function to display a well form messagebox.
             MsgBox "Can't Found the data file of " & CobMacros.List(CobMacros.ListIndex) & " macro" & Chr(13) & Chr(10) _
             & "(" & SelectedMacroName & ")", vbCritical, App.Title
             If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE
             Exit Sub
         End If
     
    '//cprin 27Jul2006: added next line to fix the bug of "can't stop play back when press alt+f10 if repeatcount is more than 1"
    If Tags.StopPlayBackLoop Then
       '//reset the StopPlayBackLoop again
       Tags.StopPlayBackLoop = False
       '//cprin 11Aug2006: changed next line to fix a bug reported by marcos
       '//where we get overflow error if repeat count is so big
       Exit Do
    End If
    
    If IsForEver = False Then
        '//cprin 11Aug2006: changed next line to fix a bug reported by marcos
        '//where we get overflow error if repeat count is so big
        'Next i
        i = i + 1
    End If
    
    Loop
    
    '//cprin 16July2006: added next line Stop the animation
    If TrayIcon1.AnimateState = True Then TrayIcon1.StopAnimateing (Me.Icon)
           
        If PlayBackPauseForValerie_Dion(2) = 1 Then
          TrayIcon1.ShowBalloon "Press " & DecodeKey(key12) & " to resume.", "Replaying " & Chr(34) & CobMacros.List(CobMacros.ListIndex) & Chr(34) & " Macro Paused", Info, 3000
        ElseIf PlayBackPauseForValerie_Dion(2) = 0 Then
          TrayIcon1.ShowBalloon "Replaying " & CobMacros.List(CobMacros.ListIndex) & " Macro Finish", App.Title, Info, 3000
        End If
      '  MsgBox "PlayBack " & CobMacros.List(CobMacros.ListIndex) & " Macro Over", vbInformation, SoftwareName
        
30      If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE
        
End Sub

Private Sub tmrTerminatePlayback_Timer()
    
    Dim ttick As Long
    Dim TotalMillySecond As Long
    Dim WhereHash As Integer
    Dim dkey2 As String * 40
     
    GetINIProfileString "tmsv", "0", dkey2, Len(dkey2)
    
    WhereHash = InStr(1, dkey2, "#")
    
    ttick = GetTickCount - TickCountNow
    
    TotalMillySecond = Left(dkey2, WhereHash - 1)

    If ttick >= TotalMillySecond Then
       
       Debug.Print 2, time
       
       CheckStopKey True
       
       tmrTerminatePlayback.Enabled = False
   
    End If
    
End Sub


Private Sub cmdPlayBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
    
    '//cprin 19July2006: added next line to dispaly a balloon tooltip whenever mousemove called
    PlayBackBalloon.ShowBalloon cmdPlayBack.hWnd, "Click to play back the selected macro in the Combo Box.", SoftwareName, Info, 2000, 5000
End Sub

Public Sub cmdRecord_Click()

Dim Pos As Long
Dim Result As Long
Dim res As String * 99
Dim Res2 As String * 30
Dim TipState As String
Dim IsFrmManageVisiable
Dim ForeGroundW As Long
Dim NewMacroName As String
Static LastState As Boolean
Static ForeGroundText As String
Static ForeGroundClass As String

If IsPlayingMode Then

Rem It is not realy a click for cmdRecord Button
Rem it is the last message Recored in the last Macro
Rem when he just attempt to stop Recording his Macro
Rem so we will Exit this sub

    Exit Sub

End If

If cmdRecord.Caption = "&Record" Then
       
   TipState = GetSetting(App.EXEName, "k", "d", "")
    
         Dim key12 As String
         Dim key13 As String
         Dim key14 As String
         Dim dkey1 As String * 15
         Dim DefChar11 As String
         Dim DefChar12 As String
         Dim DefChar8 As String
         DefChar12 = Str(vbKeyF10)
         DefChar11 = Str(vbKeyInsert)
         DefChar8 = Str(vbKeyInsert)
         GetINIProfileString Stop_Record, "0,1,0#" & DefChar12 & "*", dkey1, Len(dkey1)
         key12 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         
         GetINIProfileString Resume_Record, "0,1,0#" & DefChar11 & "*", dkey1, Len(dkey1)
         key13 = Mid(dkey1, 1, InStr(1, dkey1, "*"))

         GetINIProfileString Pause_Recording, "0,0,0#" & DefChar8 & "*", dkey1, Len(dkey1)
         key14 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         
   If Not TipState = "one" Then
      
       frmTip.lblTip.Caption = "Press " & DecodeKey(key12) & " to stop Recording" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Press " & DecodeKey(key14) & " to Pause Recording" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Press " & DecodeKey(key13) & " to Resume Recording"
      
       frmTip.Show 1, frmJMacro
       
     '  SetWindowPos frmTip.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1

   End If
               
   '//cprin 5June2007:
   StartRecord
   '//cprin 5June2007:
   If Not IsRecordingMode Then
      Exit Sub
   End If
   
   LastState = frmJMacro.Visible
   
   Me.Visible = False
    
   If frmJMacro.WindowState <> 1 Then
    
       frmJMacro.WindowState = 1
 
   End If
            
   TrayIcon1.Animate ImageList1, 100
    
   TrayIcon1.IconToolTip = "Recording..."
   
   mnuRecord.Caption = "Stop Recording"
   
   cmdRecord.Caption = "&Stop Recording"
         
   ForeGroundW = GetForegroundWindow

Rem GetForegroundWindowName
   
   Result = GetWindowText(ForeGroundW, res, 100)
    
   Pos = InStr(1, res, Chr(0))
   
   ForeGroundText = Left(res, Pos - 1)

Rem GetClassName

   Result = GetClassName(ForeGroundW, Res2, 100)
   
   Pos = InStr(1, Res2, Chr(0))
    
   ForeGroundClass = Left(Res2, Pos - 1)
   

'we must make Tags.RecordingIsOver = Flase
'because if we don't make this the record willn't run at all
'you say why i will tell you
'In JournalRecordProc you see the following statment
'If Tags.RecordingIsOver then exit sub
'do you see
 
   Tags.RecordingIsOver = False
   
   '//cprin 5June2007: comment next line
   '//StartRecord
    
ElseIf cmdRecord.Caption = "&Stop Recording" Then
        
If Not RecordPauseForValerie_Dion(2) = 1 Then
   'only stopani if there is no paused recording session
   'waiting for resume.
   Call TrayIcon1.StopAnimateing(Me.Icon)
Else
  RecordPauseForValerie_Dion(2) = 0
  TrayIcon1.ChangeIcon Me.Icon
'MsgBox "sdsdfg"
'if there is a paused macro what shall i do

End If
TrayIcon1.IconToolTip = App.Title
'Very Very Very Very Very Very Important
'Why we use Tags.RecordingIsOver?
'I will Tell you...

'when the recording Finish...
'FrmEnterMacroName Form appear
'the problem is ===>
'the recording is still running and every key you hit for saving the macro
'is recorded in the macro

'ok what i really want is?

'Stoping Macro at the moment when i hit alt-F10 for ex,
'in other word i didn't save the unneeded messages of The FrmEnterMacroName because it's not
'the macro.

'ok but i didn't understand anything of all this bullshit? can you explain more.

'ok i will tell you the problem is steps..

'1.User start a macro
'2.User hit alt-F10 to Finish the macro. "at this moment the really macro must by stop his recording"
'3.FrmEnterMacroName appear.
'4.while you write the name of the macro, the recording is still running
'5.every key you hit for the name of the macro is recorded
'6.do you imagine what is the funken of this?
'this is the problem

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Tags.RecordingIsOver = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
'Another kind of problems appear
'we must make Tags.RecordingIsOver = Flase
'because if we don't make this the record willn't run at all
'you say why i will tell you
'In JournalRecordProc you see the following statment
'If Tags.RecordingIsOver then exit sub
'do you see
     
   
   
   If LastState Then
       
       Me.Visible = True
       
       mnu_jmacro.Caption = "Hide " & SoftwareName
       
       Me.WindowState = 0
   
   End If
   
   cmdRecord.Caption = "&Record"
   
   mnuRecord.Caption = "&Record"
   
   frmEnterMacroName.Show 1, frmJMacro
    
   TrayIcon1.ChangeIcon Me.Icon
   
   If frmEnterMacroName.Tag = "ok" Then
   
               NewMacroName = frmEnterMacroName.txtMacroName
          
               Unload frmEnterMacroName
     
               On Error Resume Next
             
                      MkDir AppPath & "Data"
          
               On Error GoTo 0
          
              Call StopMacro(NewMacroName, False)
              
              SaveSetting App.EXEName, "for", NewMacroName, ForeGroundText
               
              SaveSetting App.EXEName, "forClass", NewMacroName, ForeGroundClass
              
              ForeGroundText = ""
              
              ForeGroundClass = ""
              
Rem           SaveSetting App.EXEName, "sm", NewMacroName, AppPath & "Data\" & NewMacroName & ".pmf"
                                     
Rem           SaveSetting App.EXEName, "sm", NewMacroName, ""
                                    
              FillComboBox

Rem           if the User Open Manage window while Recording his Macro
Rem           and Click "&Stop Macro"we must update frmManage with new
              
              IsFrmManageVisiable = FindWindow("ThunderRT6FormDC", "Manage your Macros")
              
              If IsFrmManageVisiable <> 0 Then
                 
                 Unload frmManage
              
                 frmManage.Show , frmJMacro
              
              End If
              
   ElseIf frmEnterMacroName.Tag = "cancel" Then
            
              Unload frmEnterMacroName
       
              Call StopMacro(NewMacroName, True)
          
   End If
   
   
End If

'Another kind of problems appear
'we must make Tags.RecordingIsOver = Flase
'because if we don't make this the record willn't run at all
'you say why i will tell you
'In JournalRecordProc you see the following statment
'If Tags.RecordingIsOver then exit sub
'do you see
 
 Tags.RecordingIsOver = False
End Sub

Private Sub cmdRecord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
    '//cprin 19July2006: added next line to dispaly a balloon tooltip whenever mousemove called
    RecordBalloon.ShowBalloon cmdRecord.hWnd, "Click to Start Recording a new macro now.", SoftwareName, Info, 2000, 5000
End Sub

Private Sub cmdreviews_Click()
    mnuExit_Click
End Sub

Private Sub Command1_Click()
'//pt 20Aug2007:
Dim PaidFor As Boolean
'//pt 20Aug2007:
SECUREBEGIN
'//pt 20Aug2007:
PaidFor = True

Dim marker As Variant
Dim TempOne As String
Dim MacroData As String
Dim StandAloneData As String
Dim StandAloneMacro As String
Dim cprsosonono3345554 As Long
Dim SelectedMacroName As String
Dim OldCap As String

On Error GoTo quitme:

If CobMacros.ListIndex = -1 Then
   MsgBox "There is no macro to compile."
   Exit Sub
End If

With CommonDialog1
     .FileName = CobMacros.List(CobMacros.ListIndex)
     .Filter = "Executable File (*.exe)|*.exe"
     .CancelError = True
     .flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist '+ cdlOFNExtensionDifferent
     .ShowSave
End With

TempOne = GetTempPathName & CommonDialog1.FileTitle

StandAloneMacro = TempOne

OldCap = Me.Caption

Me.Caption = "Wait..."

If right(StandAloneMacro, 4) <> ".exe" Then
   StandAloneMacro = StandAloneMacro & ".exe"
End If

TryAgain:

SelectedMacroName = AppPath & "Data\" & CobMacros.List(CobMacros.ListIndex) & ".pmf"


Open StandAloneMacro For Output As #1
     
     Print #1, StrConv(LoadResData(101, "CUSTOM"), vbUnicode)
     
Close #1 ' there is a 2 close downthere


'check if the resource file is still vaild or invaild by a funken hacker.

'Sleep 2000 'sleep until creating the file

'Shell StandAloneMacro & " /check"

'cprsosonono3345554 = FindWindow(vbNullString, "cprsosonono3345554")

'If cprsosonono3345554 <> 0 Then
'Else
'  Close #1
'  Kill StandAloneMacro
'  MsgBox "Critical Error,please reinstall " & SoftwareName & ".", vbCritical, SoftwareName
'  Exit Sub
'End If
'ok the file in vaild


Open StandAloneMacro For Binary As #2
        
     StandAloneData = Input(LOF(2), 2)
        
Close #2


Open SelectedMacroName For Binary As #3
        
    MacroData = Input(LOF(3), 3)
        
Close #3
    
marker = InStr(1, StandAloneData, "cprinsosonono")

StandAloneData = Left(StandAloneData, marker - 1)

StandAloneData = StandAloneData & "cprinsosonono" & MacroData


Open StandAloneMacro For Binary As #4
     
     Put #4, 1, StandAloneData
     
Close #4

FileCopy TempOne, CommonDialog1.FileName

'//cprin 17July2006: changed next line to display a well form MsgBox
MsgBox "Executable macro" & " (" & CommonDialog1.FileTitle & ") " & "saved to " & CommonDialog1.FileName

Me.Caption = OldCap

On Error Resume Next

Kill TempOne

Exit Sub
quitme:

''MsgBox Err.Description, vbCritical, App.Title


'If Err.Number = cdlDialogFailure Then
'   StandAloneMacro = AppPath & CobMacros.List(CobMacros.ListIndex) & ".exe"
'   MsgBox "Save as Dialog Failure, Standalone file saved automatically to " & StandAloneMacro
'   GoTo TryAgain:
'ElseIf Err.Number = 326 Then
'   Close #1
On Error Resume Next
Dim Pathto As Boolean
'//cprin 9Feb2007: changed next line because PathFileExists fails due to Craig=>operationivyflag@gmail.com report
'//Pathto = PathFileExists(TempOne)
Pathto = FileExists(TempOne)
If Pathto Then
   Kill StandAloneMacro
End If
'   MsgBox "Error please reinstall " & SoftwareName & ".", vbCritical, SoftwareName
'End If

'//pt 20Aug2007:
SECUREEND

'//pt 20Aug2007:
If PaidFor = False Then
   
MsgBox "In the Trial period, the executable macro file you created will be run for only five seconds, " & Chr(13) & Chr(10) & "hence you need to purchase " & SoftwareName & " to create an executable macro file" _
    & Chr(13) & Chr(10) & "that can be run without this limitation."
    
   EvulateExeFiles

End If
    
End Sub


Private Sub cmdStandaLone_Click()

Command1_Click
 
End Sub


Private Sub CprinGold_Click()

ShellExecute Me.hWnd, vbNullString, SiteName, vbNullString, "", SW_SHOWNORMAL

End Sub

Private Sub FlatButton1_Click()

cmdHelp_Click

End Sub

Private Sub cmdStandaLone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
    
    '//cprin 19July2006: added next line to dispaly a balloon tooltip whenever mousemove called
    ExeBalloon.ShowBalloon cmdStandaLone.hWnd, "Click to convert selected macro in the combo box to an EXE-file" & Chr(13) & Chr(10) & "that can runs stand alone without " & SoftwareName & ".", SoftwareName, Info, 2000, 5000
End Sub

Private Sub Form_Load()

On Error Resume Next
'//pt 23Aug2007: Hide "Enter Regsiterision Key..." if the key is enterd
Dim KeyCode As String
KeyCode = ArmEnviron$("USERKEY")
If KeyCode <> "" Then mnuRegistrationKey.Visible = False
On Error GoTo 0
 
'//cprin 26Jul2006: added next 4 lines to Test ShowSelectFolder function
'Dim res As Boolean
'Dim path As String
'res = ShowSelectFolder(hWnd, "Select folder to save BackUp files to.", path)
'If res Then MsgBox path

Dim LastRunCount As Long
Dim ShowAtStartup As Long
Dim LastRunMacro As String
Dim StartUpMacro As String
Dim SmallMacroName As String
Dim FrmJMacroState As Boolean

'cprin 13July2006: added next code block
'to ensure that jeweled buttons will remain
'as i want.
''''''''''''''''''''''''''''''''''''''''''''
cmdHelp.Caption = "&Help"
cmdreviews.Caption = "&Exit"
cmdmanage.Caption = "&My Macros"
cmdImport.Caption = "&Import"
cmdGetNow.Caption = "&GetNow!"
cmdoptions.Caption = "&Options"
cmdPlayBack.Caption = "&Play back"
cmdExport.Caption = "&Backup Macro"
cmdBackUpAll.Caption = "Backup &All"
cmdRecord.Caption = "&Record"
cmdEditMacro.Caption = "&Edit Macro"
cmdStandaLone.Caption = "E&xport to EXE"
''''''''''''''''''''''''''''''''''''''''''''

Dim Res2 As Integer
   
Res2 = GetINIProfileInt("DASK", 0)
   
If Res2 Then
   
   mnuDisableAllShortcutKeys.Caption = "Enable All Shortcut Keys"

Else
   
   mnuDisableAllShortcutKeys.Caption = "Disable All Shortcut Keys"

End If
   
If Command$ = "hide" Then
   
  mnu_jmacro_Click

Else
    
    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    If ShowAtStartup <> 0 Then
   '    frmTip1.Show
    End If

   
End If

LastRunCount = Val(GetSetting(App.EXEName, "f", "a", 0))
   
If LastRunCount = 0 Then
 
      SaveSetting App.EXEName, "f", "a", 1
      SaveSetting App.EXEName, "GHK", 2, "Alt+F9"
      SaveSetting App.EXEName, "GHK", 3, "Alt+F10"
      SaveSetting App.EXEName, "GHK", 4, "Alt+F10"
      SaveSetting App.EXEName, "GHK", 5, "Ctrl+Alt+s"
      SaveSetting App.EXEName, "GHK", 6, "Ctrl+Alt+o"
      SaveSetting App.EXEName, "GHK", 7, "Ctrl+Alt+m"
              

Else
       
      SaveSetting App.EXEName, "f", "a", 1
      
End If

TrayIcon1.Show Me.Icon, App.Title

TrayIcon1.TrackPopMenu = True

FillComboBox

'//pt 8Feb2007: added next if block to fix the problem of "Krayg Draypur"=>operationivyflag@gmail.com
'  and the problem of "Calvin Chang"=>mystery5096@gmail.com where the cmdPlayback button is still enabled
'  however no macros is in the combo box
If Not CobMacros.ListCount > 0 Then

       cmdEditMacro.Enabled = False
       
       cmdPlayBack.Enabled = False
       
       mnuPlayback.Enabled = False
        
       vismnuplayback.Enabled = False

End If

SetPlayBackSpeed

'//cprin 19July2006: added next line to subclass the combo box to
'1.display a balloon tooltip at mouse move
'2.display some info about macros "may be i won't do this feature"
SubClassCombo.BeginSubClass CobMacros.hWnd

LastRunMacro = GetSetting(App.EXEName, "las", "las", "nono")

On Error Resume Next
If LastRunMacro <> "nono" Then
   If CobMacros.ListIndex <> -1 Then
      CobMacros.ListIndex = Val(LastRunMacro)
   End If
End If
On Error GoTo 0
'StartUpMacro = GetSetting(App.EXEName, "fr", "st", "")

'SmallMacroName = StartUpMacro

'If StartUpMacro <> "" Then
   
'    StartUpMacro = AppPath & "Data\" & StartUpMacro & ".pmf"
    
Rem the Story for the only next line of code is very Important
Rem while i test the program , i found that if click the playback button from
Rem the FrmJmacro From,the playback repeat itself due to sending a click from
Rem Space button .......etc
Rem i solve this problem by minimize FrmJmacro when play back get started
Rem and we make this in line 20 and line 30
'========================================================
Rem The Brief:
Rem we minimize the FrmJmacro if she was displayed in Line 20
Rem Then we restore it to it's orginial state in Line 30
'========================================================
    
'    If Me.WindowState <> 1 Then Me.WindowState = 1: FrmJMacroState = True
    
'    Call OpenMacro(StartUpMacro, SmallMacroName)

'    StartPlay

'    WatchForCancel
     
'    If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE
         
'End If

     
End Sub

Private Sub Form_Resize()


   If Me.WindowState = 1 Then
 
   
   If mnu_jmacro.Caption = "Show " & SoftwareName Then
            
             mnu_jmacro.Caption = "Hide " & SoftwareName
        
   ElseIf mnu_jmacro.Caption = "Hide " & SoftwareName Then
          
            If Not cmdRecord.Caption = "&Stop Recording" Then
             
             Me.Visible = False
            
            End If
            
            mnu_jmacro.Caption = "Show " & SoftwareName
         
   End If

   End If
    
   
End Sub

Private Sub menuHelp_Click()
   '2
    Shell "hh " & AppPath & "help.chm", vbMaximizedFocus

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to Destroy the balloon tooltip whenever mousemove called
    RecordBalloon.Destroy
    EditBalloon.Destroy
    PlayBackBalloon.Destroy
    ContinueBallon.Destroy
    ComboBalloon.Destroy
    MyMacrosBalloon.Destroy
    OptionBalllon.Destroy
    ImportBalloon.Destroy
    BackupMacro.Destroy
    BackAllBalloon.Destroy
    ExeBalloon.Destroy
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '//cprin 19July2006: added next line to Destroy the balloon tooltip whenever mousemove called
    MyMacrosBalloon.Destroy
    OptionBalllon.Destroy
    ImportBalloon.Destroy
    BackupMacro.Destroy
    BackAllBalloon.Destroy
    ExeBalloon.Destroy
End Sub

Private Sub mnuabout_Click()
   frmAbout.Show , frmJMacro
End Sub

Private Sub mnuDisableAllShortcutKeys_Click()


    If mnuDisableAllShortcutKeys.Caption = "Disable All Shortcut Keys" Then
       WriteToINIProfile "DASK", 1
       mnuDisableAllShortcutKeys.Caption = "Enable All Shortcut Keys"
    ElseIf mnuDisableAllShortcutKeys.Caption = "Enable All Shortcut Keys" Then
       WriteToINIProfile "DASK", 0
        mnuDisableAllShortcutKeys.Caption = "Disable All Shortcut Keys"
    End If

End Sub

Private Sub mnuHelp2_Click()
'1
   Shell "hh " & AppPath & "help.chm", vbMaximizedFocus

End Sub

Private Sub mnuSendReviews_Click()

ShellExecute Me.hWnd, vbNullString, "mailto:" & Email, vbNullString, "", SW_SHOWNORMAL

End Sub

'//pt 23Aug2007:
Private Sub mnuRegistrationKey_Click()
    
    ShowEnterKeyDialog Me.hWnd
    
    '//pt 23Aug2007: Hide "Enter Regsiterision Key..." if the key is enterd
    Dim KeyCode As String
    KeyCode = ArmEnviron$("USERKEY")
    If KeyCode <> "" Then mnuRegistrationKey.Visible = False

End Sub

Private Sub mnuweb_Click()

ShellExecute Me.hWnd, vbNullString, SiteName, vbNullString, "", SW_SHOWNORMAL

End Sub

Private Sub SubClassCombo_MessageReceived(hWnd As Long, Msg As Long, wParam As Long, lParam As Long, cancel As Boolean)
    '//cprin 19July2006: added next line to disable balloon tootip if it isn't the first time for running this software
    If Not IsBalloonTooltipsEnabled Then Exit Sub
    
    '//cprin 19July2006: added next if block to dispaly a balloon tooltip whenever mousemove called
    If Msg = WM_MOUSEMOVE Then
        ComboBalloon.ShowBalloon hWnd, "Select the macro you want from this combo box in order to" & Chr(13) & Chr(10) & "playback, edit, backup or export to an executable file.", SoftwareName, Info, 2000, 5000
    End If
End Sub

Private Sub tmrCheckShortKeys_Timer()
    
   If IsPlayingMode = False Then
   
      tmrTerminatePlayback.Enabled = False
      
      '//if the icons animated while no recording or playback
      If TrayIcon1.AnimateState = True And Not IsRecordingMode Then
          TrayIcon1.StopAnimateing (Me.Icon)
      End If
      
   End If
     
   '//Ctrl + Alt + Del
   If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyDelete) And GetAsyncKeyState(18) And IsPlayingMode Then
            
      CheckStopKey True
           
      bPlaying = False
      
   ElseIf GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyDelete) And GetAsyncKeyState(18) And IsRecordingMode Then
        
      frmJMacro.cmdRecord_Click
          
   End If
  
  
   '//The New Check ShortCutKeys
   CheckOptionsShortcutKeys

   '//if there is a pause macro then
   If PlayBackPauseForValerie_Dion(2) = 1 Then
      '//cprin 28July2006: enable the button if there is a paused macro waiting for continue
      cmdContinue.Enabled = True
   Else
      cmdContinue.Enabled = False
   End If
   
   Dim Fore_Hwnd As Long
   CheckShortKeys
   
   If CobMacros.ListIndex <> -1 Then
      cmdStandaLone.Enabled = True
      '//cprin 17July2006: added next 2 lines to enable
      '//cmdExport and cmdBackUpAll when at least one macro is in the combo box
      cmdExport.Enabled = True
      cmdBackUpAll.Enabled = True
   Else
      cmdStandaLone.Enabled = False
      '//cprin 17July2006: added next 2 lines to disable
      '//cmdExport and cmdBackUpAll when at least one macro is in the combo box
      cmdExport.Enabled = False
      cmdBackUpAll.Enabled = False
      Exit Sub
   End If
     
'  If IsRecordingMode Then
    'Start Smart Features
    'here we get the window Information
    'debug.Print EventArr(EventLog).hwnd
'    Fore_Hwnd = GetForegroundWindow
'    GetWindowInfo (Fore_Hwnd)
'  End If
  
End Sub

Private Sub Form_Unload(cancel As Integer)

'//cprin 19July2006: added next line to Unsubclass the combo box.
SubClassCombo.EndSubClass

Dim isStillRecording As Boolean
Dim isStillPlaybacking As Boolean
    
    tmrCheckShortKeys.Enabled = False
    
    isStillRecording = IsRecordingMode
    
    isStillPlaybacking = IsPlayingMode
    
    If isStillRecording Then
       
       StopMacro , True, True
    
    ElseIf isStillPlaybacking Then
             
       StopMacro , True, True
    
    End If
    
    If Not CobMacros.ListCount >= 1 Then

Rem if no item in order to avoid error in the next Form_Load

Rem when you getsetting(,"las","las"
      
      SaveSetting App.EXEName, "las", "las", "nono"
        
    Else
    
      SaveSetting App.EXEName, "las", "las", CobMacros.ListIndex
    
    End If
     
'    Unload frmOptions
    Unload frmOptions2
    Unload frmAbout
    If TrayIcon1.IsDisplayed Then
       TrayIcon1.Hide
    End If
End Sub

Private Sub menuabout_Click()
  
     frmAbout.Show , frmJMacro
  
End Sub

Private Sub menuSavedMacros_Click(index As Integer)

    Dim TipState As String
    Dim FrmJMacroState As Boolean
    Dim SelectedMacroName As String
     
    TipState = GetSetting(App.EXEName, "k", "d", "")
    
         '/////////////////////////
         Dim key12 As String
         Dim key13 As String
         Dim key14 As String
         Dim dkey1 As String * 15
         Dim DefChar7 As String
         Dim DefChar9 As String
         Dim DefChar4 As String
         
         DefChar4 = Str(vbKeyPause)
         DefChar9 = Str(vbKeyPause)
         DefChar7 = Str(vbKeyF11)
         GetINIProfileString Resume_Play_back, "0,1,0#" & DefChar9 & "*", dkey1, Len(dkey1)
         key12 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         
         GetINIProfileString Stop_Play_back, "0,1,0#" & DefChar7 & "*", dkey1, Len(dkey1)
         key13 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         
         GetINIProfileString Pause_Play_back, "0,0,0#" & DefChar4 & "*", dkey1, Len(dkey1)
         key14 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         '////////////////////////////////////////////
         
         
    If Not TipState = "one" Then
        
        frmTip.lblTip.Caption = "Press " & DecodeKey(key13) & " to stop Play back." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Press " & DecodeKey(key14) & " to Pause Play back." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Press " & DecodeKey(key12) & " to Resume Play back."
      
         frmTip.Show 1
     
    End If
     
     
Rem the Story for the only next line of code is very Important
Rem while i test the program , i found that if click the playback button from
Rem the FrmJmacro From,the playback repeat itself due to sending a click from
Rem Space button .......etc
Rem i solve this problem by minimize FrmJmacro when play back get started
Rem and we make this in line 20 and line 30
'========================================================
Rem The Brief:
Rem we minimize the FrmJmacro if she was displayed in Line 20
Rem Then we restore it to it's orginial state in Line 30
'========================================================
20  If Me.WindowState <> 1 Then Me.WindowState = 1: FrmJMacroState = True
    
    SelectedMacroName = AppPath & "Data\" & menuSavedMacros.Item(index).Caption & ".pmf"
    
    Call OpenMacro(SelectedMacroName, menuSavedMacros.Item(index).Caption)

    StartPlay

    WatchForCancel
    
    If PlayBackPauseForValerie_Dion(2) = 1 Then
           
        TrayIcon1.ShowBalloon "Press " & DecodeKey(key12) & " to resume.", "Replaying " & Chr(34) & menuSavedMacros.Item(index).Caption & Chr(34) & " Macro Paused", Info, 3000
                
    ElseIf PlayBackPauseForValerie_Dion(2) = 0 Then
    
           TrayIcon1.ShowBalloon "Replaying " & menuSavedMacros.Item(index).Caption & " macro Finish", App.Title, Info, 3000
      
    End If
    
30          If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE
  
    
End Sub

Private Sub mnu_jmacro_Click()
     
     If mnu_jmacro.Caption = "Show " & SoftwareName Then
          
        Me.Visible = True
        
        Me.WindowState = 0
       
        mnu_jmacro.Caption = "Hide " & SoftwareName
        
    ElseIf mnu_jmacro.Caption = "Hide " & SoftwareName Then
         
         Me.Visible = False
     
         mnu_jmacro.Caption = "Show " & SoftwareName
       
     End If
    
End Sub

Private Sub mnuExit_Click()

Dim isStillRecording As Boolean
Dim isStillPlaybacking As Boolean
    
    tmrCheckShortKeys.Enabled = False
    
    isStillRecording = IsRecordingMode
    
    isStillPlaybacking = IsPlayingMode
    
    If isStillRecording Then
       
       StopMacro , True, True
    
    ElseIf isStillPlaybacking Then
             
       StopMacro , True, True
    
    End If
    
    If Not CobMacros.ListCount >= 1 Then

Rem if no item in order to avoid error in the next Form_Load

Rem when you getsetting(,"las","las"
      
      SaveSetting App.EXEName, "las", "las", "nono"
        
    Else
    
      SaveSetting App.EXEName, "las", "las", CobMacros.ListIndex
    
    End If
     
    If TrayIcon1.IsDisplayed Then
       TrayIcon1.Hide
    End If
    
    Unload frmJMacro

    
    
End Sub

Private Sub mnumanage_Click()
   
     frmManage.Show , frmJMacro

End Sub

Private Sub mnuoptions_Click()
   
   
     frmOptions2.Show , frmJMacro
     
     
End Sub


Private Sub mnuRecord_Click()
   
   cmdRecord_Click

End Sub

Public Sub FillComboBox()
'============================================================

' Routine Name: LoadComboBox
' Description: Load the comboBox with the saved Macros
' Programmer : Ahmed Saad
' Date: 20/08/2004 01:09:33 ã
' Copyright © 2004 CprinGold Software.
' Notes: we retrieve the Macros Name and paths from the regeidit

' Modification History:

'============================================================

    On Error GoTo Handler
    Dim Q As Long
    Dim i As Long
    Dim SavedMacros  As Variant
     
    CobMacros.Clear
    
    SavedMacros = GetAllSettings(App.EXEName, "sm")

Rem SavedMacros is a MultiDimenstional array,
Rem If you imagine it as a table you can consider
Rem that i is the Column Number and 0 is the Row number.


    For i = LBound(SavedMacros, 1) To UBound(SavedMacros, 1)
         
Rem Debug.Print SavedMacros(i, 0), SavedMacros(i, 1)
        
        CobMacros.AddItem SavedMacros(i, 0)
               
        
    Next i
    
Rem unload the last Loaded iten in the menu

    For Q = 0 To menuSavedMacros.Count - 1
        
        Rem here we unload both the systray pop up and frmjamcro popup menu
        Rem note that we make to pop up menu
        Rem one for systray and one for FrmJmacro
        If Not Q = 0 Then Unload menuSavedMacros(Q): Unload vismnuAvailbemacros(Q)
    
    Next Q
 
    For i = 0 To CobMacros.ListCount - 1
        
Rem     because we coudn't unload menu item no 0

        If Not i = 0 Then
        
              Load menuSavedMacros(i)
              
              Load vismnuAvailbemacros(i)
        
              menuSavedMacros(i).Caption = CobMacros.List(i)
             
              vismnuAvailbemacros(i).Caption = CobMacros.List(i)
              
        Else
           
             menuSavedMacros(0).Caption = CobMacros.List(i)
                      
             vismnuAvailbemacros(0).Caption = CobMacros.List(i)
        
        End If
    
    Next i


LoadComboBoxExit:

On Error Resume Next
    
    If CobMacros.ListCount > 0 Then
       
       cmdPlayBack.Enabled = True
       
       cmdEditMacro.Enabled = True
       
       mnuPlayback.Enabled = True
       
       vismnuplayback.Enabled = True
    
    Else
        
       cmdEditMacro.Enabled = False
       
       cmdPlayBack.Enabled = False
       
       mnuPlayback.Enabled = False
        
       vismnuplayback.Enabled = False
    
    End If

Rem Sets the index of the currently selected item.
If CobMacros.ListCount <> 0 Then

    CobMacros.ListIndex = 0
     
End If
   
    Exit Sub

Handler:

   
    If Err.Number <> 9 And Err.Number <> 13 Then
           
       Err.Raise Err.Number, SoftwareName & "::LoadComboBox ", Err.Description
       
    ElseIf Err.Number <> 13 Then
    
       Err.Raise Err.Number, SoftwareName & "::LoadComboBox ", Err.Description
    
    Else
       
       GoTo LoadComboBoxExit:
    
    End If
    GoTo LoadComboBoxExit

End Sub

Private Sub SetPlayBackSpeed()
'============================================================

' Routine Name: GetPlayBackSpeed
' Description:
' Author: Ahmed Saad
' Date: 23/08/2004 11:21:38 ã
' Copyright © 2004 CprinGold Software
' Notes:it set the playback speed and load option speed tab with that Speed
'

' Modification History:

'============================================================

    On Error GoTo GetPlayBackSpeedErr

Dim dkey As String * 2
Dim PlayBackSpeed As String

   GetINIProfileString "PlaybackSpeed", "n", dkey, Len(dkey)
   
'PlayBackSpeed = GetSetting(App.EXEName, "op", "sp", "n")

frmJMacro.cmdPlayBack.Tag = Left(dkey, 1)

'Load frmOptions

If PlayBackSpeed = "l" Then
        
'    frmOptions.optLowSpeed.Value = True
    
'    frmOptions.chkAlways.Value = 1
    
ElseIf PlayBackSpeed = "n" Then
       
'    frmOptions.optNormalSpeed.Value = True
      
'    frmOptions.chkAlways.Value = 1

ElseIf PlayBackSpeed = "h" Then
        
'    frmOptions.OptHigh.Value = True
    
'    frmOptions.chkAlways.Value = 1
     
End If

GetPlayBackSpeedExit:
    On Error Resume Next

    Exit Sub

GetPlayBackSpeedErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: " & SoftwareName & "_SetPlayBackSpeed" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo GetPlayBackSpeedExit

End Sub
Private Sub CheckShortKeys()
'============================================================

' Routine Name: CheckShortKeys
' Description:
' Author: Ahmed Saad
' Date: 26/08/2004 11:01:41 Õ
' Copyright © 2004 CprinGold Software
' Notes:

' Modification History:

'============================================================

On Error GoTo CheckShortKeysErr

Dim res As Integer
   
res = GetINIProfileInt("DASK", 0)
   
If res Then
    
   '//Shortcut Keys are disabled
    Exit Sub
     
End If
   
Dim g As Integer
Dim i As Integer
Dim Alt As String
Dim Ctrl As String
Dim mname As String
Dim Shift As String
Dim PressHit As String
Dim F10State As String
Dim F11State As String
Dim PressedKey As String
Dim MyMacroName As String
Dim SavedMacros As Variant
Dim RepeatCount As Variant
Dim isHitPressed As Boolean
Dim MacroFilePath As String
Dim GeneralHotKey As String
Dim ForeGroundWindow As Long
Dim FrmJMacroState As Boolean

Dim DecodedKey As String

Dim key1 As String, key2 As String

Dim dkey As String * 15  '//why 15, coz it larger than "1,1,1#256*

Dim DefChar11 As String
Dim DefChar9 As String
Dim PressedKey8 As Integer

DefChar11 = Str(vbKeyInsert)  '//Get the ASC value of N
DefChar9 = Str(vbKeyPause)  '//Get the ASC value of H
   PressedKey8 = GetPressedKey

   If GetAsyncKeyState(vbKeyControl) Then
      DecodedKey = 1
   Else
      DecodedKey = 0
   End If
   
   If GetAsyncKeyState(18) Then
      DecodedKey = DecodedKey & "," & 1
   Else
      DecodedKey = DecodedKey & "," & 0
   End If
   
   '//Tell Me why the space after the hash "# " ??
   '//for unknown reason, It added automatic in LoadShortcutKeys
   '//[i mean in GetINIProfileString lines in LoadShortcutKeys  ] and i don't know why
   '//and that's why i added it here ino order to match the keys
   
   If GetAsyncKeyState(vbKeyShift) Then
      DecodedKey = DecodedKey & "," & 1 & "# " & PressedKey8 & "*"
   Else
      DecodedKey = DecodedKey & "," & 0 & "# " & PressedKey8 & "*"
   End If

   Dim ResumePlayback As Boolean
   Dim ResumeRecording As Boolean
   
   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString Resume_Play_back, "0,1,0#" & DefChar9 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
         ResumePlayback = True
         Debug.Print Resume_Play_back
      End If
   End If
   
   
If Not IsPlayingMode Then
'if no playbacking running
   
 '  If GetAsyncKeyState(18) And GetAsyncKeyState(vbKeyInsert) Then
   If ResumePlayback Then
    
'element 0 = macropath
'element 1 = playlog
'element 2 = there is a paused macro or not

'//Hide balloon if it is still displayed
TrayIcon1.CloseBalloon

'PlayBackPauseForValerie_Dion(0) = ""
'PlayBackPauseForValerie_Dion(1) = 0
       'element 3 = there is a paused macro or not
        If PlayBackPauseForValerie_Dion(2) = 1 Then
           PlayBackPauseForValerie_Dion(2) = 0
           'Debug.Print "alt + pause"
           MyMacroName = PlayBackPauseForValerie_Dion(0)
           Call OpenMacro(MyMacroName)
           StartPlay True
           WatchForCancel
           ''TrayIcon1.ShowBalloon "PlayBack Over", App.Title, Info, 3000
           
       End If
   End If
   
End If


   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString Resume_Record, "0,1,0#" & DefChar11 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
         ResumeRecording = True
         Debug.Print Resume_Record
      End If
   End If
   
If IsRecordingMode Then
    ' If GetAsyncKeyState(vbKeyInsert) Then
     
    '     If GetAsyncKeyState(18) Then
    '     If GetAsyncKeyState(18) Then
          If ResumeRecording Then
          If RecordPauseForValerie_Dion(2) = 1 Then
             'the following line will be execuated only if
             'there is a recording Mode
             'User press ALT+PAUSE
             'there is a paused recording session waiting for resume
             RecordPauseForValerie_Dion(2) = 0
             Tags.RecordingIsOver = False
             TrayIcon1.Animate ImageList1, 100
          End If
          End If
    ' End If

End If

Rem Start Record if you press F10

F10State = GetSetting(App.EXEName, "noF", "noF", "")

F11State = GetSetting(App.EXEName, "noF", "no11", "")

Rem Alt+F9
Rem Note that:we use "noF" if the user select a hotkey for recording rather than Alt+F9
If GetAsyncKeyState(vbKeyF9) And GetAsyncKeyState(18) Then

      If F10State = "noF" Then Exit Sub

StartRecord:

      If cmdRecord.Caption = "&Record" Then
         
         ForeGroundWindow = GetForegroundWindow
         
Rem      here you save the Hwnd of the last ForeGroundedWindow Before hiting the key for
Rem      Start Record ok why...?
Rem      i will tell you
Rem      you will use this hwnd again after stoping recording in order to
Rem      return this form to its last state (ForeGrounded)
       
         '//pt 28Aug2007: commented next line coz i will use frmoption2 and the new Shortcut Model
         'cmdRecord_Click
       End If

Rem Alt+F10
Rem Note that:we use "no11" if the user select a hotkey for stop recording rather than Alt+F9
ElseIf GetAsyncKeyState(vbKeyF10) And GetAsyncKeyState(18) Then
      
      If F11State = "no11" Then Exit Sub

StopRecord:
        
          
      If cmdRecord.Caption = "&Stop Recording" Then
         ''RecordPauseForValerie_Dion(2) = 0
         '//pt 28Aug2007: commented next line coz i will use frmoption2 and the new Shortcut Model
         '//cmdRecord_Click
         
      End If

ElseIf GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyDelete) And GetAsyncKeyState(18) Then
       
       
       If F11State = "no11" Then Exit Sub

Rem StopRecord

      If cmdRecord.Caption = "&Stop Recording" Then
         ''RecordPauseForValerie_Dion(2) = 0
         '//pt 28Aug2007: commented next line coz i will use frmoption2 and the new Shortcut Model
         '//cmdRecord_Click
         
      End If
      
Rem to hide the program in the tray when hit windows key and m key

ElseIf GetAsyncKeyState(91) And GetAsyncKeyState(77) Then
 
       If mnu_jmacro.Caption = "Hide " & SoftwareName Then
          
              If Not cmdRecord.Caption = "&Stop Recording" Then
             
                  Me.Visible = False
            
              End If
            
               mnu_jmacro.Caption = "Show " & SoftwareName
         
       End If

ElseIf GetAsyncKeyState(vbKeyControl) Or GetAsyncKeyState(18) Or GetAsyncKeyState(vbKeyShift) Then
                   
    
       If GetAsyncKeyState(vbKeyControl) Then
       
          Ctrl = "Ctrl+"
           
       End If
         
Rem    Alt Key
       
       If GetAsyncKeyState(18) Then
       
          Alt = "Alt+"
           
       End If
   
       If GetAsyncKeyState(vbKeyShift) Then
       
          Shift = "Shift+"
       
       End If
        
       For i = 32 To 128
             
           If GetAsyncKeyState(i) And Not i = 116 Then
               
              '//pt 30Aug2007: changed next line to enable the user to user f1 and f2 and other keys to run macros
              '//PressHit = LCase(Chr(i))
              PressHit = TranslateShortCutKey(i)
              
              If PressHit = "none" Then
                 PressHit = LCase(Chr(i))
              End If
              
              isHitPressed = True
              
           End If
          
       Next i

       
       If isHitPressed Then

            PressedKey = Ctrl & Alt & Shift & PressHit
   
            SavedMacros = GetAllSettings(App.EXEName, "sm")
             
            For i = LBound(SavedMacros, 1) To UBound(SavedMacros, 1)
         
        
                If PressedKey = SavedMacros(i, 1) Then
                   
                           Dim res22 As Integer
                           Dim IsForEver As Boolean
                           res22 = GetINIProfileInt("repeat", 1)
            
                           Select Case res22
                                   
                            Case 1
                                        
                                tmrTerminatePlayback.Enabled = False '//just to make sure
                                txtRepeat = GetINIProfileInt("repeatcount", 1)
                                   
                            Case 2
                                                    
                                Debug.Print 1, time
                                
                                TickCountNow = GetTickCount
                                        
                                tmrTerminatePlayback.Enabled = True
                                         
                                IsForEver = True
                                
                                RepeatCount = 1E+21 '//1E+21 == 1000000000000000000000 forever ya3ny ;)
                                         
                            Case 3
                                         
                                tmrTerminatePlayback.Enabled = False '//just to make sure
                                RepeatCount = 1E+21 '//1E+21 == 1000000000000000000000 forever ya3ny ;)
                                IsForEver = True
                            
                           End Select
                        
                        If txtRepeat <> "" Then RepeatCount = txtRepeat Else RepeatCount = 1
                        
                        mname = SavedMacros(i, 0)
                        
                        MacroFilePath = AppPath & "Data\" & SavedMacros(i, 0) & ".pmf"
                                           
                        For g = 1 To RepeatCount
                              
                             Rem the Story for the only next line of code is very Important
                             Rem while i test the program , i found that if click the playback button from
                             Rem the FrmJmacro From,the playback repeat itself due to sending a click from
                             Rem Space button .......etc
                             Rem i solve this problem by minimize FrmJmacro when play back get started
                             Rem and we make this in line 20 and line 30
                             '========================================================
                             Rem The Brief:
                             Rem we minimize the FrmJmacro if she was displayed in Line 20
                             Rem Then we restore it to it's orginial state in Line 30
                             '========================================================
                             If Me.WindowState <> 1 Then Me.WindowState = 1: FrmJMacroState = True
                                                   
                             Call OpenMacro(MacroFilePath, mname)
                  
                             StartPlay
                  
                             WatchForCancel
                              
                             '//cprin 27Jul2006: added next line to fix the bug of "can't stop play back when press alt+f10 if repeatcount is more than 1"
                             If Tags.StopPlayBackLoop Then
                                '//reset the StopPlayBackLoop again
                                Tags.StopPlayBackLoop = False
                                '//Exit for because the user hit Alt+F10
                                Exit For
                             End If
                             
                        Next g
                
                        If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE
                       
                     '/////////////////////////
         Dim key12 As String
         Dim key13 As String
         Dim dkey1 As String * 15
         Dim DefChar7 As String
         DefChar7 = Str(vbKeyF11)
         GetINIProfileString Resume_Play_back, "0,1,0#" & DefChar9 & "*", dkey1, Len(dkey1)
         key12 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         
         GetINIProfileString Stop_Play_back, "0,1,0#" & DefChar7 & "*", dkey1, Len(dkey1)
         key13 = Mid(dkey1, 1, InStr(1, dkey1, "*"))
         '////////////////////////////////////////////
         
            If PlayBackPauseForValerie_Dion(2) = 1 Then
               TrayIcon1.ShowBalloon "Press " & DecodeKey(key12) & " to resume.", "Replaying " & Chr(34) & mname & Chr(34) & " Macro Paused", Info, 3000
            ElseIf PlayBackPauseForValerie_Dion(2) = 0 Then
               TrayIcon1.ShowBalloon "Replaying " & mname & " Macro Finish", App.Title, Info, 3000
            End If
            
            
    '  If PlayBackPauseForValerie_Dion(2) = 1 Then
    '         ElseIf PlayBackPauseForValerie_Dion(2) = 0 Then
    '        End If
           
                     
                End If
    
   
             Next i
            
            
               
            
CheckShortKeysExit:

            For i = 2 To 7
                  
                GeneralHotKey = GetSetting(App.EXEName, "GHK", i, "")
                
                If PressedKey = GeneralHotKey Then
                   
                   Select Case i
                          
                          Case 2
                               
                               GoTo StartRecord:
                          
                          Case 3
                               
                               GoTo StopRecord:

                          Case 4
                               
                           '  keybd_event vbKeyF11, 0, &H2, 0

                          
                          Case 5
                          
                            '//pt 28Aug2007: commented next line coz i will use frmoption2 and the new Shortcut Model
                            'Me.Show
                            'Me.SetFocus
                            
                          Case 6
                         
                              '//pt 26Aug2007: commented next line coz i will use frmoption2 and the new Shortcut Model
                              'frmOptions.Show
                                 
                          Case 7
                            
                              '//pt 26Aug2007: commented next line coz i will use frmoption2 and the new Shortcut Model
                              'frmManage.Show
                          
                   End Select
                
                End If
            
            Next i
            
      
       End If

End If


    On Error Resume Next

    Exit Sub

CheckShortKeysErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: " & SoftwareName & "_CheckShortKeys" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        'MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo CheckShortKeysExit

End Sub

Private Sub tmrTransperent_Timer()
FadeMe
tmrTransperent.Enabled = False
End Sub

Private Sub TrayIcon1_BalloonLeftClick()
  TrayIcon1.CloseBalloon
End Sub

Private Sub TrayIcon1_LeftMouseUp()

    If frmJMacro.Visible = False Then
          'prevent user form restore frmJmacro while recording
          If cmdRecord.Caption = "&Record" Then
             frmJMacro.Visible = True
             ShowWindow hWnd, SW_RESTORE
          End If
    End If

 End Sub

Private Sub TrayIcon1_RightMouseUp()
If cmdRecord.Caption = "&Record" Then
    frmJMacro.PopupMenu frmJMacro.JMacroMenu
End If
End Sub

Private Sub txtRepeat_Change()

If IsNumeric(txtRepeat) And Val(txtRepeat) > 0 Then


Else
    
    txtRepeat.Text = ""

End If

End Sub

Private Sub visEXE_Click()
  Command1_Click
End Sub

Private Sub visExport_Click()

'//pt 20Aug2007:
Dim PaidFor As Boolean
'//pt 20Aug2007:
SECUREBEGIN_C
'//pt 20Aug2007:
PaidFor = True


On Error GoTo quitme:

Dim OF As OFSTRUCT
Dim IsThere As Long
Dim FileData As String
Dim MacFileData As String
Dim BackupFilePath As String
Dim SelectedMacroName As String
Dim SelectedDecMacroName As String

If CobMacros.ListCount = 0 Then
   MsgBox "There is no macro to backup."
   Exit Sub
End If

SelectedMacroName = AppPath & "Data\" & CobMacros.List(CobMacros.ListIndex) & ".pmf"
    
'//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
'IsThere = OpenFile(SelectedMacroName, OF, OF_READ)
'//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
'CloseHandle IsThere

'//cprin 29Jul2006: changed next line because if file name is very large, OpenFile return -1
'If IsThere = -1 Then
'//cprin 9Feb2007: changed next line because PathFileExists fails due to Craig=>operationivyflag@gmail.com report
'//If PathFileExists(SelectedMacroName) = False Then
If FileExists(SelectedMacroName) = False Then
   '/cprin 17July2006: changed next MsgBox function to display a well form messagebox.
   'MsgBox "Data of [" & CobMacros.List(CobMacros.ListIndex) & "] macro missed or manually deleted by the user."
    MsgBox "Can't Found the data file of " & CobMacros.List(CobMacros.ListIndex) & " macro" & Chr(13) & Chr(10) _
    & "(" & SelectedMacroName & ")", vbCritical, App.Title
   Exit Sub
End If

IsThere = 0

SelectedDecMacroName = AppPath & "Data\" & CobMacros.List(CobMacros.ListIndex) & ".ptf"
          
'//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
'IsThere = OpenFile(SelectedDecMacroName, OF, OF_READ)

'//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
'CloseHandle IsThere
   
'//cprin 29Jul2006: changed next line because if file name is very large, OpenFile return -1
'If IsThere = -1 Then
'//cprin 9Feb2007: changed next line because PathFileExists fails due to Craig=>operationivyflag@gmail.com report
'//If PathFileExists(SelectedMacroName) = False Then
If FileExists(SelectedMacroName) = False Then
    '/cprin 17July2006: changed next MsgBox function to display a well form messagebox.
    'MsgBox "Data of [" & CobMacros.List(CobMacros.ListIndex) & "] macro missed or manually deleted by the user."
    MsgBox "Can't Found the data file of " & CobMacros.List(CobMacros.ListIndex) & " macro" & Chr(13) & Chr(10) _
    & "(" & SelectedDecMacroName & ")", vbCritical, App.Title
    Exit Sub
End If

With CommonDialog1
     .FileName = CobMacros.List(CobMacros.ListIndex)
     .Filter = SoftwareName & " backup File (*.pbak)|*.pbak"
     .CancelError = True
     .flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist  '+ cdlOFNExtensionDifferent
     'cprin 13July2006: added next line to add a title for showsave button
     .DialogTitle = "Backup Macro"
     .ShowSave
End With

BackupFilePath = CommonDialog1.FileName

Open SelectedDecMacroName For Binary As #1
        
     FileData = Input(LOF(1), 1)
        
Close #1

FileData = FileData & "cprinsosonono" & Len(CobMacros.List(CobMacros.ListIndex)) & CobMacros.List(CobMacros.ListIndex)

Open SelectedMacroName For Binary As #2
        
    MacFileData = Input(LOF(2), 2)
        
Close #2
 
FileData = FileData & MacFileData
 
Open BackupFilePath For Binary As #3
        
    Put #3, , FileData
        
Close #3
 
Exit Sub
quitme:

'//pt 20Aug2007:
SECUREEND_C

'//pt 20Aug2007:
If PaidFor = False Then
   
   MsgBox "Sorry, this function is only available to registered users."

End If


End Sub

Private Sub visImport_Click()

'//pt 20Aug2007:
Dim PaidFor As Boolean
'//pt 20Aug2007:
SECUREBEGIN_A
'//pt 20Aug2007:
PaidFor = True

On Error GoTo quitme:
Dim i As Integer
Dim res As Variant
Dim marker As Variant
Dim MacData As String
Dim DecData As String
Dim NameLen As Integer
Dim MacroName As String
Dim IsBandleFile As Long
Dim BackUpData As String
Dim DecFilePath As String
Dim MacFilePath As String
Dim Iscprinsosonono As String
 
 Dim ahmed As String
With CommonDialog1
     .Filter = SoftwareName & " backup File (*.pbak)|*.pbak"
     .CancelError = True
     '//cprin 30Jul2006: change next line to allow multiselection in the showopen
     '//.flags = cdlOFNPathMustExist '+ cdlOFNAllowMultiselect
     .flags = cdlOFNPathMustExist '+ cdlOFNExplorer '+ cdlOFNAllowMultiselect
     'Note for the pervious line: we must add cdlOFNExplorer to the flags when use cdlOFNAllowMultiselect
     'otherwise we showopen will be the silly old ShowOpen Dialog Box.
     'cprin 13July2006 added next line to display ShowOpen dialog withno defualt file name
     .FileName = ""
     .DialogTitle = "Import Macro to " & SoftwareName
     .ShowOpen
End With

'Dim ExitLoop As Boolean
'Dim MacrosDir As String
'Dim WhereSpace As Integer
'Dim WhereBackSlash As Integer
'Dim TempWhereBackSlash As Integer
'
'WhereBackSlash = 1

'Do Until ExitLoop
'    WhereBackSlash = InStr(WhereBackSlash + 1, CommonDialog1.FileName, "\")
'    If WhereBackSlash <> 0 Then TempWhereBackSlash = WhereBackSlash
'    If WhereBackSlash = 0 Then ExitLoop = True
'Loop

'ExitLoop = False

'MacrosDir = Left(CommonDialog1.FileName, TempWhereBackSlash)

Open CommonDialog1.FileName For Binary As #2
        
     BackUpData = Input(LOF(2), 2)
        
Close #2

'//cprin 27Jul2006: First we will check if this is a bundle file that contains more than one macro
IsBandleFile = InStr(1, BackUpData, "cpringold79#%^&***&(@#$%ewrwear")

'//cprin 27Jul2006: if true then this is a bandle file that contains more than one macro
If IsBandleFile Then
   '//Pass the data in order to import the bandle to PMR
   ImportBandleToPMR (BackUpData)
   '//Exit Sub Since no need to process the following lines
   Exit Sub
End If

marker = InStr(1, BackUpData, "cprinsosonono")

If marker = 0 Then
   '//cprin 17July2006: changed next line to display a well form for MsgBox
   'MsgBox "invaild backup file."
   MsgBox "Invaild backup file.", vbCritical, SoftwareName
   Exit Sub
End If

NameLen = Mid(BackUpData, marker + 13, 1)

''''''''''''''''''''''''''''''''''''''Fix of NameLen exceeds one digits bug''''''''
'//cprin 29Jul2006: added next code block to fix the bug that appears if the namelen
'//is more than 1 digits for ex 10 or 100
Dim NextChar As String

'//Since we doesn't allow the macro name to start with numbers hence
'//we will check if the next char is a number thus if it is a number then
'//name len exceeds 9
If IsNumeric(Mid(BackUpData, marker + 14, 1)) Then
     Dim NameLen2 As Integer
     NameLen2 = Mid(BackUpData, marker + 13, 2)
     '//Since we doesn't allow the macro name to start with numbers hence
     '//we will check if the next char is a number thus if it is a number then
     '//name len exceeds 99
     If IsNumeric(Mid(BackUpData, marker + 15, 1)) Then
        Dim NameLen3 As Integer
        NameLen3 = Mid(BackUpData, marker + 13, 3)
     End If
End If

If NameLen2 > NameLen Then
     NameLen = NameLen2
     marker = marker + 1   '//shift the marker one digit because we have another digit
End If

If NameLen3 > NameLen2 Then
     NameLen = NameLen3
     marker = marker + 1   '//shift the marker one digit because we have another digit
End If
'''''''''''''''''''''''''''''''''''''''''End of the fix'''''''''''''''''''''''''''''''

MacroName = Mid(BackUpData, marker + 14, NameLen)

For i = 0 To CobMacros.ListCount - 1
   
       If MacroName = CobMacros.List(i) Then
          '/cprin 17July2006: changed next line to display a well form messagebox
          'Res = MsgBox(MacroName & " macro already exists. do you want to overwrite it?", vbYesNo)
          res = MsgBox("A macro with the name " & CobMacros.List(i) & _
          " Already Exist, Do you want to overwrite it?", vbYesNo, App.Title)
          If res = vbNo Then
             Exit Sub
          End If
       End If
   
Next i

MacFilePath = AppPath & "Data\" & MacroName & ".pmf"

DecFilePath = AppPath & "Data\" & MacroName & ".ptf"
   
On Error Resume Next
Kill MacFilePath
Kill DecFilePath
On Error GoTo 0

MacData = Mid(BackUpData, marker + 13 + NameLen + 1, Len(BackUpData) - marker + 13 + NameLen)

Open MacFilePath For Binary As #4
     
     Put #4, 1, MacData
     
Close #4
   
DecData = Left(BackUpData, marker - 1)


Open DecFilePath For Binary As #4
     
     Put #4, 1, DecData
     
Close #4
   
SaveSetting App.EXEName, "sm", MacroName, ""

FillComboBox
   
Exit Sub
quitme:

'//pt 20Aug2007:
SECUREEND_A

'//pt 20Aug2007:
If PaidFor = False Then
   
   MsgBox "Sorry, Importing Macros is only available to registered users."

End If


End Sub

Private Sub vismnuabout_Click()
mnuabout_Click
End Sub

Private Sub vismnuAvailbemacros_Click(index As Integer)
menuSavedMacros_Click index
End Sub

Private Sub vismnuCprinGoldonline_Click()
mnuweb_Click
End Sub

Private Sub vismnuExit_Click()
mnuExit_Click
End Sub

Private Sub vismnuhelpt_Click()
mnuHelp2_Click
End Sub

Private Sub vismnuHide_Click()
mnu_jmacro_Click
End Sub

Private Sub vismnuLanguage_Click()
cmdLanguage_Click
End Sub

Private Sub vismnumanage_Click()
mnumanage_Click
End Sub

Private Sub vismnuoptions_Click()
mnuoptions_Click
End Sub

Private Sub vismnuRecord_Click()
mnuRecord_Click
End Sub

Private Sub vismnuSendReviews_Click()
mnuSendReviews_Click
End Sub

'//cprin added next function to BackupAll the macros
Private Sub BackUpAll()

End Sub

'//pt 26Nov2007: copycat to Command1_Click with some changes to evulate exe files with limitations
Private Sub EvulateExeFiles()

Dim marker As Variant
Dim TempOne As String
Dim MacroData As String
Dim StandAloneData As String
Dim StandAloneMacro As String
Dim cprsosonono3345554 As Long
Dim SelectedMacroName As String
Dim OldCap As String

On Error GoTo quitme:

If CobMacros.ListIndex = -1 Then
   MsgBox "There is no macro to compile."
   Exit Sub
End If

With CommonDialog1
     .FileName = CobMacros.List(CobMacros.ListIndex)
     .Filter = "Executable File (*.exe)|*.exe"
     .CancelError = True
     .flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist '+ cdlOFNExtensionDifferent
     .ShowSave
End With

TempOne = GetTempPathName & CommonDialog1.FileTitle

StandAloneMacro = TempOne

OldCap = Me.Caption

Me.Caption = "Wait..."

If right(StandAloneMacro, 4) <> ".exe" Then
   StandAloneMacro = StandAloneMacro & ".exe"
End If

TryAgain:

SelectedMacroName = AppPath & "Data\" & CobMacros.List(CobMacros.ListIndex) & ".pmf"


Open StandAloneMacro For Output As #1
     
     Print #1, StrConv(LoadResData(102, "CUSTOM"), vbUnicode)
     
Close #1 ' there is a 2 close downthere


'check if the resource file is still vaild or invaild by a funken hacker.

'Sleep 2000 'sleep until creating the file

'Shell StandAloneMacro & " /check"

'cprsosonono3345554 = FindWindow(vbNullString, "cprsosonono3345554")

'If cprsosonono3345554 <> 0 Then
'Else
'  Close #1
'  Kill StandAloneMacro
'  MsgBox "Critical Error,please reinstall " & SoftwareName & ".", vbCritical, SoftwareName
'  Exit Sub
'End If
'ok the file in vaild


Open StandAloneMacro For Binary As #2
        
     StandAloneData = Input(LOF(2), 2)
        
Close #2


Open SelectedMacroName For Binary As #3
        
    MacroData = Input(LOF(3), 3)
        
Close #3
    
marker = InStr(1, StandAloneData, "cprinsosonono")

StandAloneData = Left(StandAloneData, marker - 1)

StandAloneData = StandAloneData & "cprinsosonono" & MacroData


Open StandAloneMacro For Binary As #4
     
     Put #4, 1, StandAloneData
     
Close #4

FileCopy TempOne, CommonDialog1.FileName

'//cprin 17July2006: changed next line to display a well form MsgBox
MsgBox "Executable macro" & " (" & CommonDialog1.FileTitle & ") " & "saved to " & CommonDialog1.FileName

Me.Caption = OldCap

On Error Resume Next

Kill TempOne

Exit Sub
quitme:

''MsgBox Err.Description, vbCritical, App.Title


'If Err.Number = cdlDialogFailure Then
'   StandAloneMacro = AppPath & CobMacros.List(CobMacros.ListIndex) & ".exe"
'   MsgBox "Save as Dialog Failure, Standalone file saved automatically to " & StandAloneMacro
'   GoTo TryAgain:
'ElseIf Err.Number = 326 Then
'   Close #1
On Error Resume Next
Dim Pathto As Boolean
'//cprin 9Feb2007: changed next line because PathFileExists fails due to Craig=>operationivyflag@gmail.com report
'//Pathto = PathFileExists(TempOne)
Pathto = FileExists(TempOne)
If Pathto Then
   Kill StandAloneMacro
End If
'   MsgBox "Error please reinstall " & SoftwareName & ".", vbCritical, SoftwareName
'End If
    
End Sub
