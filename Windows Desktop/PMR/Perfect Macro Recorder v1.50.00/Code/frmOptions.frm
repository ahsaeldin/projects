VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Old Options"
   ClientHeight    =   3225
   ClientLeft      =   2580
   ClientTop       =   1515
   ClientWidth     =   5040
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkMinToSysTray 
      Caption         =   "Minimize to system tray on startup"
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
      Height          =   375
      Left            =   360
      TabIndex        =   58
      Top             =   4080
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.OptionButton Optchange 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Break Playbacking :"
      Height          =   255
      Index           =   4
      Left            =   12480
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "General"
      Height          =   585
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   4815
      Begin VB.CheckBox chkStartUp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A&uto run on windows startup"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   57
         Top             =   240
         Width           =   4215
      End
   End
   Begin JwldButn2b.JeweledButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
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
   End
   Begin JwldButn2b.JeweledButton cmdApply 
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Apply"
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
   End
   Begin JwldButn2b.JeweledButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Global Hot Keys"
      Height          =   5175
      Left            =   5520
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   5295
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "My Macros"
            Enabled         =   0   'False
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Options "
            Enabled         =   0   'False
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Perfect Macro Recorder"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+Alt+s"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   50
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+Alt+o"
            Height          =   255
            Index           =   6
            Left            =   3000
            TabIndex        =   49
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+Alt+m"
            Height          =   255
            Index           =   7
            Left            =   3000
            TabIndex        =   48
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   120
         TabIndex        =   37
         Top             =   3120
         Width           =   5295
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pause Play back"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Resume Play back"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Stop Play back"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Pause"
            Height          =   255
            Left            =   3000
            TabIndex        =   43
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Alt+Insert"
            Height          =   255
            Left            =   3000
            TabIndex        =   42
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Alt+F10"
            Height          =   255
            Left            =   3000
            TabIndex        =   41
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   5295
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Record"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Stop Record"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Resume Record"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pause Record"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "F10"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   36
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "F9"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   35
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label lbl3333 
            BackStyle       =   0  'Transparent
            Caption         =   "Pause"
            Height          =   255
            Left            =   3000
            TabIndex        =   34
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblssss 
            BackStyle       =   0  'Transparent
            Caption         =   "Alt+Insert"
            Height          =   255
            Left            =   3000
            TabIndex        =   33
            Top             =   840
            Width           =   975
         End
      End
      Begin JwldButn2b.JeweledButton cmdChange 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "S&et"
         PictureSize     =   0
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         Enabled         =   0   'False
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
      End
      Begin JwldButn2b.JeweledButton cmdDefualt 
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Defualt"
         PictureSize     =   0
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         Enabled         =   0   'False
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
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   5400
         X2              =   120
         Y1              =   4560
         Y2              =   4560
      End
   End
   Begin VB.CheckBox chkPlaySound 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PlaySound when stop Recording or Playbacking"
      Height          =   375
      Left            =   7680
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      Top             =   9120
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "S&how Tips at Startup."
      Height          =   315
      Left            =   7680
      TabIndex        =   11
      Top             =   8640
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame fraSample1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Play Back Speed"
      Height          =   1785
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   4815
      Begin VB.CheckBox chkAlways 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Al&ways Use this Play Back Speed."
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   2775
      End
      Begin VB.OptionButton optLowSpeed 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Low Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   3135
      End
      Begin VB.OptionButton optNormalSpeed 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Normal Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton OptHigh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&High Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   90
         Left            =   23760
         TabIndex        =   5
         Top             =   7875
         Width           =   0
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
         Begin VB.Label Label12 
            Caption         =   "Show balloon tooltop when playback finish"
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   2520
            Width           =   3735
         End
         Begin VB.Label Label10 
            Caption         =   "Add this new option : start animation when playback or record"
            Height          =   495
            Left            =   0
            TabIndex        =   22
            Top             =   1800
            Width           =   3735
         End
         Begin VB.Label Label8 
            Caption         =   "Add this new option : when minimize the window what do you want 1. goto system tray or wahat"
            Height          =   615
            Left            =   0
            TabIndex        =   21
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label Label7 
            Caption         =   "Add this new option : when left click the icon in system tray what do you want to do 1. start new macro or what..."
            Height          =   615
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   3735
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Marcos wants Yes, I know this option. But it makes Smack start maximized and not minimized in the tray in the right corner below."
      Height          =   615
      Left            =   240
      TabIndex        =   56
      Top             =   3960
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label13 
      Caption         =   "Add this new option : when left click the icon in system tray what do you want to do 1. start new macro or what..."
      Height          =   1335
      Left            =   480
      TabIndex        =   55
      Top             =   6600
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label14 
      Caption         =   "Add this new option : when minimize the window what do you want 1. goto system tray or wahat"
      Height          =   735
      Left            =   360
      TabIndex        =   54
      Top             =   4560
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label15 
      Caption         =   "Add this new option : start animation when playback or record"
      Height          =   375
      Left            =   240
      TabIndex        =   53
      Top             =   5160
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label16 
      Caption         =   "Show balloon tooltop when playback finish"
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   5640
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label11 
      Caption         =   "Don't Display Balloon Tooltip for buttons"
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   6000
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "F10"
      Height          =   375
      Index           =   4
      Left            =   12720
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "PLAYBACK KEY"
      Height          =   255
      Left            =   12600
      TabIndex        =   26
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Tab Index"
      Height          =   255
      Left            =   12840
      TabIndex        =   24
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   4920
      X2              =   120
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaveReg As Boolean
Dim RegSettings(1 To 10) As String
Private Sub RetOldReg()
'============================================================

' Routine Name: RetOldReg
' Description:
' Author: Ahmed Saad
' Date: 28/08/2004 12:17:37 ã
' Copyright © 2004
' Notes:

' Modification History:

'============================================================

    On Error GoTo RetOldRegErr
    
    SaveSetting App.EXEName, "GHK", 2, RegSettings(5)
    
    SaveSetting App.EXEName, "GHK", 3, RegSettings(6)
    
    SaveSetting App.EXEName, "GHK", 4, RegSettings(7)
    
    SaveSetting App.EXEName, "GHK", 5, RegSettings(8)
    
    SaveSetting App.EXEName, "GHK", 6, RegSettings(9)
    
    SaveSetting App.EXEName, "GHK", 7, RegSettings(10)
    
    SaveSetting App.EXEName, "op", "sp", RegSettings(3)
    
    SaveSetting App.EXEName, "chk", "chk", RegSettings(1)
    
    SaveSetting App.EXEName, "noF", "noF", RegSettings(2)
     
    SaveSetting App.EXEName, "noF", "no11", RegSettings(4)
    
    
RetOldRegExit:
    On Error Resume Next

    Exit Sub

RetOldRegErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: Options_RetOldReg" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo RetOldRegExit

End Sub

Private Sub GetOldReg()
'============================================================

' Routine Name: GetOldReg
' Description:
' Author: Ahmed Saad
' Date: 28/08/2004 11:58:51 Õ
' Copyright © 2004
' Notes:

' Modification History:

'============================================================

    On Error GoTo GetOldRegErr
      
     
    RegSettings(1) = GetSetting(App.EXEName, "chk", "chk", "")
    
    RegSettings(2) = GetSetting(App.EXEName, "noF", "noF", "")
     
    RegSettings(3) = GetSetting(App.EXEName, "op", "sp", "")
      
    RegSettings(4) = GetSetting(App.EXEName, "noF", "no11", "")
      
    RegSettings(5) = lbl(2)
    
    RegSettings(6) = lbl(3)
    
    RegSettings(7) = lbl(4)
    
    RegSettings(8) = lbl(5)
    
    RegSettings(9) = lbl(6)
    
    RegSettings(10) = lbl(7)
    
  
    
GetOldRegExit:
    On Error Resume Next

    Exit Sub

GetOldRegErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: Options_GetOldReg" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo GetOldRegExit

End Sub

Private Sub chkAlways_Click()
Dim i As Integer

     cmdApply.Enabled = True
     
If chkAlways.Value = 1 Then

      
      If OptHigh.Value = True Then
          
         SaveSetting App.EXEName, "op", "sp", "h"

      ElseIf optNormalSpeed.Value = True Then
                         
         SaveSetting App.EXEName, "op", "sp", "n"
         
      ElseIf optLowSpeed.Value = True Then
         
         SaveSetting App.EXEName, "op", "sp", "l"
        
      Else
         
         SaveSetting App.EXEName, "op", "sp", "n"
      
      End If
     
        
Else
     
         SaveSetting App.EXEName, "op", "sp", "unchecked"
      

End If

End Sub

Private Sub chkLoadTipsAtStartup_Click()

cmdApply.Enabled = True

SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value

End Sub

Private Sub chkStartUp_Click()

Dim AppExeName As String

cmdApply.Enabled = True

If chkStartUp.Value = 1 Then
     
     SaveSetting App.EXEName, "chk", "chk", "ok"
     
     AppExeName = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & App.EXEName
             
     If Not right(AppExeName, 4) = ".exe" Then AppExeName = AppExeName & ".exe"
     
      '//cprin 10Aug2006: change next line to change registery entry from Smart Macro to PMR
      'savestring HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Smart Macro", AppExeName & Space(1) & "hide"
      savestring HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PMR", AppExeName & Space(1) & "hide"
Else
      
      SaveSetting App.EXEName, "chk", "chk", ""
      '//cprin 10Aug2006: change next line to change registery entry from Smart Macro to PMR
      'savestring HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Smart Macro", ""
      savestring HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PMR", ""


End If

End Sub

Private Sub cmdApply_Click()
    
    
    
    cmdApply.Enabled = False
    
End Sub

 Private Sub cmdCancel_Click()
               
     SetOldOptions
     
     Me.Hide

End Sub

Private Sub cmdChange_Click()
Dim i As Integer
Dim Key As String
Dim Selected As Integer
Dim IsContain As String
      

For i = 2 To 7

    If Optchange(i).Value = True Then
    
       Selected = i
       
       Rem to store the selected to use in frmSetShortKey.GetSelectedKey
       Rem to check if the user select "Alt+F9" or "Alt+F10"
       Rem why
       Rem if the user change Stop Record hot key from "Alt+F9" to another key
       Rem and after while decide to return ro use it again
       Rem frmSetShortKey.GetSelectedKey will generate Msgbox "already selected from StopPlaybackinh
       
       cmdApply.Tag = i
    
    End If
   

Next i

frmSetShortKey.Tag = "options"

'frmSetShortKey.Show 1, frmOptions

Unload frmManage

Key = Me.Tag

Rem why we will do the following
Rem if i only save selected Options ,this may
Rem make error when getsettings in frmJMacro.tmrCheckShortKeys
Rem because he didn't found the remaining keys for ex: i select 4 and
Rem save 4 he getsettings 2,3,4,5,6,7  and he willn't found 2,3,5,6,7 that's it


For i = 2 To 7

   If i = Selected Then
      
      lbl(i).Caption = Key
      
      SaveSetting App.EXEName, "GHK", i, Key
       
      If i = 2 Then

         If Key <> "Alt+F9" Then
         
            SaveSetting App.EXEName, "noF", "noF", "noF"
         
         Else

Rem      if the user want to reuse "Alt+F9" again so we will make "nof" = ""
            SaveSetting App.EXEName, "noF", "noF", ""
            
         End If
      
      ElseIf i = 3 Then
         
         If Key <> "Alt+F10" Then
            
            SaveSetting App.EXEName, "noF", "no11", "no11"
         
         Else
            
            SaveSetting App.EXEName, "noF", "no11", ""
         
         End If
      
      End If
      
   Else
       
      IsContain = GetSetting(App.EXEName, "GHK", i, "")
      
      If IsContain = "" Then
      
         SaveSetting App.EXEName, "GHK", i, ""
      
      End If
   
   End If

Next i

End Sub

Private Sub cmdDefualt_Click()

SaveSetting App.EXEName, "f", "a", 1
SaveSetting App.EXEName, "noF", "no11", ""
SaveSetting App.EXEName, "noF", "noF", ""
SaveSetting App.EXEName, "GHK", 2, "Alt+F9"
SaveSetting App.EXEName, "GHK", 3, "Alt+F10"
SaveSetting App.EXEName, "GHK", 4, "Alt+F10"
SaveSetting App.EXEName, "GHK", 5, "Ctrl+Alt+s"
SaveSetting App.EXEName, "GHK", 6, "Ctrl+Alt+o"
SaveSetting App.EXEName, "GHK", 7, "Ctrl+Alt+m"
Form_Activate



End Sub

Private Sub cmdOk_Click()
    
     Me.Hide
     
End Sub


Private Sub Form_Activate()
  
  Dim i As Integer

  Dim TipsState As String
 
  Dim LastHotKeys As String
  
  Dim StartUpState As String
   
  
    cmdApply.Enabled = False
     
    StartUpState = GetSetting(App.EXEName, "chk", "chk", "")
      
    If StartUpState = "ok" Then
      
      chkStartUp.Value = 1
   
    End If
    
    TipsState = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 0)
      
    chkLoadTipsAtStartup.Value = TipsState
   
    SaveOldOptions

    For i = 2 To 7
       
      LastHotKeys = GetSetting(App.EXEName, "GHK", i, "")
      
      If LastHotKeys <> "" Then
               
         lbl(i) = LastHotKeys
         
      End If
      
    Next i

If SaveReg Then

   GetOldReg
   
   SaveReg = False
   
End If
End Sub
Private Sub Form_Load()
    'cprin 18July2006: added next code block
    'to ensure that jeweled buttons will remain
    'as i want.
    ''''''''''''''''''''''''''''''''''''''''''''
    cmdOk.Caption = "&OK"
    cmdApply.Caption = "&Apply"
    cmdCancel.Caption = "&Cancel"
    cmdChange.Caption = "C&hange"
    cmdDefualt.Caption = "&Defualt"
    ''''''''''''''''''''''''''''''''''''''''''''

    SaveReg = True
End Sub











Private Sub Optchange_Click(index As Integer)

cmdApply.Enabled = True

If Optchange(index).Value = True Then

   cmdChange.Enabled = True

Else

   cmdChange.Enabled = False

End If

End Sub

Private Sub OptHigh_Click()
                   
          cmdApply.Enabled = True
          
          frmJMacro.cmdPlayBack.Tag = "h"
          
Rem        if the User choose this while he checked the chkAlways
Rem        so we call it for Update the Regedit
           
           chkAlways_Click
           
End Sub

Private Sub optLowSpeed_Click()
        
          cmdApply.Enabled = True
          
          frmJMacro.cmdPlayBack.Tag = "l"

Rem        if the User choose this while he checked the chkAlways
Rem        so we call it for Update the Regedit
                     
           chkAlways_Click
End Sub

Private Sub optNormalSpeed_Click()
   
         cmdApply.Enabled = True
          
         frmJMacro.cmdPlayBack.Tag = "n"
Rem        if the User choose this while he checked the chkAlways
Rem        so we call it for Update the Regedit
        
          chkAlways_Click
End Sub

Private Sub SaveOldOptions()
'============================================================

' Routine Name: SaveOldOptions
' Description:
' Author: Ahmed Saad
' Date: 24/08/2004 01:06:29 Õ
' Copyright © 2004 CprinGold Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo SaveOldOptionsErr

      Dim i As Integer
      
      
      
      If OptHigh.Value = True Then
          
         cmdCancel.Tag = "h"

      ElseIf optNormalSpeed.Value = True Then
                         
         cmdCancel.Tag = "n"
 
      ElseIf optLowSpeed.Value = True Then
         
         cmdCancel.Tag = "l"

      Else
         
         cmdCancel.Tag = "none"

      End If
     
      If chkAlways.Value = 1 Then
      
         cmdOk.Tag = "always"
      
      Else
         cmdOk.Tag = "once"
        
      End If
       
      If chkStartUp.Value = 1 Then
      
         chkStartUp.Tag = "sel"
      
      Else
         
         chkStartUp.Tag = ""
      
      End If
         
      chkLoadTipsAtStartup.Tag = chkLoadTipsAtStartup.Value
      
      SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value


      For i = 2 To 7
           
          If Optchange(i).Value = True Then
             
             cmdCancel.Tag = i
             
          End If
           
      Next i
        
SaveOldOptionsExit:
    On Error Resume Next

    Exit Sub

SaveOldOptionsErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: Options_SaveOldOptions" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo SaveOldOptionsExit

End Sub

Private Sub SetOldOptions()
'============================================================

' Routine Name: SetOldOptions
' Description:
' Author: Ahmed Saad
' Date: 24/08/2004 01:12:23 Õ
' Copyright © 2004 CprinGold Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo SetOldOptionsErr
    Dim i As Integer
    
    RetOldReg
    
    If cmdCancel.Tag = "h" Then
       
       OptHigh.Value = True
    
    ElseIf cmdCancel.Tag = "n" Then
      
       optNormalSpeed.Value = True
    
    ElseIf cmdCancel.Tag = "l" Then
       
       optLowSpeed.Value = True
        
    ElseIf cmdCancel.Tag = "none" Then
    
       OptHigh.Value = False
      
       optNormalSpeed.Value = False
       
       optLowSpeed.Value = False
      
    End If
    
    If cmdOk.Tag = "always" Then
       
       chkAlways.Value = 1
    
    ElseIf cmdOk.Tag = "once" Then
      
       chkAlways.Value = 0
    
    End If
    
    
      If chkStartUp.Tag = "sel" Then
      
         chkStartUp.Value = 1
      
      Else
         
         chkStartUp.Value = 0
      
      End If
         

      For i = 2 To 7
           
          If Val(cmdCancel.Tag) = i Then
             
             Optchange(i).Value = True
             
          End If
           
      Next i
    
      chkLoadTipsAtStartup.Value = chkLoadTipsAtStartup.Tag
      
SetOldOptionsExit:
    On Error Resume Next

    Exit Sub

SetOldOptionsErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: Options_SetOldOptions" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo SetOldOptionsExit

End Sub

