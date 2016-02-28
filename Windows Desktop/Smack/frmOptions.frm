VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   Caption         =   "Smack Options"
   ClientHeight    =   5160
   ClientLeft      =   2580
   ClientTop       =   1515
   ClientWidth     =   5040
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmOptions.frx":08CA
   ScaleHeight     =   5160
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   """&PlayBack Speed"""
      TabPicture(0)   =   "frmOptions.frx":CAF8C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraSample1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Startup"
      TabPicture(1)   =   "frmOptions.frx":CAFA8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Shortcut &Keys"
      TabPicture(2)   =   "frmOptions.frx":CAFC4
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   3855
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   4575
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Stop Playback"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   2400
            Width           =   1695
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Resume Playback"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   2040
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pause Playback"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1680
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pause Recording"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Resume Recording"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "C&hange"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3120
            TabIndex        =   27
            Top             =   2760
            Width           =   1335
         End
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Stop Recording"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   1320
            Width           =   1455
         End
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Start Recording"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Break Playbacking :"
            Height          =   375
            Index           =   4
            Left            =   2520
            TabIndex        =   24
            Top             =   360
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Smack"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   23
            Top             =   2760
            Width           =   1335
         End
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Options :"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   22
            Top             =   3120
            Width           =   1215
         End
         Begin VB.OptionButton Optchange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Manage :"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   21
            Top             =   3480
            Width           =   1575
         End
         Begin VB.CommandButton cmdDefualt 
            Caption         =   "&Defualt"
            Height          =   375
            Left            =   3120
            TabIndex        =   20
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "ALT+F10"
            Height          =   255
            Left            =   1920
            TabIndex        =   43
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "ALT+Insert"
            Height          =   255
            Left            =   1920
            TabIndex        =   41
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Pause"
            Height          =   255
            Left            =   1920
            TabIndex        =   40
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "ALT+Insert"
            Height          =   255
            Left            =   1920
            TabIndex        =   36
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Pause"
            Height          =   255
            Left            =   1920
            TabIndex        =   35
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "F9"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   34
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "F10"
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   33
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+Alt+m"
            Height          =   255
            Index           =   7
            Left            =   1920
            TabIndex        =   32
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+Alt+o"
            Height          =   255
            Index           =   6
            Left            =   1920
            TabIndex        =   31
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "F10"
            Height          =   375
            Index           =   4
            Left            =   3600
            TabIndex        =   30
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl+Alt+s"
            Height          =   255
            Index           =   5
            Left            =   1920
            TabIndex        =   29
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "PLAYBACK KEY"
            Height          =   255
            Left            =   3000
            TabIndex        =   28
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   4575
         Begin VB.CheckBox chkLoadTipsAtStartup 
            BackColor       =   &H00FFFFFF&
            Caption         =   "S&how Tips at Startup."
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.CheckBox chkPlaySound 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PlaySound when stop Recording or Playbacking"
            Height          =   375
            Left            =   240
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   17
            Top             =   1680
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.CheckBox chkStartUp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Run Smack when computer Startup."
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   16
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame fraSample1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Choose a Playback Speed"
         Height          =   2385
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   4575
         Begin VB.OptionButton OptHigh 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&High Speed"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton optNormalSpeed 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Normal Speed"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton optLowSpeed 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Low Speed"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   1200
            Width           =   3135
         End
         Begin VB.CheckBox chkAlways 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Al&ways PlayBack with this Speed"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1800
            Width           =   3375
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   90
         Left            =   23760
         TabIndex        =   8
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   4920
      X2              =   120
      Y1              =   4560
      Y2              =   4560
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
     
      savestring HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Smart Macro", AppExeName & Space(1) & "hide"
      
Else
      
      SaveSetting App.EXEName, "chk", "chk", ""

      savestring HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Smart Macro", ""


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

frmSetShortKey.Show 1, frmOptions

Unload frmManage

Key = Me.Tag

Rem why we will do the following
Rem if i only save selected Options ,this may
Rem make error when getsettings in frmJMacro.tmrCheckShortKeys
Rem because he didn't found the remaining keys for ex: i select 4 and
Rem save 4 he getsettings 2,3,4,5,6,7  and he willn't found 2,3,5,6,7 that's it

For i = 2 To 7

   If i = Selected Then
      
      lbl(i).caption = Key
      
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

Private Sub cmdOK_Click()
    
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
     
     SaveReg = True
   
  
End Sub




Private Sub Optchange_Click(Index As Integer)

cmdApply.Enabled = True

If Optchange(Index).Value = True Then

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
      
         cmdOK.Tag = "always"
      
      Else
         cmdOK.Tag = "once"
        
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
    
    If cmdOK.Tag = "always" Then
       
       chkAlways.Value = 1
    
    ElseIf cmdOK.Tag = "once" Then
      
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

