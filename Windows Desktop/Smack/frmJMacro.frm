VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4E8B3075-42A6-11DA-A1C2-803C58C10000}#1.0#0"; "astc.ocx"
Begin VB.Form frmJMacro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smack 1.06"
   ClientHeight    =   3015
   ClientLeft      =   5055
   ClientTop       =   3495
   ClientWidth     =   5640
   Icon            =   "frmJMacro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmJMacro.frx":08CA
   ScaleHeight     =   3015
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ASTC.TrayIcon TrayIcon1 
      Left            =   3360
      Top             =   1680
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Text            =   "fade in first time only"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer tmrTransperent 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3000
      Top             =   1320
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStandaLone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xport to EXE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdAliasRecord 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Record New Macro"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Record your Mouse  and keyboards events"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdLanguage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Language"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Help"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdreviews 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJMacro.frx":655EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJMacro.frx":65906
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJMacro.frx":65C20
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtRepeat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   3
      ToolTipText     =   "Set the Repeat count of playback."
      Top             =   1320
      Width           =   855
   End
   Begin VB.Timer tmrCheckShortKeys 
      Interval        =   1
      Left            =   3360
      Top             =   2280
   End
   Begin VB.CommandButton cmdmanage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Manage"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Manage your Recorded Macros"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdoptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Options"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdPlayBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Playback"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Playback the Macro selected in Combo Box"
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox CobMacros 
      Height          =   315
      ItemData        =   "frmJMacro.frx":65F3A
      Left            =   1920
      List            =   "frmJMacro.frx":65F3C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "choose macro for playback"
      Top             =   900
      Width           =   1935
   End
   Begin VB.CommandButton cmdRecord 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Record New Macro"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Record your Mouse  and keyboards events"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "write that in the regeister version you can record as much as you can"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Repeat Count"
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      Visible         =   0   'False
      X1              =   5640
      X2              =   5880
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   2775
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5415
   End
   Begin VB.Menu JMacroMenu 
      Caption         =   "Smack"
      Visible         =   0   'False
      Begin VB.Menu mnu_jmacro 
         Caption         =   "Hide Smack"
      End
      Begin VB.Menu a 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuRecord 
         Caption         =   "Record New Macro"
      End
      Begin VB.Menu mnuPlayback 
         Caption         =   "Pla&yBack Macro"
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
         Caption         =   "Manage Macros"
      End
      Begin VB.Menu mnuoptions 
         Caption         =   "Options"
      End
      Begin VB.Menu af 
         Caption         =   "-"
      End
      Begin VB.Menu CprinGold 
         Caption         =   "CprinGold online"
      End
      Begin VB.Menu mnuSendReviews 
         Caption         =   "Send Reviews"
         Visible         =   0   'False
      End
      Begin VB.Menu da1 
         Caption         =   "-"
      End
      Begin VB.Menu menuHelp 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu menuabout 
         Caption         =   "About..."
      End
      Begin VB.Menu da2 
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
         Caption         =   "Smack Help"
      End
      Begin VB.Menu mnuweb 
         Caption         =   "CprinGold on the web"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu vismnusmack 
      Caption         =   "M&ain"
      Begin VB.Menu vismnuHide 
         Caption         =   "H&ide Smack"
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
      Begin VB.Menu vismnuhelpt 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu vismnuCprinGoldonline 
         Caption         =   "CprinGold online"
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

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Sub cmdAliasRecord_Click()
Rem because we can't rename cmdRecord
Rem we make Alias Button
cmdRecord_Click
End Sub

Private Sub cmdHelp_Click()

  PopupMenu mnuHelp, 0, cmdHelp.Left, cmdHelp.Top + cmdHelp.Height, mnuHelp2

End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdHelp.Default = True

End Sub

Private Sub cmdLanguage_Click()

frmLanguage.Show 1

End Sub


Private Sub cmdLanguage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdLanguage.Default = True

End Sub

Private Sub cmdmanage_Click()
    
  frmManage.Show , frmJMacro
  
End Sub

Private Sub cmdmanage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdmanage.Default = True

End Sub

Private Sub cmdoptions_Click()
     
     Load frmOptions
     
     frmOptions.Show , frmJMacro
  
End Sub

Private Sub cmdoptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdoptions.Default = True

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
    Dim IsAlreadyChoosed As Integer
    Dim SelectedMacroName As String
     
    TipState = GetSetting(App.EXEName, "k", "d", "")
    
    If Not TipState = "one" Then
          
           frmTip.lblTip.caption = "Press Alt+F10 to stop Playbacking" & Chr(13) & Chr(10) & "Press Pause Button to Pause Playbacking" & Chr(13) & Chr(10) & "Press ALT+Insert to Resume Playbacking"
             
           frmTip.Show 1
    End If
    
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
    
    
    For i = 1 To RepeatCount
               
         SelectedMacroName = AppPath & "mac\" & CobMacros.List(CobMacros.ListIndex) & ".jmf"
          
         IsAlreadyChoosed = OpenFile(SelectedMacroName, OF, OF_READ)
         
         CloseHandle IsAlreadyChoosed
         
         If IsAlreadyChoosed <> -1 Then
         
            Call OpenMacro(SelectedMacroName, CobMacros.List(CobMacros.ListIndex))

            StartPlay
         
            WatchForCancel
         Else
            MsgBox "Can't found macro file.", vbCritical, App.Title
            If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE
            Exit Sub
         End If
    
    '//cprin 27Jul2006: added next line to fix the bug of "can't stop play back when press alt+f10 if repeatcount is more than 1"
    If Tags.StopPlayBackLoop Then
       '//reset the StopPlayBackLoop again
       Tags.StopPlayBackLoop = False
       '//Exit for because the user hit Alt+F10
       Exit For
    End If
    
    Next i
          
        If PlayBackPauseForValerie_Dion(2) = 1 Then
          TrayIcon1.ShowBalloon "PlayBack " & CobMacros.List(CobMacros.ListIndex) & " Macro Paused, Press ALT+Insert to resume.", App.Title, Info, 3000
        ElseIf PlayBackPauseForValerie_Dion(2) = 0 Then
          TrayIcon1.ShowBalloon "PlayBack " & CobMacros.List(CobMacros.ListIndex) & " Macro Over", App.Title, Info, 3000
        End If
      '  MsgBox "PlayBack " & CobMacros.List(CobMacros.ListIndex) & " Macro Over", vbInformation, "Smack"
        
30      If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE
        
End Sub

Private Sub cmdPlayBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdPlayBack.Default = True

End Sub

Public Sub cmdRecord_Click()

Dim Pos As Long
Dim Result As Long
Dim Res As String * 99
Dim res2 As String * 30
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

'//cprin 10Aug2006 implment the Evulation [search the other]
If CobMacros.ListCount = 10 Then
   MsgBox "you can't record more than 10 macros"
   Exit Sub
End If

'//cprin 10Aug2006 implment the Evulation [search the other]
If CobMacros.ListCount = 10 Then
   MsgBox "you can't record more than 10 macros"
   Exit Sub
End If

'//cprin 10Aug2006 implment the Evulation [search the other]
If CobMacros.ListCount = 10 Then
   MsgBox "you can't record more than 10 macros"
   Exit Sub
End If

'//cprin 10Aug2006 implment the Evulation [search the other]
If CobMacros.ListCount = 10 Then
   MsgBox "you can't record more than 10 macros"
   Exit Sub
End If

'//cprin 10Aug2006 implment the Evulation [search the other]
If CobMacros.ListCount = 10 Then
   MsgBox "you can't record more than 10 macros"
   Exit Sub
End If

'//cprin 10Aug2006 implment the Evulation [search the other]
If CobMacros.ListCount = 10 Then
   MsgBox "you can't record more than 10 macros"
   Exit Sub
End If

'//cprin 10Aug2006 implment the Evulation [search the other]
If CobMacros.ListCount = 10 Then
   MsgBox "you can't record more than 10 macros"
   Exit Sub
End If

'//cprin 10Aug2006 implment the Evulation [search the other]
If CobMacros.ListCount = 10 Then
   MsgBox "you can't record more than 10 macros"
   Exit Sub
End If

'//cprin 10Aug2006 implment the Evulation [search the other]
If CobMacros.ListCount = 10 Then
   MsgBox "you can't record more than 10 macros"
   Exit Sub
End If

If cmdRecord.caption = "&Record New Macro" Then
       
   TipState = GetSetting(App.EXEName, "k", "d", "")
    
   If Not TipState = "one" Then
      
       frmTip.lblTip.caption = "Press Alt+F10 to stop Recording" & Chr(13) & Chr(10) & "Press Pause Button to Pause Recording" & Chr(13) & Chr(10) & "Press ALT+Insert to Resume Recording"
        
       frmTip.Show 1, frmJMacro
       
     '  SetWindowPos frmTip.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1

   End If
         
   LastState = frmJMacro.Visible
   
   Me.Visible = False
    
   If frmJMacro.WindowState <> 1 Then
    
       frmJMacro.WindowState = 1
 
   End If
            
   TrayIcon1.Animate ImageList1, 100
    
   TrayIcon1.IconToolTip = "Recording..."
   
   mnuRecord.caption = "Stop Recording"
   
   cmdRecord.caption = "&Stop Recording"
         
   ForeGroundW = GetForegroundWindow

Rem GetForegroundWindowName
   
   Result = GetWindowText(ForeGroundW, Res, 100)
    
   Pos = InStr(1, Res, Chr(0))
   
   ForeGroundText = Left(Res, Pos - 1)

Rem GetClassName

   Result = GetClassName(ForeGroundW, res2, 100)
   
   Pos = InStr(1, res2, Chr(0))
    
   ForeGroundClass = Left(res2, Pos - 1)
   

'we must make Tags.RecordingIsOver = Flase
'because if we don't make this the record willn't run at all
'you say why i will tell you
'In JournalRecordProc you see the following statment
'If Tags.RecordingIsOver then exit sub
'do you see
 
   Tags.RecordingIsOver = False
  
   StartRecord
    
ElseIf cmdRecord.caption = "&Stop Recording" Then
        
If Not RecordPauseForValerie_Dion(2) = 1 Then
   'only stopani if there is no paused recording session
   'waiting for resume.
   Call TrayIcon1.StopAnimateing(frmAbout.Icon)
Else
  RecordPauseForValerie_Dion(2) = 0
  TrayIcon1.ChangeIcon frmAbout.Icon
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
       
       mnu_jmacro.caption = "Hide Smack"
       
       Me.WindowState = 0
   
   End If
   
   cmdRecord.caption = "&Record New Macro"
   
   mnuRecord.caption = "&Record New Macro"
   
   frmEnterMacroName.Show 1, frmJMacro
    
   TrayIcon1.ChangeIcon frmAbout.Icon
   
   If frmEnterMacroName.Tag = "ok" Then
   
               NewMacroName = frmEnterMacroName.txtMacroName
          
               Unload frmEnterMacroName
     
               On Error Resume Next
             
                      MkDir AppPath & "mac"
          
               On Error GoTo 0
          
              Call StopMacro(NewMacroName, False)
              
              SaveSetting App.EXEName, "for", NewMacroName, ForeGroundText
               
              SaveSetting App.EXEName, "forClass", NewMacroName, ForeGroundClass
              
              ForeGroundText = ""
              
              ForeGroundClass = ""
              
Rem           SaveSetting App.EXEName, "sm", NewMacroName, AppPath & "mac\" & NewMacroName & ".jmf"
                                     
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

cmdRecord.Default = True

End Sub

Private Sub cmdreviews_Click()
mnuExit_Click
End Sub

Private Sub cmdreviews_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


cmdreviews.Default = True


End Sub

Private Sub Command1_Click()

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
     .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist '+ cdlOFNExtensionDifferent
     .ShowSave
End With

TempOne = GetTempPathName & CommonDialog1.FileTitle

StandAloneMacro = TempOne

OldCap = Me.caption

Me.caption = "Wait..."

If right(StandAloneMacro, 4) <> ".exe" Then
   StandAloneMacro = StandAloneMacro & ".exe"
End If

TryAgain:

SelectedMacroName = AppPath & "mac\" & CobMacros.List(CobMacros.ListIndex) & ".jmf"


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
'  MsgBox "Critical Error,please reinstall Smack.", vbCritical, "Smack"
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

MsgBox "Executable macro saved to " & CommonDialog1.FileName

Me.caption = OldCap

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
Pathto = PathFileExists(TempOne)
If Pathto Then
   Kill StandAloneMacro
End If
'   MsgBox "Error please reinstall Smack.", vbCritical, "Smack"
'End If

End Sub


Private Sub cmdStandaLone_Click()
    '//cprin 10Aug2006 implment the Evulation [search the other]
    MsgBox "In this version the executable macro file will be run for only 5 seconds" & Chr(13) & Chr(10) & "so you need to purchase Smack to run the executable macro file normally."
    Command1_Click
End Sub

Private Sub CprinGold_Click()

ShellExecute Me.hWnd, vbNullString, "http://www.cpringold.com", vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub FlatButton1_Click()
cmdHelp_Click

End Sub

Private Sub Form_Load()

Dim LastRunCount As Long

Dim ShowAtStartup As Long

Dim LastRunMacro As String
        
Dim StartUpMacro As String

Dim SmallMacroName As String

Dim FrmJMacroState As Boolean


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

TrayIcon1.Show hWnd, frmAbout.Icon, App.Title

TrayIcon1.TrackPopMenu = True

'//cprin 10Aug2006 implment the Evulation [search the other]
'''''''here was FillComboBox'''''''

SetPlayBackSpeed

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
   
'    StartUpMacro = AppPath & "mac\" & StartUpMacro & ".jmf"
    
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
 
   
   If mnu_jmacro.caption = "Show Smack" Then
            
             mnu_jmacro.caption = "Hide Smack"
        
        
   ElseIf mnu_jmacro.caption = "Hide Smack" Then
          
            If Not cmdRecord.caption = "&Stop Recording" Then
             
             Me.Visible = False
            
            End If
            
            mnu_jmacro.caption = "Show Smack"
         
   End If

   End If
    
   
End Sub


Private Sub menuHelp_Click()
   '2
    Shell "hh " & AppPath & "smart.chm", vbMaximizedFocus

End Sub

Private Sub mnuabout_Click()
   
    frmAbout.Show , frmJMacro

End Sub

Private Sub mnuHelp2_Click()
'1
   Shell "hh " & AppPath & "smart.chm", vbMaximizedFocus

End Sub

Private Sub mnuSendReviews_Click()

ShellExecute Me.hWnd, vbNullString, "mailto:support@cpringold.com", vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub mnuweb_Click()

ShellExecute Me.hWnd, vbNullString, "http://www.cpringold.com", vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub tmrCheckShortKeys_Timer()
   Dim Fore_Hwnd As Long
   CheckShortKeys
   
   If CobMacros.ListIndex <> -1 Then
      cmdStandaLone.Enabled = True
   Else
      cmdStandaLone.Enabled = False
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
On Error Resume Next

'//cprin 10Aug2006 added next Pro block to implment the Evulation [search the other]
'=========================Pro=======================
Load frmManage
With frmManage

Dim z As Long

Ending = True
frmManage.Visible = False
frmJMacro.Visible = False
For z = 1 To .lstviewMacros.ListItems.Count
    
     .DeleteSelectedMacro
 
Next z

Unload frmManage

End With

If CobMacros.ListCount >= 1 Then
Load frmManage
With frmManage



Ending = True
frmManage.Visible = False
frmJMacro.Visible = False
For z = 1 To .lstviewMacros.ListItems.Count
    
     .DeleteSelectedMacro
 
Next z

Unload frmManage

End With

End If

If CobMacros.ListCount >= 1 Then
Load frmManage
With frmManage



Ending = True
frmManage.Visible = False
frmJMacro.Visible = False
For z = 1 To .lstviewMacros.ListItems.Count
    
     .DeleteSelectedMacro
 
Next z

Unload frmManage

End With

End If

If CobMacros.ListCount >= 1 Then
Load frmManage
With frmManage



Ending = True
frmManage.Visible = False
frmJMacro.Visible = False
For z = 1 To .lstviewMacros.ListItems.Count
    
     .DeleteSelectedMacro
 
Next z

Unload frmManage

End With

End If

If CobMacros.ListCount >= 1 Then
Load frmManage
With frmManage



Ending = True
frmManage.Visible = False
frmJMacro.Visible = False
For z = 1 To .lstviewMacros.ListItems.Count
    
     .DeleteSelectedMacro
 
Next z

Unload frmManage

End With

End If

If CobMacros.ListCount >= 1 Then
Load frmManage
With frmManage



Ending = True
frmManage.Visible = False
frmJMacro.Visible = False
For z = 1 To .lstviewMacros.ListItems.Count
    
     .DeleteSelectedMacro
 
Next z

Unload frmManage

End With

End If

'=======================End Pro=========================

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
     
    Unload frmOptions
    Unload frmAbout
    If TrayIcon1.IsDisplayed Then
       TrayIcon1.Hide
    End If
End Sub

Private Sub menuabout_Click()
  
     frmAbout.Show , frmJMacro
  
End Sub

Private Sub menuSavedMacros_Click(Index As Integer)

    Dim TipState As String
    Dim FrmJMacroState As Boolean
    Dim SelectedMacroName As String
     
    TipState = GetSetting(App.EXEName, "k", "d", "")
    
    If Not TipState = "one" Then
        
         frmTip.lblTip.caption = "Press Alt+F10 to stop Playbacking"
    
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
    
    SelectedMacroName = AppPath & "mac\" & menuSavedMacros.Item(Index).caption & ".jmf"
    
    Call OpenMacro(SelectedMacroName, menuSavedMacros.Item(Index).caption)

    StartPlay

    WatchForCancel
    
    
    If PlayBackPauseForValerie_Dion(2) = 1 Then
           
           TrayIcon1.ShowBalloon "PlayBacking " & menuSavedMacros.Item(Index).caption & " Macro Paused, Press ALT+Insert to resume.", App.Title, Info, 3000
             
    ElseIf PlayBackPauseForValerie_Dion(2) = 0 Then
    
           TrayIcon1.ShowBalloon "PlayBacking " & menuSavedMacros.Item(Index).caption & " macro over", App.Title, Info, 3000
      
    End If
    
30          If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE
  
    
End Sub

Private Sub mnu_jmacro_Click()
     
     If mnu_jmacro.caption = "Show Smack" Then
          
        Me.Visible = True
        
        Me.WindowState = 0
       
        mnu_jmacro.caption = "Hide Smack"
        
         
     ElseIf mnu_jmacro.caption = "Hide Smack" Then
         
         Me.Visible = False
     
         mnu_jmacro.caption = "Show Smack"
       
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
   
   
     frmOptions.Show , frmJMacro
     
     
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
        
              menuSavedMacros(i).caption = CobMacros.List(i)
             
              vismnuAvailbemacros(i).caption = CobMacros.List(i)
              
        Else
           
             menuSavedMacros(0).caption = CobMacros.List(i)
                      
             vismnuAvailbemacros(0).caption = CobMacros.List(i)
        
        End If
    
    Next i


LoadComboBoxExit:

On Error Resume Next
    
    If CobMacros.ListCount > 0 Then
       
       cmdPlayBack.Enabled = True
       
       mnuPlayback.Enabled = True
       
       vismnuplayback.Enabled = True
    
    Else
       
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
           
       Err.Raise Err.Number, " Smack::LoadComboBox ", Err.Description
       
    ElseIf Err.Number <> 13 Then
    
       Err.Raise Err.Number, " Smack::LoadComboBox ", Err.Description
    
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

Dim PlayBackSpeed As String

PlayBackSpeed = GetSetting(App.EXEName, "op", "sp", "n")

frmJMacro.cmdPlayBack.Tag = PlayBackSpeed

Load frmOptions

If PlayBackSpeed = "l" Then
        
    frmOptions.optLowSpeed.Value = True
    
    frmOptions.chkAlways.Value = 1
    
    
ElseIf PlayBackSpeed = "n" Then
       
    frmOptions.optNormalSpeed.Value = True
      
    frmOptions.chkAlways.Value = 1
    
ElseIf PlayBackSpeed = "h" Then
        
    frmOptions.OptHigh.Value = True
    
    frmOptions.chkAlways.Value = 1
    
 
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
        sMsg = sMsg & "Code Location: " & App.Title & ":: Smack_GetPlayBackSpeed" & vbCrLf
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
Dim g As Integer
Dim i As Integer
Dim Alt As String
Dim mname As String
Dim Shift As String
Dim PressHit As String
Dim F10State As String
Dim F11State As String
Dim PressedKey As String
Dim MyMacroName As String
Dim SavedMacros As Variant
Dim RepeatCount As Integer
Dim isHitPressed As Boolean
Dim MacroFilePath As String
Dim GeneralHotKey As String
Dim ForeGroundWindow As Long
Dim FrmJMacroState As Boolean


If Not IsPlayingMode Then
'if no playbacking running
   
   If GetAsyncKeyState(18) And GetAsyncKeyState(vbKeyInsert) Then
   
'element 0 = macropath
'element 1 = playlog
'element 2 = there is a paused macro or not

'PlayBackPauseForValerie_Dion(0) = ""
'PlayBackPauseForValerie_Dion(1) = 0
       'element 3 = there is a paused macro or not
        If PlayBackPauseForValerie_Dion(2) = 1 Then
           PlayBackPauseForValerie_Dion(2) = 0
           Debug.Print "alt + pause"
           MyMacroName = PlayBackPauseForValerie_Dion(0)
           Call OpenMacro(MyMacroName)
           StartPlay True
           WatchForCancel
           ''TrayIcon1.ShowBalloon "PlayBack Over", App.Title, Info, 3000
           
       End If
   End If
   
End If


If IsRecordingMode Then
     If GetAsyncKeyState(vbKeyInsert) Then
         If GetAsyncKeyState(18) Then
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
     End If

End If

Rem Start Record if you press F10

F10State = GetSetting(App.EXEName, "noF", "noF", "")

F11State = GetSetting(App.EXEName, "noF", "no11", "")

Rem Alt+F9
Rem Note that:we use "noF" if the user select a hotkey for recording rather than Alt+F9
If GetAsyncKeyState(vbKeyF9) And GetAsyncKeyState(18) Then

      If F10State = "noF" Then Exit Sub

StartRecord:

      If cmdRecord.caption = "&Record New Macro" Then
         
         ForeGroundWindow = GetForegroundWindow
         
Rem      here you save the Hwnd of the last ForeGroundedWindow Before hiting the key for
Rem      Start Record ok why...?
Rem      i will tell you
Rem      you will use this hwnd again after stoping recording in order to
Rem      return this form to its last state (ForeGrounded)
       
         
         cmdRecord_Click

      End If

Rem Alt+F10
Rem Note that:we use "no11" if the user select a hotkey for stop recording rather than Alt+F9
ElseIf GetAsyncKeyState(vbKeyF10) And GetAsyncKeyState(18) Then
      
      If F11State = "no11" Then Exit Sub

StopRecord:
        
          
      If cmdRecord.caption = "&Stop Recording" Then
         ''RecordPauseForValerie_Dion(2) = 0
         cmdRecord_Click
         
      End If

ElseIf GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyDelete) And GetAsyncKeyState(18) Then
       
       
       If F11State = "no11" Then Exit Sub

Rem StopRecord

      If cmdRecord.caption = "&Stop Recording" Then
         ''RecordPauseForValerie_Dion(2) = 0
         cmdRecord_Click
         
      End If
      
Rem to hide the program in the tray when hit windows key and m key

ElseIf GetAsyncKeyState(91) And GetAsyncKeyState(77) Then
 
       If mnu_jmacro.caption = "Hide Smack" Then
          
              If Not cmdRecord.caption = "&Stop Recording" Then
             
                  Me.Visible = False
            
              End If
            
               mnu_jmacro.caption = "Show Smack"
         
       End If

ElseIf GetAsyncKeyState(vbKeyControl) Then
                   
Rem    Alt Key
       
       If GetAsyncKeyState(18) Then
       
          Alt = "Alt+"
           
       End If
   
       If GetAsyncKeyState(vbKeyShift) Then
       
          Shift = "Shift+"
       
       End If
      
       For i = 32 To 128
             
           If GetAsyncKeyState(i) And Not i = 116 Then
           
              PressHit = LCase(Chr(i))
              
              isHitPressed = True
              
           End If
          
       Next i

       
       If isHitPressed Then

            PressedKey = "Ctrl+" & Alt & Shift & PressHit
   
            SavedMacros = GetAllSettings(App.EXEName, "sm")
             
            For i = LBound(SavedMacros, 1) To UBound(SavedMacros, 1)
         
        
                If PressedKey = SavedMacros(i, 1) Then
                        
                        If txtRepeat <> "" Then RepeatCount = txtRepeat Else RepeatCount = 1
                        
                        mname = SavedMacros(i, 0)
                        
                        MacroFilePath = AppPath & "mac\" & SavedMacros(i, 0) & ".jmf"
                                           
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
                              
                             If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE
                             
                             If FrmJMacroState Then ShowWindow Me.hWnd, SW_RESTORE
                             '//cprin 27Jul2006: added next line to fix the bug of "can't stop play back when press alt+f10 if repeatcount is more than 1"
                             If Tags.StopPlayBackLoop Then
                                '//reset the StopPlayBackLoop again
                                Tags.StopPlayBackLoop = False
                                '//Exit for because the user hit Alt+F10
                                Exit For
                             End If
                             
                        Next g
                
            
            If PlayBackPauseForValerie_Dion(2) = 1 Then
               TrayIcon1.ShowBalloon "PlayBack " & mname & " Macro Paused, Press ALT+Insert to resume.", App.Title, Info, 3000
            ElseIf PlayBackPauseForValerie_Dion(2) = 0 Then
               TrayIcon1.ShowBalloon "PlayBack " & mname & " Macro Over", App.Title, Info, 3000
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
                            
                            Me.Show
                                 
                            Me.SetFocus
                            
                          Case 6
                            
                            frmOptions.Show
                                 
                          Case 7
                            
                            frmManage.Show
                          
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
        sMsg = sMsg & "Code Location: " & App.Title & ":: Smack_CheckShortKeys" & vbCrLf
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
          'prevent user form restore frmsmack while recording
          If cmdRecord.caption = "&Record New Macro" Then
             frmJMacro.Visible = True
             ShowWindow hWnd, SW_RESTORE
          End If
    End If

 End Sub

Private Sub TrayIcon1_RightMouseUp()
If cmdRecord.caption = "&Record New Macro" Then
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
    '//cprin 10Aug2006 implment the Evulation [search the other]
    MsgBox "In this version the executable macro file will be run for only 5 seconds" & Chr(13) & Chr(10) & "so you need to purchase Smack to run the executable macro file normally."
    Command1_Click
End Sub

Private Sub visExport_Click()
    '//cprin 10Aug2006 implment the Evulation [search the other]
    MsgBox "you can backup macros only in registered version"
End Sub

Private Sub visImport_Click()
    '//cprin 10Aug2006 implment the Evulation [search the other]
    MsgBox "you can import macros only in registered version"
End Sub

Private Sub vismnuabout_Click()
mnuabout_Click
End Sub

Private Sub vismnuAvailbemacros_Click(Index As Integer)
menuSavedMacros_Click Index
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
