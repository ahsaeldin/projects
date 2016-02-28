VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00A3FCCF-E2C6-11DA-AB1C-00E04C770EEB}#1.0#0"; "mrac.ocx"
Begin VB.Form Form1 
   Caption         =   "Vb Macro Recorder"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin MacroRecorderActX.MacroRecorder MacroRecorder1 
      Left            =   5760
      Top             =   240
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin VB.Timer tmrState 
      Interval        =   100
      Left            =   5040
      Top             =   1200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdReplay 
         Caption         =   "&Replay"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdStopRecord 
         Caption         =   "&Stop Record"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   1440
         List            =   "Form1.frx":000D
         TabIndex        =   2
         Text            =   "Replay Speed"
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Record"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "3.Click Replay button to Replay the macro you had saved"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   5775
      End
      Begin VB.Label Label2 
         Caption         =   "2. Click Stop Record button to stop recording and save the macro to the disk."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "1. Click Record button to start recording mouse clicks, keystrokes ."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrates how to use MacroRecorder
'               ActiveX control v1.50 to record and replay mouse clicks, keystrokes
'               and bundle them into a file [macro] in order to replay later.
'=============================================================================
Option Explicit
 
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'==============================================================================
' Method:        cmdRecord_Click
'
' Description:  Start recording a new macro.
'==============================================================================
Private Sub cmdRecord_Click()
    If Me.WindowState <> 1 Then Me.WindowState = 1
    MacroRecorder1.Record
End Sub

'==============================================================================
' Method:        cmdStopRecord_Click
'
' Description:  Stop recording and save the macro to the disk.
'==============================================================================
Private Sub cmdStopRecord_Click()

  '  On Error GoTo quitme:

    'Only if State is recording.
    If Not MacroRecorder1.IsRecord Then
        Exit Sub
    End If

    Me.WindowState = 0

    'Stop recording while get the path from the user
    MacroRecorder1.StopRecord

    With CommonDialog1
        'You can use any extension for the macro file name, or you can use no extension.
        'hence it will be easy to integrate the MacroRecorder ActiveX Control in your
        'Applications by using any file extension you want for the macro file generated by
        'MacroRecorder ActiveX Control.
      
        .FileName = "macro1"
        .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
        .ShowSave
    End With

    'You can use any extension for the macro file name, or you can use no extension.
    'hence it will be easy to integrate the MacroRecorder ActiveX Control in your
    'Applications by using any file extension you want for the macro file generated by
    'MacroRecorder ActiveX Control.
    
    'For example suppose that you want to save the macro to c:\
    'If you want to assign (*.xxx) as an extension for the generated macro file,
    'then all you have to do is
    'MacroRecorder1.Save ("c:\mymacro.xxx")
    'If you want to use no extension then use the following:
    'MacroRecorder1.Save ("c:\mymacro")
    
     MacroRecorder1.Save (CommonDialog1.FileName)

End Sub

'==============================================================================
' Method:        cmdReplay_Click
'
' Description:  Start Replay a macro.
'==============================================================================
Private Sub cmdReplay_Click()

    On Error GoTo quitme:
    If MacroRecorder1.IsReplay Then
        Exit Sub
    End If
    With CommonDialog1
        .CancelError = True
        .ShowOpen
    End With

    'Pass to MacroPath parameter any file you had saved by Save Method,
    'passing invalid macro file has no effect and no Replaying will happen.
    MacroRecorder1.StartReplay (CommonDialog1.FileName)

    Exit Sub
quitme:

End Sub

'==============================================================================
' Method:        Combo1_Click
'
' Description:  Change the Replay speed.
'==============================================================================
Private Sub Combo1_Click()
    'To Replay with high speed set Speed parameter to 0
    'To Replay with Normal speed set Speed parameter to 1
    'To Replay with Low speed set Speed parameter to 2
    MacroRecorder1.SetReplaySpeed Combo1.ListIndex
End Sub

Private Sub MacroRecorder1_RecordStart()
    MsgBox "Press Alt+F10 to Stop Recording."
End Sub

Private Sub MacroRecorder1_ReplayFinish(ByVal MacFilePath As String)
    MsgBox "Replay " & MacFilePath & " Macro Finish"
End Sub

Private Sub MacroRecorder1_ReplayStart(ByVal MacFilePath As String)
    MsgBox "Press Alt+F9 to terminate Replay."
End Sub

Private Sub MacroRecorder1_SaveComplete(ByVal MacFilePath As String)
    If MacFilePath <> "" Then MsgBox "Macro Saved To " & MacFilePath
End Sub

Private Sub tmrState_Timer()

    'Stop Recording when you Click Alt+F10
    If GetAsyncKeyState(vbKeyF10) And GetAsyncKeyState(18) Then
        cmdStopRecord_Click
    End If

    'Stop Replaying when you Click Alt+F9
    If GetAsyncKeyState(vbKeyF9) And GetAsyncKeyState(18) Then
        MacroRecorder1.StopReplay
    End If

    If MacroRecorder1.IsRecord And Not Me.Caption = "Recording ..." Then
        Me.Caption = "Recording ..."
    ElseIf MacroRecorder1.IsReplay And Not Me.Caption = "Replaying..." Then
        Me.Caption = "Replaying..."
    ElseIf Not MacroRecorder1.IsReplay And Not MacroRecorder1.IsRecord Then
        If Not Me.Caption = "Vb Macro Recorder" Then
            Me.Caption = "Vb Macro Recorder"
        End If
    End If

End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 1
End Sub

