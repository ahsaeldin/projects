VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{42E02B6F-DBB2-11DA-AB1C-00E04C770EEB}#1.0#0"; "mrac.ocx"
Begin VB.Form Form1 
   Caption         =   "Vb Macro Recorder"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin MacroRecorderActX.MacroRecorder MacroRecorder1 
      Left            =   6000
      Top             =   360
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin VB.Timer tmrState 
      Interval        =   100
      Left            =   5040
      Top             =   360
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmdStopRecord 
         Caption         =   "&Save as Exe"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "&Record"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "2.Click Save as Exe button to stop recording and save the macro as an exectutable file."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "1. Click Record button to start recording mouse clicks, keystrokes ."
         Height          =   255
         Left            =   120
         TabIndex        =   3
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
'====================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrates how to use MacroRecorder
'               ActiveX control v1.50 to record and replay mouse clicks, keystrokes
'               and bundle them into an executble file [.exe] in order to replay later.
'=====================================================================================
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
    
    'Only if State is recording.
    If Not MacroRecorder1.IsRecord Then
        Exit Sub
    End If

    Me.WindowState = 0

    MacroRecorder1.StopRecord

    With CommonDialog1
        .FileName = "macro1"
        .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
        .ShowSave
    End With


    'Note that
    'Macro Recorder ActiveX Control will add the ".exe" extension
    'to the file path automatically
    'For example suppose that you want to save the macro to c:\
    'you can use one of the following
    'MacroRecorder1.SaveAsExE ("c:\macro.exe")
    'MacroRecorder1.SaveAsExE ("c:\macro")     'Macro Recorder will add the .exe extension to the file
    
     MacroRecorder1.SaveAsExE (CommonDialog1.FileName)
     
     'You can use command line parameters for the executable macro file to change the replay speed
     'For example if you save an executable macro to macro1.exe
     'macro1 /h    for high replay speed.
     'macro1 /n    for normal replay speed "the default one"
     'macro1 /l    for low replay speed.

End Sub

Private Sub MacroRecorder1_RecordStart()
    MsgBox "Press Alt+F10 to Stop Recording."
End Sub

Private Sub MacroRecorder1_SaveComplete(ByVal MacFilePath As String)
    If MacFilePath <> "" Then MsgBox "Macro Saved To " & MacFilePath
End Sub

Private Sub tmrState_Timer()

    'Stop Recording when you Click Alt+F10
    If GetAsyncKeyState(vbKeyF10) And GetAsyncKeyState(18) Then
        cmdStopRecord_Click
    End If

   If MacroRecorder1.IsRecord And Not Me.Caption = "Recording ..." Then
        Me.Caption = "Recording ..."
   ElseIf Not MacroRecorder1.IsReplay And Not MacroRecorder1.IsRecord Then
        If Not Me.Caption = "Vb Macro Recorder" Then
            Me.Caption = "Vb Macro Recorder"
        End If
    End If

End Sub
