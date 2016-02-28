VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\AMRAC.vbp"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   2505
   StartUpPosition =   2  'CenterScreen
   Begin MacroRecorderActX.MacroRecorder MacroRecorder1 
      Left            =   1560
      Top             =   2040
      _extentx        =   926
      _extenty        =   926
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Convert"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1080
      Top             =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Playback"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Set Speed"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Get Speed"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "IsRecord"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Interval        =   8000
      Left            =   600
      Top             =   2040
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Command1_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
    If frmTest.WindowState <> 1 Then
    
        frmTest.WindowState = 1
 
    End If
   
    MacroRecorder1.Record

    Timer2.Enabled = True

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Command2_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    'Because there is no an explicit extension for the file, you can use your own file extension or you can use a file path with no extension.

    If MacroRecorder1.IsRecord Then
       frmTest.WindowState = 0
       MacroRecorder1.StopRecord
       CommonDialog1.ShowSave
       MacroRecorder1.Save CommonDialog1.FileName
    End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Command3_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    'Pass to MacroPath parameter any file you had saved by  SaveMacro Method,
    'passing unvalid macro file has no effect and no playbacking will happen.
        
    CommonDialog1.ShowOpen
    
    If Not MacroRecorder1.IsReplay Then
       MacroRecorder1.StartReplay CommonDialog1.FileName
    End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Command4_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    MacroRecorder1.SetReplaySpeed LowSpeed

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Command5_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    MsgBox MacroRecorder1.GetReplaySpeed

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Command7_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    MsgBox MacroRecorder1.IsRecord

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Command8_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

   CommonDialog1.ShowOpen
   
   MacroRecorder1.ConvertToExE CommonDialog1.FileName, "CommonDialog1.FileTitle"

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub MacroRecorder1_ConvertComplete(ByVal MacFilePath As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    MsgBox "Macro Converted to " & MacFilePath

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub MacroRecorder1_RecordFinish()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    MsgBox "Record Finish "

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub MacroRecorder1_RecordStart()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    MsgBox "Record Start "

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub MacroRecorder1_ReplayFinish(ByVal MacFilePath As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    MsgBox "pal finsih " & MacFilePath

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub MacroRecorder1_ReplayStart(ByVal MacFilePath As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    MsgBox "Playback statrt " & MacFilePath

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub MacroRecorder1_SaveComplete(ByVal MacFilePath As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    MsgBox "Macro saved to " & MacFilePath

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Timer1_Timer()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


    If GetAsyncKeyState(vbKeyF10) And GetAsyncKeyState(18) Then
       
       Command2_Click
       
       If MacroRecorder1.IsReplay Then
          
          MacroRecorder1.StopReplay
       
       End If
    
    End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Timer2_Timer()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    MacroRecorder1.StopReplay

    Timer2.Enabled = False

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Timer3_Timer()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Debug.Print MacroRecorder1.IsRecord, MacroRecorder1.IsReplay

End Sub
