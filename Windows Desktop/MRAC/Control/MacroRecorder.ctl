VERSION 5.00
Begin VB.UserControl MacroRecorder 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   525
   ScaleWidth      =   525
   ToolboxBitmap   =   "MacroRecorder.ctx":0000
   Begin VB.PictureBox Picture1 
      Height          =   530
      Left            =   0
      Picture         =   "MacroRecorder.ctx":0312
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   0
      Width           =   530
   End
End
Attribute VB_Name = "MacroRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Dim CopyFromTempFolder As Boolean
 
    Public Enum ReplaySpeed
        HighSpeed = 0
        NormalSpeed = 1
        LowSpeed = 2
    End Enum

    Public Event RecordStart()
    Public Event SaveComplete(ByVal MacFilePath As String)
    Public Event ConvertComplete(ByVal MacFilePath As String)
    Public Event ReplayFinish(ByVal MacFilePath As String)
    Public Event ReplayStart(ByVal MacFilePath As String)
    Public Event RecordFinish()

    Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub Record()
Attribute Record.VB_Description = "Start recording a new macro."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

        If IsPlayingMode Then
    
        Rem It is not realy a click for cmdRecord Button
        Rem it is the last message Recored in the last Macro
        Rem when he just attempt to stop Recording his Macro
        Rem so we will Exit this sub
       
            Debug.Print "Exit Sub Man"
        
            Exit Sub
    
        End If
    
        RaiseEvent RecordStart
    
        IsRecording = True
    
        StartRecord

    End Sub
    
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub Save(ByVal Path As String)
Attribute Save.VB_Description = "Save the macro to a file."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

        If IsPlayingMode Then

            Rem It is not realy a click for cmdRecord Button
            Rem it is the last message Recored in the last Macro
            Rem when he just attempt to stop Recording his Macro
            Rem so we will Exit this sub
            Exit Sub

        End If

        IsRecording = False

        'That's means, there is a recording session and the user call Save before calling StopMacro
        If bRecording Then
           StopMacro Path
           RaiseEvent SaveComplete(Path)
           Exit Sub
        End If
      
        'That's means, there is a recording session and the user call stopmacro before call calling stopmacro
        If CopyFromTempFolder Then
           On Error Resume Next
           FileCopy GetTempPathName & "4564", Path
           On Error GoTo 0
           RaiseEvent SaveComplete(Path)
           CopyFromTempFolder = False
        Else
          'MsgBox "Recording didn't start", vbCritical, "Macro" 'cprin
        End If

    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub SaveAsExE(ByVal Path As String)
Attribute SaveAsExE.VB_Description = "Save the macro as an executable file."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
  
        Dim IsDeleted As Long
        
        If IsPlayingMode Then
        
            Rem It is not realy a click for cmdRecord Button
            Rem it is the last message Recored in the last Macro
            Rem when he just attempt to stop Recording his Macro
            Rem so we will Exit this sub
            Exit Sub
            
        End If
        
        IsRecording = False
        
        'That's means, there is a recording session and the user call Save before calling StopMacro
        If bRecording Then
           
           On Error Resume Next
           IsDeleted = DeleteFile(GetTempPathName & "4563")
           On Error GoTo 0
           StopMacro GetTempPathName & "4563"
           ExecutableMacro GetTempPathName & "4563", Path
           RaiseEvent SaveComplete(Path)
           Exit Sub
         
         End If
              
         'That's means, there is a recording session and the user call stopmacro before call calling stopmacro
         If CopyFromTempFolder Then
            ExecutableMacro GetTempPathName & "4564", Path
            RaiseEvent SaveComplete(Path)
            CopyFromTempFolder = False
         Else
            'MsgBox "Recording didn't start", vbCritical, "Macro" 'cprin
         End If
    
    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub ConvertToExE(ByVal Source As String, ByVal Destination As String)
Attribute ConvertToExE.VB_Description = "Converts a macro file to an executable macro file."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

        Dim IsDeleted As Long

        If IsPlayingMode Then
           Rem It is not realy a click for cmdRecord Button
           Rem it is the last message Recored in the last Macro
           Rem when he just attempt to stop Recording his Macro
           Rem so we will Exit this sub
           Exit Sub
        End If
    
        ExecutableMacro Source, Destination
    
        RaiseEvent ConvertComplete(Destination)
    
    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub StopRecord()
Attribute StopRecord.VB_Description = "Stop recording keyboard and mouse events."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
        Dim IsDeleted As Long
        
        If IsPlayingMode Then
        
            Rem It is not realy a click for cmdRecord Button
            Rem it is the last message Recored in the last Macro
            Rem when he just attempt to stop Recording his Macro
            Rem so we will Exit this sub
            Exit Sub
        
        End If
          
        IsRecording = False
           
        If bRecording Then
          
            IsDeleted = DeleteFile(GetTempPathName & "4564")
            StopMacro GetTempPathName & "4564"
            CopyFromTempFolder = True
            RaiseEvent RecordFinish

        Else
            
            'MsgBox "Recording didn't start", vbCritical, "Macro" 'cprin
        
        End If
      
    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub StartReplay(ByVal Path As String)
Attribute StartReplay.VB_Description = " Start replay a  macro."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
  
        Dim res As Boolean
          
        If bRecording Then
           'MsgBox "Can't Playback while Recording.", vbCritical, "Macro"  'cprin
           Exit Sub
        End If
        
        If IsPlayingMode Then
           'MsgBox "you already playback.", vbCritical, "Macro"  'cprin
           Exit Sub
        End If
        
        RaiseEvent ReplayStart(Path)
            
        res = OpenMacro(Path)
            
        If Not res Then Exit Sub
            
        StartPlay
                 
        WatchForCancel
                
        RaiseEvent ReplayFinish(Path)
        
    End Sub
    
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub StopReplay()
Attribute StopReplay.VB_Description = "Teminate a repaly session."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
        If IsRecord Then Exit Sub
           
        If Not IsReplay Then Exit Sub
        
        StopMacro
           
        'THERE IS NO NEED HERE TO CALL
        'RaiseEvent RecordFinish
        'why
        'because it called automatically after WatchForCancel
        
    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub UserControl_Resize()
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

        UserControl.Height = 530
        UserControl.Width = 530

    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function GetReplaySpeed() As Integer
Attribute GetReplaySpeed.VB_Description = "Returns the current Replay speed."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
        Select Case speedo
               Case "h"
                   GetReplaySpeed = HighSpeed
               Case "l"
                   GetReplaySpeed = LowSpeed
               Case Else
                   GetReplaySpeed = NormalSpeed
        End Select
    
    End Function

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function SetReplaySpeed(ByVal Speed As ReplaySpeed)
Attribute SetReplaySpeed.VB_Description = "Set the repaly speed."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
        Select Case Speed
               Case 0
                   speedo = "h"
               Case 1
                   speedo = "n"
               Case 2
                   speedo = "l"
        End Select
    
    End Function

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Property Get IsRecord() As Boolean
Attribute IsRecord.VB_Description = "Returns true if the MacroRecorder control is in Recording session."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
        If Not IsRecording And bRecording Then
            'that's means that the user check the IsRecord method in between
            'calling StopRecord Method and Save Method for ex:
            'MacroRecorder1.StopRecord
            'MacroRecorder1.IsRecording
            IsRecord = IsRecording
            Exit Property
        End If
        
        IsRecord = bRecording
    
    End Property

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Property Get IsReplay() As Boolean
Attribute IsReplay.VB_Description = "Returns true if the MacroRecorder control is in Replay session."
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
        IsReplay = IsPlayingMode
    
    End Property

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub UserControl_Terminate()
     '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
        If bRecording Then
           StopMacro ""
        End If
    
    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub ExecutableMacro(ByVal MacPath As String, SaveTo As String)
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
        Dim marker As Variant
        Dim TempOne As String
        Dim MacroData As String
        Dim StandAloneData As String
        Dim StandAloneMacro As String
        Dim cprsosonono3345554 As Long
        Dim SelectedMacroName As String
        Dim OldCap As String
        
        On Error GoTo quitme:
        
        TempOne = GetTempPathName & "s3"
        
        StandAloneMacro = TempOne
        
        If Right(SaveTo, 4) <> ".exe" Then
           SaveTo = SaveTo & ".exe"
        End If
        
TryAgain:
        
        SelectedMacroName = MacPath
        
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
        
        FileCopy TempOne, SaveTo
        
        On Error Resume Next
        
        Kill TempOne
        
        Exit Sub
quitme:
        
        'MsgBox Err.Description, vbCritical, App.Title
             
        'If Err.Number = cdlDialogFailure Then
        '   StandAloneMacro = AppPath & CobMacros.List(CobMacros.ListIndex) & ".exe"
        '   MsgBox "Save as Dialog Failure, Standalone file saved automatically to " & StandAloneMacro
        '   GoTo TryAgain:
        'ElseIf Err.Number = 326 Then
        '   Close #1
            On Error Resume Next
            Dim Pathto As Boolean
            'Pathto = PathFileExists(TempOne)
            If Pathto Then
               Kill StandAloneMacro
            End If
       'End If
    
    End Sub
