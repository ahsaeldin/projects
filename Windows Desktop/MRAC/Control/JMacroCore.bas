Attribute VB_Name = "JMacroCore"
Option Explicit

    'How to record Macro:
    '1.call startrecord
    '2.call stopmacro

    'How to playback Macro:
    '1.openmacro
    '2.call startplay
    '3.call watchforcanvel

    'Hook Constants
    Private Const WH_KEYBOARD = 2
    Private Const WH_JOURNALRECORD = 0
    Private Const WH_JOURNALPLAYBACK = 1

    'Hook Call Constants
    Private Const HC_SKIP = 2
    Private Const HC_ACTION = 0
    Private Const HC_GETNEXT = 1
    Private Const VK_MENU = &H12
    Private Const PM_REMOVE = &H1
    Private Const VK_CONTROL = &H11
    Private Const HC_SYSMODALON = 4
    Private Const HC_SYSMODALOFF = 5
    Private Const GWL_HINSTANCE = (-6)
    Private Const KEYEVENTF_KEYUP = &H2
    Private Const WM_CANCELJOURNAL = &H4B

    Private Type POINTAPI
            X As Long
            Y As Long
    End Type

    Private Type MSG
        hWnd As Long
        message As Long
        wParam As Long
        lParam As Long
        time As Long
        pt As POINTAPI
    End Type

    Private Type EVENTMSG
            message As Long
            paramL As Long
            paramH As Long
            time As Long
            hWnd As Long
    End Type

    Private Declare Function WaitMessage Lib "user32" () As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
    Private Declare Sub CopyMemoryT2H Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As Long, Source As EVENTMSG, ByVal Length As Long)
    Private Declare Sub CopyMemoryH2T Lib "kernel32" Alias "RtlMoveMemory" (Dest As EVENTMSG, ByVal Source As Long, ByVal Length As Long)
    Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


    Private MacroName
    Private MacroPath
    Private recOK As Long
    Private hHook As Long
    Private hPlay As Long
    Private sEXT As String
    Private pos As POINTAPI
    Private PlayLog As Long
    Private EventLog As Long
    Private bDelay As Boolean
    Private canPlay As Boolean
    Private lStartTime As Long
    Private ParenthWnd As Long
    Private lbstarttime As Long
    Private bPlaying As Boolean
    Private bKeyStop As Boolean
    Private sFileName As String
    Private iFileNumber As String
    Public bRecording As Boolean
    Private EventArr() As EVENTMSG
    Public speedo As String
    Public IsRecording As Boolean

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function IsPlayingMode() As Boolean
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
          
           IsPlayingMode = bPlaying

    End Function
    
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function OpenMacro(FileName As String) As Boolean
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        
        On Error GoTo fin

        Dim bHotKey
        Dim sfilter As String
        Dim FileLength As Long

        iFileNumber = FreeFile()
        
        'If Right(FileName, 4) <> ".maf" Then
            'MsgBox "Invalid Macro File"
            'OpenMacro = False
            'Exit Function
        'End If
        
        If bHotKey = False Then
           sFileName = FileName
        End If

        If sFileName <> vbNullChar Then
    
           ReDim EventArr(0)
    
           Open sFileName For Random Access Read As iFileNumber Len = Len(EventArr(0))
        
               FileLength = 1
            
               Do
                    ReDim Preserve EventArr(FileLength - 1)
                
                    Get #iFileNumber, FileLength, EventArr(FileLength - 1)
                
                    FileLength = FileLength + 1
            
               Loop Until EOF(iFileNumber)
            
               EventLog = FileLength - 2
    
           Close iFileNumber

        Else
   

        End If

        OpenMacro = True

        Exit Function

fin:

        Select Case Err.Number
            Case Else
                'MsgBox Err.Number & " OpenMacro " & Err.Description 'cprin
         End Select

    End Function
    
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub SaveMacro(Optional MacroName As String)
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
        On Error GoTo fin

        Dim projectpath
        Dim iIndex As Long
        Dim sfilter As String
        
        sEXT = ".maf"

        iFileNumber = FreeFile()

        DoEvents
        
        If MacroName = "" Then
           'MacroName = InputBox("Enter Macro Name", "Reco")
           Exit Sub
        End If
        
        ''If Right(MacroName, 4) <> ".maf" Then
        ''   MsgBox "Macro File extension must be .maf"
        ''   Exit Sub
        ''End If
        
        sFileName = MacroName 'ShowSave(frmmangmacro.hwnd, True, sfilter)
        
        If sFileName <> vbNullChar Then
            
            If Mid(Right(sFileName, 4), 1, 1) <> "." Then
                
                ''''''sFileName = sFileName & sEXT
                
            End If
            
            If Len(Dir(sFileName)) > 0 Then
                
                Kill sFileName
            
            End If
            
            MacroPath = sFileName
            
            Open sFileName For Random Access Write As iFileNumber Len = Len(EventArr(EventLog))
                        
                For iIndex = 1 To EventLog '- 100
                    
                    Put iFileNumber, iIndex, EventArr(iIndex - 1)
                
                Next iIndex
            
            Close iFileNumber
           
        End If
        
        Exit Sub
        
fin:
        Select Case Err.Number
            
            Case Else
                
                ''''MsgBox Err.Number & " SaveMacro " & Err.Description 'cprin
                
        End Select

End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Function JournalRecordProc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        
        On Error GoTo fin
        Dim m As MSG
        Dim Result As Long
        Dim sMsg As String
        
        ReDim Preserve EventArr(EventLog)
        
        recOK = 1
        
        Result = 0
        
        If iCode < 0 Then
            
            Result = CallNextHookEx(hHook, iCode, wParam, lParam)
        
        ElseIf iCode = HC_SYSMODALON Then
            
            recOK = 0
        
        ElseIf iCode = HC_SYSMODALOFF Then
            
            recOK = 1
        
        ElseIf ((recOK > 0) And (iCode = HC_ACTION)) Then
            
            If CheckStopKey = True Then
                
                StopMacro , True
                
                bKeyStop = True
                
                JournalRecordProc = CallNextHookEx(hHook, iCode, wParam, lParam)
                
                Exit Function
            
            End If
            
            If IsRecording = False Then
               GoTo Finish:
            End If
         
            CopyMemoryH2T EventArr(EventLog), lParam, Len(EventArr(EventLog))
            
            EventArr(EventLog).time = EventArr(EventLog).time - lbstarttime
            
            EventLog = EventLog + 1
            
            Result = CallNextHookEx(hHook, iCode, wParam, lParam)
        
        End If
        
Finish:
        JournalRecordProc = Result
        
        Exit Function
        
fin:
        
        Select Case Err.Number
            
            Case Else
                
                'MsgBox Err.Number & " JournalRecordProc " & Err.Description  'cprin
        
        End Select
        
    End Function

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Function JournalPlaybackProc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

        On Error GoTo fin
        
        Dim sMsg As String
        Dim lPause As Long
        Dim Result As Long
        Dim evtMsg As EVENTMSG
        Dim PlayBackSpeed As String
        
        canPlay = True
        
        Result = 0
        
        If iCode < 0 Then
            
            Result = CallNextHookEx(hPlay, iCode, wParam, lParam)
        
        ElseIf iCode = HC_SYSMODALON Then
            
            canPlay = False
        
        ElseIf iCode = HC_SYSMODALOFF Then
            
            canPlay = True
        
        ElseIf ((canPlay = True) And (iCode = HC_GETNEXT)) Then
            'This code controls the timing between system messages
            If bDelay Then
                
                bDelay = False
                         
                'if you wish a normal speed run don't run line 1
                'if you wish a lighting speed make result = 1
                'if you wish a slow speed don't run line 7
                'as you see we don't need to use Line1 at-all
        
                 
                 
                Rem      we don't need line one at-all, i put it only for the unexpected Modifies
1               Rem   EventArr(PlayLog).time = EventArr(PlayLog).time + lStartTime
Normal_Speed:
        
2               If PlayLog > 0 And PlayLog < EventLog Then
                    
3                  lPause = (EventArr(PlayLog).time - lStartTime)
                
4               Else
                    
5                  lPause = 0
                
6               End If
                
                PlayBackSpeed = speedo
                  
                If PlayBackSpeed <> "l" Then
        
7                  lStartTime = EventArr(PlayLog).time
                
                End If
                
                If PlayBackSpeed = "h" Then
                 
                   Result = 0.1
                 
                Else
                    
                   Result = lPause
                    
                End If
            
            End If
            
             ' ForeGroundIt
              CopyMemoryT2H lParam, EventArr(PlayLog), Len(EventArr(PlayLog))
             'old = EventArr(PlayLog).time
        
        ElseIf ((canPlay = True) And (iCode = HC_SKIP)) Then
            
            If PlayLog >= EventLog Or CheckStopKey = True Then
                
                StopMacro
                
                bPlaying = False
                
                JournalPlaybackProc = Result
                
                Exit Function
            
            End If
            
            bDelay = True
            
            Result = CallNextHookEx(hPlay, iCode, wParam, lParam)
            
            PlayLog = PlayLog + 1
        
        End If
        
        JournalPlaybackProc = Result
        Exit Function
fin:
        Select Case Err.Number
            Case Else
            
                StopMacro
                'MsgBox Err.Number & " JournalPlaybackProc " & Err.Description
            
        End Select
        
    End Function
        
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub StartRecord()
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
      
        On Error GoTo fin
        
        EventLog = 0
        
        lbstarttime = GetTickCount
        
        hHook = SetWindowsHookEx(WH_JOURNALRECORD, AddressOf JournalRecordProc, App.hInstance, 0)
        
        If hHook <> 0 Then bRecording = True
        
        Exit Sub
        
fin:
        Select Case Err.Number
            
            Case Else
                
                 'MsgBox Err.Number & " StartRecord " & Err.Description 'cprin
         
        End Select
    
    End Sub
    
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub StartPlay()
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 
        PlayLog = 0
        
        lStartTime = EventArr(0).time
        
        hPlay = SetWindowsHookEx(WH_JOURNALPLAYBACK, AddressOf JournalPlaybackProc, App.hInstance, 0)
        
        bKeyStop = False
        
        If hPlay <> 0 Then
            
            bPlaying = True
        
        Else
            
            bPlaying = False
        
        End If
    
    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub StopMacro(Optional MacroName As String, Optional cancel As Boolean)
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
      
        On Error GoTo fin
        Dim retval As Long
        
        If bPlaying Then
            
            retval = UnhookWindowsHookEx(hPlay)
            
            DoEvents
            
            If retval Then
                
                bPlaying = False
            
            End If
        
        ElseIf bRecording Then
            
            retval = UnhookWindowsHookEx(hHook)
            
            DoEvents
            
            If retval Then
               
               bRecording = False
                
               If Not cancel Then
                 
                 Call SaveMacro(MacroName)
               
               End If
            
            End If
        
        End If
        
        PlayLog = 0
        
        EventLog = 0
        
        Erase EventArr
        
        Exit Sub
        
fin:
        
        Select Case Err.Number
            
            Case Else
                
               ''' MsgBox Err.Number & " StopMacro " & Err.Description 'cprin
        
        End Select
        
    End Sub
    
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Function CheckStopKey() As Boolean
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 
        On Error GoTo fin
        
        If GetAsyncKeyState(vbKeyF12) Then ' And GetAsyncKeyState(VK_MENU) And GetAsyncKeyState(Asc("s")) Then
            
            CheckStopKey = True
            
            bKeyStop = True
            
            Call keybd_event(VK_CONTROL, &H45, KEYEVENTF_KEYUP, 0)
            
            Call keybd_event(VK_MENU, &H45, KEYEVENTF_KEYUP, 0)
        
        Else
            
            CheckStopKey = False
            
            bKeyStop = False
        
        End If
        
        Exit Function
        
fin:
        Select Case Err.Number
            
            Case Else
                
                 'MsgBox Err.Number & " CheckStopKey " & Err.Description 'cprin
        
        End Select
        
    End Function
        
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub WatchForCancel()
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 
        On Error GoTo fin
        
        Dim message As MSG
        
        Do Until bKeyStop = True Or bPlaying = False
           
           WaitMessage
           
            'check if it's a HOTKEY-message
            
           If PeekMessage(message, 0, WM_CANCELJOURNAL, WM_CANCELJOURNAL, PM_REMOVE) Then
                
              bKeyStop = True
                
              bPlaying = False
                
              StopMacro
            
            End If
            
            'let the operating system process other events
            
            DoEvents
        
        Loop
        
        Exit Sub
        
fin:
        
        Select Case Err.Number
            
            Case Else
                
                 'MsgBox Err.Number & " WatchForCancel " & Err.Description 'cprin
        
        End Select
        
    
    End Sub
        
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function AppPath()
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        
        'use the function to get the path of the parameter filename
        If Right(App.Path, 1) <> "\" Then AppPath = App.Path + "\" Else AppPath = App.Path
           
    End Function

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Property Get MacGenNameAndVersion() As Variant
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
       
        MacGenNameAndVersion = App.Title & " " & App.Major & "." & App.Minor & App.Revision
    
    End Property
    
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function GetTempPathName() As String
    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
        Dim sBuffer As String
        Dim lRet As Long
        
        sBuffer = String$(255, vbNullChar)
        
        lRet = GetTempPath(255, sBuffer)
        
        If lRet > 0 Then
            
            sBuffer = Left$(sBuffer, lRet)
        
        End If
        
        GetTempPathName = sBuffer
        
    End Function
