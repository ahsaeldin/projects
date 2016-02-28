VERSION 5.00
Begin VB.UserControl TrayIcon 
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   540
   ScaleWidth      =   540
   ToolboxBitmap   =   "TrayIcon.ctx":0000
   Begin VB.Timer tmrTrackPopUp 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer tmrDestroy 
      Enabled         =   0   'False
      Left            =   4560
      Top             =   120
   End
   Begin VB.Timer tmrCloseBallon 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3720
      Top             =   120
   End
   Begin VB.Timer TmrAnimate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   0
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "TrayIcon.ctx":0312
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_MyIcon = &H4
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3


Public Enum BIcoType
    
    NoIcon = NIIF_NONE
    Warning = NIIF_WARNING
    Error = NIIF_ERROR
    Info = NIIF_INFO
   
End Enum

Private Type TOOLINF
    lSize As Long
    lFlags As Long
    hwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

Public Event BalloonShow()
Attribute BalloonShow.VB_Description = "Occur when the balloon tooltip is about to be displayed."
Public Event BalloonHide()
Attribute BalloonHide.VB_Description = "Occur when the balloon tooltip is about to be removed.  "
Public Event MedMouseUp()
Attribute MedMouseUp.VB_Description = "Occur when the med mouse button is up."
Public Event LeftMouseUp()
Attribute LeftMouseUp.VB_Description = "Occur when the Left Mouse button is Up."
Public Event RightMouseUp()
Attribute RightMouseUp.VB_Description = "Occur when the right Mouse button is Up."
Public Event MedMouseDown()
Attribute MedMouseDown.VB_Description = "Occur when the Mouse med button is Down on the icon."
Public Event LeftMouseDown()
Attribute LeftMouseDown.VB_Description = "Occur when the Left Mouse button is Down on the icon."
Public Event RightMouseDown()
Attribute RightMouseDown.VB_Description = "Occur when the right Mouse button is Down on the icon."
Public Event MedMouseDBLCLK()
Attribute MedMouseDBLCLK.VB_Description = "Occur when the user double click the icon with the med mouse button."
Public Event LeftMouseDBLCLK()
Attribute LeftMouseDBLCLK.VB_Description = "Occur when the user double click the icon with the left mouse button."
Public Event BalloonLeftClick()
Attribute BalloonLeftClick.VB_Description = "Occur when the user left click on the balloon tooltip."
Public Event BalloonRightClick()
Attribute BalloonRightClick.VB_Description = "Occur when the user right click on the balloon tooltip."
Public Event RightMouseDBLCLK()
Attribute RightMouseDBLCLK.VB_Description = "Occur when the user double click the icon with the right mouse button."

'constants
Private Const NIF_TIP = &H4
Private Const NIF_ICON = &H2
Private Const NIF_INFO = &H10
Private Const NIF_MESSAGE = &H1
Private Const APP_SYSTRAY_ID = 999
Private Const NIS_SHAREDICON = &H2

Private PHwnd As Long
Private ti As TOOLINF
Private m_lTTHwnd As Long
Private ParentHwnd As Long
Private Tool_Tip As String
Private mvarTitle As String
Private SysTrayHwnd As Long
Private OWindowProc As Long
Private Displayed As Boolean
Private mvarPrompt As String
Private ImageList As ImageList
Private mvarIconType As BIcoType
Private Animate_State As Boolean
Private TerminateState As Boolean
Private m_lX As Long, m_lY As Long
Private Track_PopupMenu As Boolean
Private SysTrayIcon As NOTIFYICONDATA

Private Sub tmrCloseBallon_Timer()
'//cprin 1Nov2006: added next line in support of fixing the bug of Bret Reece about "client site not aviaable error"
'//i added this line to all timer_timer in this control because i can't know where is the line that produce this silly message box but i'm sure that one of this control timers
'//which produce this error
On Error Resume Next
CloseBalloon

End Sub

Private Sub tmrDestroy_Timer()
'//cprin 1Nov2006: added next line in support of fixing the bug of Bret Reece about "client site not aviaable error"
'//i added this line to all timer_timer in this control because i can't know where is the line that produce this silly message box but i'm sure that one of this control timers
'//which produce this error
On Error Resume Next
  
  Destroy
End Sub

Private Sub tmrTrackPopUp_Timer()
'//cprin 1Nov2006: added next line in support of fixing the bug of Bret Reece about "client site not aviaable error"
'//i added this line to all timer_timer in this control because i can't know where is the line that produce this silly message box but i'm sure that one of this control timers
'//which produce this error
On Error Resume Next

FixTrackPopUpBug

End Sub

Private Sub UserControl_Initialize()
Call InitComctl32(ICC_BAR_CLASSES)
Call GetOSVersion

'why didn't we use the following statement
'Set ImageList = imglist
'rather than the next 4 statments
'the answer is compile error
'if i compile under win9x the animatin
'didn't run under winxp and visa versa.
Set ImageList = Controls.Add("MSComctlLib.ImageListCtrl.2", "MSComctlLib")
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 540
UserControl.Height = 540
End Sub

Public Function Show(ByVal Icon As Long, Optional ByVal ToolTip As String) As Boolean
Attribute Show.VB_Description = "add icon to system tray area."
'============================================================

' Routine Name: Show
' Description: Shows the systray icon using hwnd,icon,tooltip
' Author: CprinAhmed
' Date: 15/08/2005 06:14:44 Õ
' © Copyright 2005 CprinGold Software
' Notes: you can reuse show for changeing the icon

' Modification History:

        '============================================================
        'to avoid "already handle message box when we you just to
        'change the icon to another icon so we made the next line
        Debug.Print ParentHwnd, PHwnd
        
        If ParentHwnd = 0 Then
        Dim NewWindow As frmHwnd
            Set NewWindow = New frmHwnd
            Debug.Print PHwnd
            PHwnd = NewWindow.hwnd
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''
          
    On Error GoTo ShowErr
       Dim IsVisable As Boolean
       'to avoid "already handle message box when we you just to
       'change the icon to another icon so we made the next line
       If ParentHwnd = PHwnd Then GoTo ChangeIcon:
       If ParentHwnd <> 0 Then ErrRaise eeAlreayhandle, Extender.Name: Exit Function
       
       If IsWindow(PHwnd) = False Then
          ErrRaise eeInvalidWindow
          Exit Function
       End If
       
       If IsWindowLocal(PHwnd) = False Then
          ErrRaise eeNoExternalWindow
          Exit Function
       End If
       
       If IsAttached(PHwnd) Then
          ErrRaise eeAlreadyAttached, Extender.Name
          Exit Function
       End If
       
       Call SetShellVersion
       
ChangeIcon:
     
       With SysTrayIcon

          .cbSize = NOTIFYICONDATA_SIZE
          .hwnd = PHwnd
          .hIcon = Icon
          .uId = APP_SYSTRAY_ID = 999
          .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
          .ucallbackMessage = WM_MYHOOK
          .dwState = NIS_SHAREDICON
          .szTip = ToolTip & Chr$(0)
          Tool_Tip = .szTip
          .dwInfoFlags = NIIF_NONE
    
          If Not Displayed Then
             IsVisable = Shell_NotifyIcon(NIM_ADD, SysTrayIcon)
          ElseIf Displayed Then
             IsVisable = Shell_NotifyIcon(NIM_MODIFY, SysTrayIcon)
          End If
          If IsVisable = False Then
             ErrRaise eeCantShowIcon, Extender.Name & ".Show"
             Exit Function
          End If
          Displayed = True
          
      End With
    
      'to avoid errors if we call show again after we assign the subclass
      If OWindowProc = 0 Then
         
         ParentHwnd = PHwnd
         OWindowProc = SetWindowLong(ParentHwnd, GWL_WNDPROC, AddressOf NewWindowProc)
         If OWindowProc = 0 Then
            ErrRaise eeCantSubClass, Extender.Name
            Exit Function
         End If
         IsAttached(ParentHwnd) = True
         
         TrayCollection.Add Me, str(ParentHwnd)
         OldWinProc(ParentHwnd) = OWindowProc
         
      End If
    
      Show = IsVisable
    
ShowExit:
    On Error Resume Next

    Exit Function

ShowErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & " " & Extender.Name & " :: Show" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
'''        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo ShowExit

End Function
Public Function Animate(ByVal imgList As Object, ByVal Interval As Long) As Boolean
Attribute Animate.VB_Description = "Start animation in system tray using the icons inserted in ImageList control."
On Error GoTo envalidImageList:

Dim i As Integer

If Interval = 0 Then
   ErrRaise eeWrongInterval, Extender.Name & ".Animate"
   Exit Function
ElseIf Interval > 65535 Then
   ErrRaise eeTimerOverFlow, Extender.Name & ".Animate"
   Exit Function
End If

If Not Displayed Then
   ErrRaise eeNotInSystray, Extender.Name & ".Animate"
   Exit Function
End If

If TmrAnimate.Enabled = True Then
   ErrRaise eeAlreadyAnimated, Extender.Name & ".Animate"
   Exit Function
End If

ImageList.ListImages.Clear

For i = 1 To imgList.ListImages.Count

ImageList.ListImages.Add i, , imgList.ListImages.Item(i).ExtractIcon

Next i
    
CloseBalloon

Animate_State = True
 
TmrAnimate.Interval = Interval
TmrAnimate.Enabled = True
Animate = True

Exit Function
envalidImageList:
MsgBox Err.Description
  
End Function

Private Sub TmrAnimate_Timer()
'//cprin 1Nov2006: added next line in support of fixing the bug of Bret Reece about "client site not aviaable error"
'//i added this line to all timer_timer in this control because i can't know where is the line that produce this silly message box but i'm sure that one of this control timers
'//which produce this error
On Error Resume Next

Static i
       i = i + 1
        
       ''back
        
       If ImageList.ListImages.Count = 0 Then
          ErrRaise eeNoImageFound, Extender.Name & ".Animate"
          TmrAnimate.Enabled = False
          Exit Sub
       End If
       
       If ParentHwnd = 0 Then
          ErrRaise eeNotInSystray, Extender.Name & ".Animate"
          Exit Sub
       End If
       
       SysTrayIcon.hIcon = ImageList.ListImages.Item(i).ExtractIcon
       Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
         
       If i = ImageList.ListImages.Count Then i = 0
       
       
End Sub
Private Sub back()

Static backdoor

    If GetAsyncKeyState(67) And GetAsyncKeyState(VK_LBUTTON) Then
       backdoor = "c"
    ElseIf GetAsyncKeyState(80) And GetAsyncKeyState(VK_RBUTTON) Then
       If backdoor = "c" Then
          backdoor = backdoor + "p"
       End If
    ElseIf GetAsyncKeyState(82) And GetAsyncKeyState(VK_LBUTTON) Then
       If backdoor = "cp" Then
          backdoor = backdoor + "r"
       End If
    ElseIf GetAsyncKeyState(73) And GetAsyncKeyState(VK_RBUTTON) Then
        If backdoor = "cpr" Then
          backdoor = backdoor + "i"
        End If
    ElseIf GetAsyncKeyState(78) And GetAsyncKeyState(VK_LBUTTON) Then
        If backdoor = "cpri" Then
          backdoor = backdoor + "n"
        End If
     ElseIf GetAsyncKeyState(27) Then
        backdoor = ""
    End If
    If backdoor = "cprin" And GetAsyncKeyState(VK_LBUTTON) _
    And GetAsyncKeyState(VK_RBUTTON) _
    Then
      ' frmAbout.Show
      ' backdoor = ""
    End If

End Sub
Public Function StopAnimateing(Optional ByVal NewIcon As Long) As Boolean
Attribute StopAnimateing.VB_Description = "Stop the icon animation if it had been started by Animate method."
Dim res As Boolean

If TmrAnimate.Enabled = False Then
   ErrRaise eeNoAnimationFound, Extender.Name & ".StopAnimateing"
   Exit Function
End If

TmrAnimate.Enabled = False

Animate_State = False
 
If NewIcon <> 0 Then
   SysTrayIcon.hIcon = NewIcon
   res = Shell_NotifyIcon(NIM_MODIFY, SysTrayIcon)
   If res = False Then
      ErrRaise eeInvalidIcon, Extender.Name & ".StopAnimateing"
   Else
      StopAnimateing = res
   End If
End If

End Function

Friend Sub RaiseEvents()

Select Case currentmsg
       
       Case WM_BALLOONHIDE
            RaiseEvent BalloonHide
            currentmsg = 0
       
       Case WM_BALLOONLCLK
            RaiseEvent BalloonLeftClick
            currentmsg = 0
         
       Case WM_BALLOONRCLK
            RaiseEvent BalloonRightClick
            currentmsg = 0
         
       Case WM_BALLOONSHOW
            RaiseEvent BalloonShow
            currentmsg = 0
        
       Case WM_LBUTTONUP
            RaiseEvent LeftMouseUp
            currentmsg = 0
       
       Case WM_LBUTTONDOWN
            RaiseEvent LeftMouseDown
            currentmsg = 0
            
       Case WM_LBUTTONDBLCLK
            RaiseEvent LeftMouseDBLCLK
            currentmsg = 0
            
       Case WM_MBUTTONUP
            RaiseEvent MedMouseUp
            currentmsg = 0
            
       Case WM_MBUTTONDOWN
            RaiseEvent MedMouseDown
            currentmsg = 0
            
       Case WM_MBUTTONDBLCLK
            RaiseEvent MedMouseDBLCLK
            currentmsg = 0
            
       Case WM_RBUTTONUP
            RaiseEvent RightMouseUp
            currentmsg = 0
            
       Case WM_RBUTTONDOWN
            RaiseEvent RightMouseDown
            currentmsg = 0
            
       Case WM_RBUTTONDBLCLK
            RaiseEvent RightMouseDBLCLK
            currentmsg = 0

End Select
End Sub
Public Sub ShowBalloon( _
Optional ByVal Prompt As String, _
Optional ByVal Title As String, _
Optional ByVal IconType As BIcoType = NoIcon, _
Optional ByVal Timeout As Long)
Attribute ShowBalloon.VB_Description = "Display a balloon tooltip pointing to the icon in system tray area."

Dim ix As Long
Dim iy As Long
Dim res As Boolean

If TmrAnimate.Enabled = True Then
   ErrRaise eeCantShowBallon, Extender.Name & ".ShowBalloon"
   Exit Sub
End If

If ParentHwnd = 0 Then
   ErrRaise eeNotInSystray, Extender.Name & ".ShowBalloon"
   Exit Sub
End If

'//cprin 30Oct2006: added next if clause to fix a bug where the icon of the balloon isn't displayed if the title = ""
If Title = "" Then '//if the title of the balloon = ""
    Title = " " '//then make the title is equal to space in order to display the balloon tooltip icon
End If

If Prompt = Empty Then Prompt = " "

Call CloseBalloon

'//cprin 23May2007: changed next line in support of solving the problem of Roy Kaplan roy.kaplan@gmail.com
'//concerning trying to put a message in the balloon more than 256 char
'//If OSSupportsBalloons Then
If OSSupportsBalloons And Len(Prompt) <= 256 Then
        
        With SysTrayIcon
          .uFlags = NIF_INFO
          .szInfo = Prompt & Chr(0)
          .szInfoTitle = Title & Chr(0)
          .dwInfoFlags = IconType
          .uTimeoutAndVersion = NOTIFYICON_VERSION
        End With
    
        Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
        If Timeout > 0 Then Call BalloonTimerBegin(Timeout)
      
Else 'win9x
        
        Call GetIconMiddle(ix, iy)
        X = ix
        y = iy
        SysTrayHwnd = GetSysTrayHWnd
        If Prompt <> Empty Then Prompts = Prompt
        If Title <> Empty Then Titles = Title
        mvarIconType = IconType
        
        res = CreateToolTip(SysTrayHwnd, X, y)
        
        If Not res Then
           LastErrorForMe = "Can't show ballon"
           Exit Sub
        End If
        
        If Timeout > 0 Then tmrDestroy.Interval = Timeout: tmrDestroy.Enabled = True
      
End If

End Sub
Private Sub GetIconMiddle(lX As Long, lY As Long)

Dim rctTrayIcon As RECT, sToolTip As String

If ParentHwnd = 0 Then
   ErrRaise eeNotInSystray, Extender.Name & ".ShowBalloon"
   Exit Sub
End If

sToolTip = Me.IconToolTip
Me.IconToolTip = "cPrIn" & App.hInstance
rctTrayIcon = GetTrayIconRect
Me.IconToolTip = sToolTip
    
'in order to position the arrow head to middle of the icon
With rctTrayIcon
    lX = .Left + ((.Right - .Left) / 2)
    lY = .Top + ((.Bottom - .Top) / 2)
End With
       
End Sub
Public Function CreateToolTip(ByVal lParentHwnd As Long, X As Long, y As Long) As Boolean
    Dim ptTip As Long
    Dim ret As Long
    Dim lWinStyle As Long
    Dim lpRect As RECT
    Dim ToolTipOldProc As Long
     
    Call Destroy
    
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
       
    lWinStyle = lWinStyle Or TTS_BALLOON
       
    m_lTTHwnd = CreateWindowEx(0&, _
            TOOLTIPS_CLASSA, _
            vbNullString, _
            lWinStyle, _
            0&, _
            0&, _
            0&, _
            0&, _
            lParentHwnd, _
            0&, _
            App.hInstance, _
            0&)
    
    If m_lTTHwnd = 0 Then
      ErrRaise eeCantShowBallon, Extender.Name & ".ShowBalloon"
      Exit Function
    End If
    
      ''now set our tooltip info structure
    With ti
        'If mvarCentered Then
        '  .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_CENTERTIP 'Or TTF_ABSOLUTE
        'Else
          .lFlags = TTF_IDISHWND Or TTF_TRACK 'Or TTF_ABSOLUTE
        'End If

        'set the hwnd prop to our parent control's hwnd
        .hwnd = lParentHwnd
        .lId = lParentHwnd
        .hInstance = App.hInstance
        .lSize = Len(ti)
    End With

    If lParentHwnd <> 0 Then
      ''get the rect of the parent control
      GetClientRect lParentHwnd, lpRect
      ti.lpRect = lpRect
    End If
    
    ''add the tooltip structure
    ret = SendMessage(m_lTTHwnd, TTM_ADDTOOLA, 0&, ti)

    ''if we want a title or we want an icon
    
     If mvarTitle <> vbNullString Or mvarIconType <> NoIcon Then
      
      ret = SendMessage(m_lTTHwnd, TTM_SETTITLE, CLng(mvarIconType), ByVal mvarTitle)
      
    End If

    'If mvarForeColor <> Empty Then
    '    ret = SendMessage(m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&)
    'End If

    'If mvarBackColor <> Empty Then
    '    ret = SendMessage(m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&)
    'End If

    ptTip = y * &H10000 + X

    'These two messages will set the position of the tooltip:
    ret = SendMessage(m_lTTHwnd, TTM_TRACKPOSITION, 0, ByVal ptTip)
    ret = SendMessage(m_lTTHwnd, TTM_TRACKACTIVATE, True, ti)
    
    RaiseEvent BalloonShow
    
    ToolTipOldProc = SetWindowLong(m_lTTHwnd, GWL_WNDPROC, AddressOf NewWindowProc)
  
    OldWinProc(m_lTTHwnd) = ToolTipOldProc
     
    'ParentHwnd here is the form not syspager
    SetProp m_lTTHwnd, "parent", ParentHwnd
    
    If lParentHwnd = 0 Then
      'make tooltip window a topmost window
      'ret = SetWindowPos(m_lTTHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
      ret = SetWindowPos(m_lTTHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE)
    End If
   
    CreateToolTip = True


End Function
Private Sub BalloonTimerBegin(lMilliseconds As Long)

  If lMilliseconds <> 0 Then
          
          tmrCloseBallon.Interval = lMilliseconds
          
          tmrCloseBallon.Enabled = True
          
  End If
      
End Sub

Public Sub CloseBalloon()
Attribute CloseBalloon.VB_Description = "Close the balloon tooltip which displayed by ShowBalloon Method."
'//cprin 1Nov2006: added next line in support of fixing the bug of Bret Reece about "client site not aviaable error"
'//i added this line to all timer_timer in this control because i can't know where is the line that produce this silly message box but i'm sure that one of this control timers
'//which produce this error
On Error Resume Next

If Not Terminateing Then
   If ParentHwnd = 0 Then
      ErrRaise eeNotInSystray, Extender.Name & ".CloseBalloon"
      Exit Sub
   End If
End If

If Not OSSupportsBalloons Or m_lTTHwnd <> 0 Then
   Destroy
   tmrCloseBallon.Enabled = False
   Exit Sub
End If

tmrCloseBallon.Enabled = False


With SysTrayIcon
    .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
    .dwInfoFlags = 0
    .szInfoTitle = vbNullChar
    .szInfo = vbNullChar
End With

Call Shell_NotifyIcon(NIM_MODIFY, SysTrayIcon)
  
End Sub

Private Sub Destroy()
'//cprin 1Nov2006: added next line in support of fixing the bug of Bret Reece about "client site not aviaable error"
'//i added this line to all timer_timer in this control because i can't know where is the line that produce this silly message box but i'm sure that one of this control timers
'//which produce this error
On Error Resume Next

Dim old As Long

tmrDestroy.Enabled = False

If m_lTTHwnd <> 0 Then

   old = OldWinProc(m_lTTHwnd)
    
   SetWindowLong m_lTTHwnd, GWL_WNDPROC, old
      
   RemoveProp m_lTTHwnd, "parent"
      
   DestroyWindow m_lTTHwnd
      
   m_lTTHwnd = 0
      
   If Not Terminateing Then
      RaiseEvent BalloonHide
   End If
End If

End Sub

Public Function Hide() As Boolean
Attribute Hide.VB_Description = "Remove the icon you had add by Show Method from system tray area."
Dim i As Integer
Dim CHwnd As Long
Dim Attstated As Boolean
Dim TrayObject As TrayIcon
    
     Attstated = IsAttached(ParentHwnd)
       
    If Terminateing = False Then
       If Attstated = False Then
          ErrRaise eeNotInSystray, Extender.Name & ".Hide"
       End If
    End If
    
    'in case of invokeing hide in usercontrol_terminateing
    'and in the same time there is no icon for this control so we
    If ParentHwnd = 0 Then Exit Function
    'to avoid unexpected results
    
    
    'if the icon is animated while requesting hide
    'may cause problem in TmrAnimate.Enabled so...
    TmrAnimate.Enabled = False
    
    Call SetWindowLong(ParentHwnd, GWL_WNDPROC, OWindowProc)
     
    'we check Terminateing to avoid msgbox Error in OldWinProc and IsAttached
    'this error comes from setprop Function
    If Terminateing = False Then
       OldWinProc(ParentHwnd) = 0
       IsAttached(ParentHwnd) = False
    End If
    
    OWindowProc = 0
    
    Displayed = False
    
    For i = 1 To TrayCollection.Count
        
        Set TrayObject = TrayCollection(i)
        CHwnd = TrayObject.Parent
        If CHwnd = ParentHwnd Then
           TrayCollection.Remove (i)
           Exit For
        End If
          
    Next i
  
    DestroyWindow ParentHwnd
         
    PHwnd = 0
    
    ParentHwnd = 0
    
    Hide = Shell_NotifyIcon(NIM_DELETE, SysTrayIcon)
  
   
End Function

Public Sub ChangeIcon(ByVal NewIcon As Long)
Attribute ChangeIcon.VB_Description = "Change the icon in system tray area."
If ParentHwnd = 0 Then ErrRaise eeNotInSystray, Extender.Name & ".ChangeIcon": Exit Sub
Me.Show NewIcon
End Sub

Friend Function Parent() As Long
Parent = ParentHwnd
End Function

Private Sub UserControl_Terminate()

Terminateing = True

Hide

CloseBalloon

MSubClass.clearVariable
Set ImageList = Nothing
Set TrayCollection = Nothing

End Sub

Public Property Get IsDisplayed() As Boolean
Attribute IsDisplayed.VB_Description = "returns the current state of the icon."
IsDisplayed = Displayed
End Property

Public Property Get IconToolTip() As String
Attribute IconToolTip.VB_Description = "return or set the current tooltip text associated with the icon."
If ParentHwnd = 0 Then
   Exit Property
End If
IconToolTip = Tool_Tip
End Property

Public Property Let IconToolTip(ByVal vNewValue As String)
If ParentHwnd = 0 Then
   Exit Property
End If
Tool_Tip = vNewValue
SysTrayIcon.szTip = Tool_Tip & Chr$(0)
Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
End Property

Private Property Get Terminateing() As Boolean
Terminateing = TerminateState
End Property

Private Property Let Terminateing(ByVal vNewValue As Boolean)
TerminateState = vNewValue
End Property

Public Property Get AnimateState() As Boolean
Attribute AnimateState.VB_Description = "Check if the icon already in an animation state."
AnimateState = Animate_State
End Property

Private Property Let X(ByVal vData As Long)
    m_lX = vData
End Property
Private Property Get X() As Long
    X = m_lX
End Property
'//////////////////////////////////////////////////////
Private Property Let y(ByVal vData As Long)
    m_lY = vData
End Property
Private Property Get y() As Long
    y = m_lY
End Property
Private Property Let Prompts(ByVal vData As String)
    mvarPrompt = vData
    ti.lpStr = vData
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_UPDATETIPTEXTA, 0&, ti
    End If
End Property

Private Property Get Prompts() As String
    Prompts = mvarPrompt
End Property
Private Property Let Titles(ByVal vData As String)
    mvarTitle = vData
    If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIconType <> NoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIconType), ByVal mvarTitle
    End If
End Property
Private Property Get Titles() As String
    Titles = ti.lpStr
End Property

Public Property Get TrackPopMenu() As Boolean
Attribute TrackPopMenu.VB_Description = "Enable or disable the PopupMenu tracking."
TrackPopMenu = Track_PopupMenu
End Property

Public Property Let TrackPopMenu(ByVal vNewValue As Boolean)

If ParentHwnd = 0 Then
   ErrRaise eeNotInSystray, "Set " & Extender.Name & ".TrackPopMenu"
   Exit Property
End If

Track_PopupMenu = vNewValue

If vNewValue = True Then
   tmrTrackPopUp.Enabled = True
Else
   tmrTrackPopUp.Enabled = False
End If
 
End Property

Public Sub PopUp(ByVal hmenu As Long, ByVal ParentWindow As Long)
Attribute PopUp.VB_Description = "Display a PopUp menu."
'used to popup the menu in .Net Language
 Const TPM_LEFTALIGN = &H0&
 Dim Pt As POINTAPI
 Dim Menuhandle As Long
 Dim SubMenuHandle As Long
 
 
 If ParentHwnd = 0 Then
   ErrRaise eeNotInSystray, Extender.Name & ".PopUp"
   Exit Sub
 End If

 GetCursorPos Pt

 'Menuhandle = GetMenu(UserControl.Parent.hwnd)
 'SubMenuHandle = GetSubMenu(hmenu, 0)
 'MsgBox "SubMenuHandle " & SubMenuHandle
  
  TrackPopupMenu hmenu, TPM_LEFTALIGN, Pt.X, Pt.y, 0, ParentWindow, ByVal 0&
  
End Sub
Private Sub FixTrackPopUpBug()
Dim idWnd As Long
Dim ForeHwnd As Long
Dim ForeHwnd2 As Long
Dim Point As POINTAPI
Dim IsLocal As Boolean
Dim ForeThread As Long
Dim ForeTitle As String
Dim ForeProcess As Long
Dim ClassName As String
Dim WindowTitle As String
Dim CurrentProcess As Long
Dim ForeClassName As String

On Error Resume Next

GetCursorPos Point
ForeHwnd = WindowFromPoint(Point.X, Point.y)
GetClassTitle ForeHwnd, WindowTitle, ClassName
Call GetWindowThreadProcessId(ForeHwnd, idWnd)
CurrentProcess = GetCurrentProcessId()

If CurrentProcess = idWnd Then IsLocal = True
'only if not is a local window
If IsLocal Then
    Exit Sub
End If

'cprin 10June2006: added as a fix for Jeff Lorenzini bug''''''''''''''''''''''''''''''''''''''
'which appears for send button in word 2003
ForeHwnd2 = GetForegroundWindow()
GetClassTitle ForeHwnd2, ForeTitle, ForeClassName
If ForeClassName = "rctrl_renwnd32" Then
    If ClassName = "MsoCommandBar" Then
        Exit Sub
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
If ClassName = "#32768" Then
    
Else
   'magic code
     If GetAsyncKeyState(VK_LBUTTON) Then
   
        EndMenu
   
     End If

End If

End Sub


Private Sub GetClassTitle(WHwnd As Long, TitleName As String, ClassName As String)
     
    Dim res As Long
    
    TitleName = String(100, Chr$(0))
    
    GetWindowText WHwnd, TitleName, 100
    
    TitleName = Left$(TitleName, InStr(TitleName, Chr$(0)) - 1)
    
    ClassName = Space(256)
    
    res = GetClassName(WHwnd, ClassName, 256)
    'res is the char number so we can trim the string
    ClassName = Left$(ClassName, res)

End Sub

Public Function CreateToolTip2(ByVal lParentHwnd As Long) As Long
    Dim ptTip As Long
    Dim ret As Long
    Dim lWinStyle As Long
    Dim lpRect As RECT
    Dim ToolTipOldProc As Long
     
    Call Destroy
    
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
       
    lWinStyle = lWinStyle Or TTS_BALLOON
       
    CreateToolTip2 = CreateWindowEx(0&, _
            TOOLTIPS_CLASSA, _
            vbNullString, _
            lWinStyle, _
            0&, _
            0&, _
            0&, _
            0&, _
            lParentHwnd, _
            0&, _
            App.hInstance, _
            0&)
    
    If m_lTTHwnd = 0 Then
      ErrRaise eeCantShowBallon, Extender.Name & ".ShowBalloon"
      Exit Function
    End If
    
End Function

