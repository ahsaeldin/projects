VERSION 5.00
Begin VB.UserControl Balloon 
   Appearance      =   0  'Flat
   BackColor       =   &H008080FF&
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Balloon.ctx":0000
   PropertyPages   =   "Balloon.ctx":0974
   ScaleHeight     =   405
   ScaleWidth      =   405
   ToolboxBitmap   =   "Balloon.ctx":0986
   Begin ASTC.UserControl1 UserControl11 
      Left            =   1560
      Top             =   360
      _ExtentX        =   1296
      _ExtentY        =   1296
   End
   Begin ASTC.TrayIcon TrayIcon1 
      Left            =   1680
      Top             =   1920
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4920
      Top             =   1560
   End
   Begin VB.Timer tmrDestroy 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6720
      Top             =   1080
   End
End
Attribute VB_Name = "Balloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Public Event BalloonShow()
Attribute BalloonShow.VB_Description = "Occur when the balloon tooltip is about to be displayed."
Public Event BalloonHide()
Attribute BalloonHide.VB_Description = "Occur when the balloon tooltip is about to be removed."
Public Event BalloonLeftClick()
Attribute BalloonLeftClick.VB_Description = "Occur when the user click the balloon tooltip with the left Mouse button."
Public Event BalloonRightClick()
Attribute BalloonRightClick.VB_Description = "Occur when the user click the balloon tooltip with the Right Mouse button."

'PHWnd Optional

'use this Optional parameter to insure that the balloon tooltip willn't appear
'outside a specifed window or a child window region like a command button or a text box.

'the only usefullness of this parameter is to avoid allowing
'the balloon tooltip form appearing at any part of window or it's Controls

'note that

'1.you can't set this paramter to any extenal window 'i.e you can't set to an extenal Apps window

'2.you can set this paramter to the window which contains the current instance of the balloon control
'i.e
'balloon1.ShowBalloon From1.hWnd
'this hwnd is the handle to the window in which contains the Balloon control
'and that's means that the balloon will only appear if the mouse is over the Form itself
'not any of it's child controls [child windows] in which it contains.
   
'3.you can set this paramter to any control of the window that contains the
'current instance of the balloon control
'i.e if you have a Form called Form1 which has a command button called Command1
'and a text box called text1,you can set the parameter as the following
'balloon1.ShowBalloon command1.hwnd
'balloon1.ShowBalloon text1.hwnd
'and that's means the balloon will only appear if the mouse is over this control itself
     
'4.if you didn't enter this parameter Balloon1 object will use it's parent window automatically
'i.e if you have a form called Form1 ,ASTC will assign PHwnd silently to be a Form1.Hwnd
'but
'in this case the balloon tooltip may appear at any part of form or it's Controls
 
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

Dim X As Long
Dim y As Long

Dim OLD_X As Long
Dim OLD_y As Long

Private ti As TOOLINF

Dim ParentHwnd As Long
Dim Forced_Hwnd As Long

Public Enum BalloonStyle
    BalloonType
    RectangleType
End Enum

Private m_lTTHwnd As Long
Private mvarTitle As String
Private mvarPrompt As String
Private mvarForeColor As Long
Private mvarBackColor As Long
Private mvarStyle As BalloonStyle
Private mvarIconType As BIcoType
Private TerminateState As Boolean


Public Sub ShowBalloon(Optional PHwnd As Long, Optional ByVal Prompt As String, _
Optional ByVal Title As String, _
Optional ByVal IconType As BIcoType = NoIcon, Optional ByVal DelayTime As Long = 2000, _
Optional ByVal Timeout As Long = 5000)
Attribute ShowBalloon.VB_Description = "Display a balloon tooltip."
        
        '//cprin 30Oct2006: added next if clause to fix a bug where the icon of the balloon isn't displayed if the title = ""
        If Title = "" Then '//if the title of the balloon = ""
            Title = " " '//then make the title is equal to space in order to display the balloon tooltip icon
        End If

        'if the user didn't enter a PHwnd parameter
        If PHwnd <> 0 Then
           If IsWindow(PHwnd) = False Then
              ErrRaise eeInvalidWindow, Extender.Name & ".ShowBalloon error"
              Exit Sub
           End If
           
           If IsWindowLocal(PHwnd) = False Then
              ErrRaise eeNoExternalWindow, Extender.Name & ".ShowBalloon error "
              Exit Sub
           End If
        End If
        
        If m_lTTHwnd <> 0 Then Exit Sub
        If Prompt <> Empty Then Prompts = Prompt
        If Title <> Empty Then Titles = Title
        mvarIconType = IconType
        
        If PHwnd <> 0 Then
           'clear ParentHwnd if it had a value from a pervious calling
           'in order to avoid Error in tmrShow_Timer
           ParentHwnd = 0
           ForcedHwnd = PHwnd
        Else 'that's means user didn't enter the PHwnd paramter
           'clear ForcedHwnd if it had a value from a pervious
           'in order to avoid Error in tmrShow_Timer
           ForcedHwnd = 0
           On Error Resume Next
           ParentHwnd = UserControl.Parent.hwnd
           On Error GoTo 0
        End If
        
        
'        Debug.Print ParentHwnd
        tmrShow.Interval = DelayTime
        tmrShow.Enabled = True
        If Timeout > 0 Then tmrDestroy.Interval = Timeout: tmrDestroy.Enabled = True
      


End Sub

Private Function CreateToolTip(ByVal lParentHwnd As Long, X As Long, y As Long) As Boolean
    Dim ptTip As Long
    Dim ret As Long
    Dim lWinStyle As Long
    Dim lpRect As RECT
    Dim ToolTipOldProc As Long
     

lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
       
If mvarStyle = BalloonType Then lWinStyle = lWinStyle Or TTS_BALLOON
    
If m_lTTHwnd <> 0 Then Destroy
    
m_lTTHwnd = UserControl11.CreateToolTip2(lParentHwnd, lWinStyle)


    
    
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

    If mvarForeColor <> Empty Then
        ret = SendMessage(m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&)
    End If

    If mvarBackColor <> Empty Then
        ret = SendMessage(m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&)
    End If

    ptTip = y * &H10000 + X

    'These two messages will set the position of the tooltip:
    ret = SendMessage(m_lTTHwnd, TTM_TRACKPOSITION, 0, ByVal ptTip)
    ret = SendMessage(m_lTTHwnd, TTM_TRACKACTIVATE, True, ti)
    
     RaiseEvent BalloonShow
    
     ToolTipOldProc = SetWindowLong(m_lTTHwnd, GWL_WNDPROC, AddressOf BalloonWindowProc)
      
     BalloonCollection.Add Me, str(m_lTTHwnd)
   
     OldWinProc(m_lTTHwnd) = ToolTipOldProc
     
    'ParentHwnd here is the form not syspager

     SetProp m_lTTHwnd, "parent", m_lTTHwnd
     
    
     If lParentHwnd = 0 Then
      'make tooltip window a topmost window
      'ret = SetWindowPos(m_lTTHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
      ret = SetWindowPos(m_lTTHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE)
     End If
       
    
     CreateToolTip = True


End Function

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

Public Sub Destroy()
Attribute Destroy.VB_Description = "Close the balloon tooltip."

Dim old As Long
Dim i As Integer
Dim CHwnd As Long
Dim BalloonObject As Balloon

tmrDestroy.Enabled = False

If m_lTTHwnd <> 0 Then

   old = OldWinProc(m_lTTHwnd)
    
   SetWindowLong m_lTTHwnd, GWL_WNDPROC, old
      
   RemoveProp m_lTTHwnd, "parent"
      
   DestroyWindow m_lTTHwnd
      
   m_lTTHwnd = 0
      
   For i = 1 To BalloonCollection.Count
        
        Set BalloonObject = BalloonCollection(i)
        CHwnd = BalloonObject.Parent
        If CHwnd = ParentHwnd Then
           BalloonCollection.Remove (i)
           Exit For
        End If
          
    Next i
    
   If Not Terminateing Then
     RaiseEvent BalloonHide
   End If
End If

End Sub

Private Sub tmrDestroy_Timer()
  Destroy
End Sub

Private Property Get Terminateing() As Boolean
Terminateing = TerminateState
End Property

Private Property Let Terminateing(ByVal vNewValue As Boolean)
TerminateState = vNewValue
End Property

Private Sub tmrShow_Timer()

Dim res As Boolean
Dim Pointa As POINTAPI
Dim HwndFromPoint As Long

'Only show if doesn't Exist
If m_lTTHwnd = 0 Then
        
        Call GetCursorPos(Pointa)
        Xpos = Pointa.X
        YPOS = Pointa.y
        If Xpos = OLDX And YPOS = OLDy Then
           tmrShow.Enabled = False
           Exit Sub
        End If
        
        HwndFromPoint = WindowFromPoint(Xpos, YPOS)
        If ForcedHwnd <> 0 Then
           'without the following important line
           'the balloon may be displayed outside it's assosited control
           If ForcedHwnd <> HwndFromPoint Then Exit Sub
           res = CreateToolTip(ForcedHwnd, Xpos, YPOS)
        Else
           res = CreateToolTip(Parent_Hwnd, Xpos, YPOS)
        End If
        
        If Not res Then
           LastErrorForMe = "can't show ballon"
           Exit Sub
        End If
        
        OLD_X = Pointa.X
        OLD_y = Pointa.y

End If


tmrShow.Enabled = False

End Sub

Friend Sub RaiseEvents()

Select Case CurrentBalloonMsg
       
       
       Case WM_BALLOONLCLK
            RaiseEvent BalloonLeftClick
            CurrentBalloonMsg = 0
         
       Case WM_BALLOONRCLK
            RaiseEvent BalloonRightClick
            CurrentBalloonMsg = 0
         
      

End Select
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 405
UserControl.Height = 405
End Sub

Private Sub UserControl_Terminate()
Terminateing = True
Destroy
MSubClass.ClearBalloonObj
Set BalloonCollection = Nothing
End Sub

Private Property Get Parent_Hwnd() As Long
Parent_Hwnd = ParentHwnd
End Property

Private Property Let Parent_Hwnd(ByVal vNewValue As Long)
ParentHwnd = vNewValue
End Property

Private Property Get ForcedHwnd() As Long
ForcedHwnd = Forced_Hwnd
End Property

Private Property Let ForcedHwnd(ByVal vNewValue As Long)
Forced_Hwnd = vNewValue
End Property

Private Property Get Xpos() As Long
Xpos = X
End Property

Private Property Let Xpos(ByVal vNewValue As Long)
X = vNewValue
End Property

Private Property Get YPOS() As Long
YPOS = y
End Property

Private Property Let YPOS(ByVal vNewValue As Long)
y = vNewValue
End Property


Private Property Get OLDX() As Long
OLDX = OLD_X
End Property

Private Property Let OLDX(ByVal vNewValue As Long)
OLD_X = vNewValue
End Property

Private Property Get OLDy() As Long
OLDy = OLD_y
End Property

Private Property Let OLDy(ByVal vNewValue As Long)
OLD_y = vNewValue
End Property

Friend Function Parent() As Long
If ForcedHwnd <> 0 Then Parent = ForcedHwnd Else Parent = ParentHwnd
End Function
Public Property Let ForeColor(ByVal vData As Long)
Attribute ForeColor.VB_Description = "returns or sets the font color of the balloon tooltip."
   mvarForeColor = vData
   If m_lTTHwnd <> 0 Then
      SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
   End If
End Property

Public Property Get ForeColor() As Long
   ForeColor = mvarForeColor
End Property
Public Property Let BackColor(ByVal vData As Long)
Attribute BackColor.VB_Description = "returns or sets the back color of the balloon tooltip."
   mvarBackColor = vData
   If m_lTTHwnd <> 0 Then
      SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
   End If
End Property
Public Property Get BackColor() As Long
   BackColor = mvarBackColor
End Property
Public Property Let Style(ByVal vData As BalloonStyle)
Attribute Style.VB_Description = "returns or sets the balloon tooltip shape [balloon or Rectangle]."
   mvarStyle = vData
End Property

Public Property Get Style() As BalloonStyle
    Style = mvarStyle
End Property
