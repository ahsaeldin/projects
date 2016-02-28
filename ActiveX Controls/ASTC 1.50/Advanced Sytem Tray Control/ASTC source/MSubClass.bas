Attribute VB_Name = "MSubClass"
Option Explicit

Dim TrayObj As TrayIcon
Dim BalloonObj As Balloon
Public currentmsg As Long
Private OldWindowProc As Long
Private Old_BalloonProc As Long
Public CurrentBalloonMsg As Long

Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim res As Long
Dim ToolTipParent As Long

Const WM_NCDESTROY = &H82

OldWindowProc = OldWinProc(hwnd)

ToolTipParent = GetProp(hwnd, "parent")   'not yet

'we use this to avoid interfaceing with normal icon messages
If ToolTipParent <> 0 Then

Select Case Msg

          'note that
          'we handle both wm_Ballonshow and wm_Ballonhide
          'from within CreateToolTip and Destroy fucntions
          Case WM_LBUTTONUP
               currentmsg = WM_BALLOONLCLK
               Set TrayObj = TrayCollection(str(ToolTipParent))
               TrayObj.RaiseEvents
          Case WM_RBUTTONUP
               currentmsg = WM_BALLOONRCLK
               Set TrayObj = TrayCollection(str(ToolTipParent))
               TrayObj.RaiseEvents
          
             
End Select

End If



If Msg = WM_MYHOOK Then

   Select Case lParam
          
          Case WM_LBUTTONDBLCLK
               currentmsg = WM_LBUTTONDBLCLK
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_LBUTTONDOWN
               currentmsg = WM_LBUTTONDOWN
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_LBUTTONUP
               currentmsg = WM_LBUTTONUP
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_MBUTTONDBLCLK
               currentmsg = WM_MBUTTONDBLCLK
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_MBUTTONDOWN
               currentmsg = WM_MBUTTONDOWN
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_MBUTTONUP
               currentmsg = WM_MBUTTONUP
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_RBUTTONDBLCLK
               currentmsg = WM_RBUTTONDBLCLK
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_RBUTTONDOWN
               currentmsg = WM_RBUTTONDOWN
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_RBUTTONUP
               currentmsg = WM_RBUTTONUP
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_BALLOONHIDE
               currentmsg = WM_BALLOONHIDE
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_BALLOONLCLK
               currentmsg = WM_BALLOONLCLK
               Debug.Print "hwnd is ", hwnd
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_BALLOONRCLK
               currentmsg = WM_BALLOONRCLK
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
          Case WM_BALLOONSHOW
               currentmsg = WM_BALLOONSHOW
               Set TrayObj = TrayCollection(str(hwnd))
               TrayObj.RaiseEvents
   
   End Select

End If

If Msg = WM_NCDESTROY Then
       
   res = SetWindowLong(hwnd, GWL_WNDPROC, OldWindowProc)
   If res = 0 Then LastErrorForMe = "Error Unsubclass"
  
End If

NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)

End Function

Public Property Get OldWinProc(ByVal hwnd As Long) As Long
Dim sName As String
   sName = hwnd
   OldWinProc = GetProp(hwnd, sName)
End Property
Public Property Let OldWinProc(ByVal hwnd As Long, ByVal lPtr As Long)
Dim m_f As Long
Dim sName As String

   m_f = 1
   sName = hwnd
   m_f = SetProp(hwnd, sName, lPtr)
'log error here instead insteds
   If m_f = 0 Then LastErrorForMe = "Error in let OldWinProc"
   If (lPtr = 0) Then
      RemoveProp hwnd, sName
   End If
   
End Property

Public Sub clearVariable()
Set TrayObj = Nothing
End Sub

Public Sub ClearBalloonObj()
Set BalloonObj = Nothing
End Sub

Public Function BalloonWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim res As Long
Dim ToolTipHwnd As Long

Const WM_NCDESTROY = &H82

Old_BalloonProc = OldWinProc(hwnd)

ToolTipHwnd = GetProp(hwnd, "parent")   'not yet

'we use this to avoid interfaceing with normal icon messages
If ToolTipHwnd <> 0 Then

Select Case Msg

          'note that
          'we handle both wm_Ballonshow and wm_Ballonhide
          'from within CreateToolTip and Destroy fucntions
          Case WM_LBUTTONUP
               CurrentBalloonMsg = WM_BALLOONLCLK
               Set BalloonObj = BalloonCollection(str(ToolTipHwnd))
               BalloonObj.RaiseEvents
          Case WM_RBUTTONUP
               CurrentBalloonMsg = WM_BALLOONRCLK
               Set BalloonObj = BalloonCollection(str(ToolTipHwnd))
               BalloonObj.RaiseEvents
          
             
End Select

End If


If Msg = WM_NCDESTROY Then
       
   res = SetWindowLong(hwnd, GWL_WNDPROC, Old_BalloonProc)
   If res = 0 Then LastErrorForMe = "Error Unsubclass"
  
End If

BalloonWindowProc = CallWindowProc(Old_BalloonProc, hwnd, Msg, wParam, lParam)

End Function
