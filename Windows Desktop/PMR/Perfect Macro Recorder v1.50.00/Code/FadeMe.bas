Attribute VB_Name = "Module1"
Option Explicit

'cprin 13July2006: added next 3 lines in support of make a global constants for
'Email address, Software Name and SiteName in case that any of them changes not
'to change in the whole Project.
Public Const Email = "support@perfectiontools.com"
Public Const SoftwareName = "Perfect Macro Recorder"
Public Const SiteName = "http://www.perfectiontools.com"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Sub FadeMe()
Dim i As Integer
    'KPD-Team 2000
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim Ret As Long
    'Set the window style to 'Layered'
    Ret = GetWindowLong(frmJMacro.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    'Set the opacity of the layered window to 128
   For i = 0 To 51
      SetWindowLong frmJMacro.hWnd, GWL_EXSTYLE, Ret
      On Error Resume Next
      SetLayeredWindowAttributes frmJMacro.hWnd, vbBlue, i * 5, LWA_ALPHA
      If i = 1 Then
         frmJMacro.Visible = True
         'you must refresh frmmain to avoid blackness
         frmJMacro.Refresh
      End If
   Next i


 
End Sub

