Attribute VB_Name = "SubClasser"
Option Explicit
Dim SubClassObj As SubClass
Private OldWindowProc As Long

Public Function SubClassProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim res As Long

Const WM_NCDESTROY = &H82

OldWindowProc = OldSubClassProc(hwnd)

'
Set SubClassObj = SubClassCollection(str(hwnd))

SubClassObj.RaiseEvents hwnd, Msg, wParam, lParam
'
'bypass the remaing messages for this one only
If SubClassObj.CancelMessages = True Then
   SubClassObj.CancelMessages = False
   Exit Function
End If

If Msg = WM_NCDESTROY Then
'
   res = SetWindowLong(hwnd, GWL_WNDPROC, OldWindowProc)
   If res = 0 Then LastErrorForMe = "Error Unsubclass"
'
End If
'
SubClassProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)

End Function

Public Property Get OldSubClassProc(ByVal hwnd As Long) As Long
Dim sName As String
   sName = hwnd
   OldSubClassProc = GetProp(hwnd, sName)
End Property
Public Property Let OldSubClassProc(ByVal hwnd As Long, ByVal lPtr As Long)
Dim m_f As Long
Dim sName As String

   m_f = 1
   sName = hwnd
   m_f = SetProp(hwnd, sName, lPtr)
   If m_f = 0 Then LastErrorForMe = "Error in Let OldWinProc"
   If (lPtr = 0) Then
      RemoveProp hwnd, sName
   End If
End Property

Public Sub clearme()
Set SubClassObj = Nothing
End Sub


