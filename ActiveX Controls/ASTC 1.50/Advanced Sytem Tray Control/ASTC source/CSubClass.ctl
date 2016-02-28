VERSION 5.00
Begin VB.UserControl SubClass 
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   InvisibleAtRuntime=   -1  'True
   Picture         =   "CSubClass.ctx":0000
   ScaleHeight     =   540
   ScaleWidth      =   540
   ToolboxBitmap   =   "CSubClass.ctx":E977
   Begin VB.Image Image1 
      Height          =   540
      Left            =   0
      Picture         =   "CSubClass.ctx":EC89
      Stretch         =   -1  'True
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "SubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event MessageReceived(hwnd As Long, Msg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
Attribute MessageReceived.VB_Description = "Occur when a new message sent to the subclassed Control."

Private ParentHwnd As Long
Private OWindowProc As Long
Private TerminateState As Boolean
 
Dim Bypass_Messages As Boolean

Public Function BeginSubClass(ByVal PHwnd As Long) As Boolean
Attribute BeginSubClass.VB_Description = "Start subclassing a specified control. "
 Dim i As Integer
 Dim CHwnd As Long
 Dim ControlName As String
 Dim SubClassObj2 As SubClass
   
   On Error GoTo ShowErr
       
       'this is very important
       'when the developer make a multi call to BeginSubClass for the same control
       'something like that
       'subclass1.beginsubclass hwnd
       'subclass1.beginSubclass text1.hwnd
       If ParentHwnd <> 0 Then
       'with this check we will overwrite the pervious subclassing
       'and start subclass a new form
          EndSubClass
          '''ErrRaise eeAlreayhandle, Extender.Name: Exit Function
       End If
       If IsWindow(PHwnd) = False Then
          ErrRaise eeInvalidWindow, Extender.Name & "::BeginSubClass"
          Exit Function
       End If
       
       If IsWindowLocal(PHwnd) = False Then
          ErrRaise eeNoExternalWindow, Extender.Name & "::BeginSubClass"
          Exit Function
       End If
       
       If IsAttached(PHwnd) Then
           For i = 1 To SubClassCollection.Count
              Set SubClassObj2 = SubClassCollection(i)
              CHwnd = SubClassObj2.Parent
              If CHwnd = PHwnd Then
                 ControlName = SubClassObj2.MyName
                 Set SubClassObj2 = Nothing
                 Exit For
              End If
           Next i
          ErrRaise eeAlreadyAttached, Extender.Name & "::BeginSubClass", ControlName
          Exit Function
       End If
       
      'to avoid errors if we call show again after we assign the subclass
      If OWindowProc = 0 Then
         ParentHwnd = PHwnd
         OWindowProc = SetWindowLong(ParentHwnd, GWL_WNDPROC, AddressOf SubClassProc)
         If OWindowProc = 0 Then
            ErrRaise eeCantSubClass, Extender.Name
            BeginSubClass = False
            Exit Function
         End If
         IsAttached(ParentHwnd) = True
         
         SubClassCollection.Add Me, str(ParentHwnd)
         OldSubClassProc(ParentHwnd) = OWindowProc
      End If
    
      BeginSubClass = True
          
      
ShowExit:
    On Error Resume Next

    Exit Function

ShowErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & " " & Extender.Name & " :: BeginSubClass" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
'''        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo ShowExit

End Function

Public Sub EndSubClass()
Attribute EndSubClass.VB_Description = "End Subclassing."
Dim i As Integer
Dim CHwnd As Long
Dim res As Long
Dim Attstated As Boolean
Dim SubClassObject As SubClass
 
    Attstated = IsAttached(ParentHwnd)
       
    If Terminateing = False Then
       If Attstated = False Then
        '  ErrRaise eeNotInSystray, Extender.Name & ".Hide"
       End If
    End If
    
    'in case of invokeing hide in usercontrol_terminateing
    'and in the same time there is no icon for this control so we
    If ParentHwnd = 0 Then Exit Sub
    'to avoid unexpected results
    
    Call SetWindowLong(ParentHwnd, GWL_WNDPROC, OWindowProc)
               
    'we check Terminateing to avoid msgbox Error in OldWinProc and IsAttached
    'this error comes from setprop Function
    If Terminateing = False Then
       OldSubClassProc(ParentHwnd) = 0
       IsAttached(ParentHwnd) = False
    End If
    
    OWindowProc = 0
    
    For i = 1 To SubClassCollection.Count
        
        Set SubClassObject = SubClassCollection(i)
        CHwnd = SubClassObject.Parent
        If CHwnd = ParentHwnd Then
           SubClassCollection.Remove (i)
           Exit For
        End If
          
    Next i
  
    ParentHwnd = 0
     
   
End Sub

Private Property Get Terminateing() As Boolean
Terminateing = TerminateState
End Property

Private Property Let Terminateing(ByVal vNewValue As Boolean)
TerminateState = vNewValue
End Property

Friend Function Parent() As Long
Parent = ParentHwnd
End Function

Friend Function MyName() As Variant
MyName = Extender.Name
End Function

Private Sub UserControl_Resize()
UserControl.Width = 540
UserControl.Height = 540
End Sub

Private Sub UserControl_Terminate()

Terminateing = True

EndSubClass

SubClasser.clearme

Set SubClassCollection = Nothing

End Sub

Friend Sub RaiseEvents(mhWnd As Long, Msg As Long, wParam As Long, lParam As Long)
Dim Cancel As Boolean
EndSubClass
RaiseEvent MessageReceived(mhWnd, Msg, wParam, lParam, Cancel)
CancelMessages = Cancel
BeginSubClass mhWnd
End Sub

Friend Property Get CancelMessages() As Boolean
CancelMessages = Bypass_Messages
End Property

Friend Property Let CancelMessages(ByVal vNewValue As Boolean)
Bypass_Messages = vNewValue
End Property
