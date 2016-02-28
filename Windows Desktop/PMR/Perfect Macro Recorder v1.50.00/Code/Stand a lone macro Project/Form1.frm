VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private mSpeed As String
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'first checking
'sec FreeFile
Private Sub Form_Load()
         
         Dim mePath As String
         Dim TipState As String
         Dim marker As Variant
         Dim TempPath As String
         Dim FileData As String
         Dim MacroData As String
          
         If Command$ = "/h" Then
            mSpeed = "h"
         ElseIf Command$ = "/n" Then
            mSpeed = "n"
         ElseIf Command$ = "/l" Then
            mSpeed = "l"
         End If
         
       '  Unload frmAbout
         
'         TipState = GetSetting(App.EXEName, "k", "d", "")
    
'         If Not TipState = "one" Then
'            frmTip.Show 1
'         End If
         
         TempPath = GetTempPathName
         
         mePath = AppPath & App.EXEName & ".exe"
         
         Open mePath For Binary As #2
        
           FileData = Input(LOF(2), 2)
        
         Close #2
         
         marker = InStr(1, FileData, "cprinsosonono")
         
         MacroData = Mid(FileData, marker + 13, Len(FileData) - (marker + 12))
         
         Open TempPath & "a.jmf" For Binary As #4
        
           Put #4, , MacroData
        
         Close #4
         
         Call OpenMacro(TempPath & "a.jmf")

         StartPlay
         
         WatchForCancel
    On Error Resume Next
    DeleteFile TempPath & "a.jmf"
    On Error GoTo 0
'MsgBox "PlayBack Over", vbInformation, "Smack Macro Recorder"
        
End

End Sub


Public Function GetTempPathName() As String

    Dim sBuffer As String
    Dim lRet As Long
    
    sBuffer = String$(255, vbNullChar)
    
    lRet = GetTempPath(255, sBuffer)
    
    If lRet > 0 Then
        
        sBuffer = Left$(sBuffer, lRet)
    
    End If
    
    GetTempPathName = sBuffer
    
End Function

Public Function AppPath()

'use the function to get the path of the parameter filename
If right(App.Path, 1) <> "\" Then AppPath = App.Path + "\" Else AppPath = App.Path
   
End Function

Public Property Get Speed() As String
Speed = mSpeed
End Property

Public Property Let Speed(ByVal vNewValue As String)
mSpeed = vNewValue
End Property

Private Sub Timer1_Timer()
If Not IsPlayingMode Then
   
   End

End If
End Sub
