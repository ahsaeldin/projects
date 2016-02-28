Attribute VB_Name = "ModWindowInfo"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    right As Long
    Bottom As Long
End Type

Public Type WinInfo
         
          
         Top As Long
         
         Left As Long
         
         right As Long
         
         Bottom As Long

         Hwnd As Long
 
         Title As String * 70

         UniqueWindow As Boolean
  
         Class_Name As String * 70
         
    
End Type

Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal Hwnd As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal Hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


