Attribute VB_Name = "API_CALL2"
Option Explicit

Public Const OF_READ = &H0
Public Const OF_WRITE = &H1
Public Const OF_CREATE = &H1000
Public Const OFS_MAXPATHNAME = 128

Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function LogOn() As Boolean
   
End Function


Public Sub ComposeMessage()
  
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

