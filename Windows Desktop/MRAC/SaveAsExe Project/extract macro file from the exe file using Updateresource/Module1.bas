Attribute VB_Name = "Module1"
Option Explicit
Public Type SAFEARRAYBOUND
   cElements As Long      'number of elements of a dimension
   lLbound As Long        'lower bounds of the array dimension
End Type
'8 bytes

Public Type SAFEARRAY
   cDims As Integer       'number of dimensions in this array
   fFeatures As Integer   'flags used by the SAFEARRAY routines
   cbElements As Long     'size of an element
   cLocks As Long         'number of times the array has been locked without corresponding unlock
   pvData As Long         'pointer to the data.
   rgsabound(1 To 60) As SAFEARRAYBOUND   'one bound for each dimension.
'an array can have max 60 dimensions, only the first cDims items will be used
'note that rgsabound elements are in reverse order, e.g. for a 2-dimensional
'array, rgsabound(1) holds info about columns, and rgsabound(2) about rows
End Type


Public Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" _
(ByVal pFileNaME As String, ByVal bDeleteExistingResouces As Boolean) As Long

Public Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" _
(ByVal HResFile As Long, ByVal bDeleteExistingResouces As Boolean) As Long

Public Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" _
(hUpdate As Long, lpType As Long, lpName As String, wLanguage As Long, _
lpData As Variant, cbData As Long) As Long

Public Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function FindResourceByNum Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Long) As Long
Public Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Public Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Public Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long

Public hModule As Long
Dim resoop As Long
Private m_tSA As SAFEARRAY

Private m_lSAFEARRAY As Long
Function ResTypesCallBack(ByVal hMod As Long, ByVal ResType As Long, ByVal lParam As Long) As Long

   Dim Data As Variant
     If (ResType And &HFFFF0000) = 0 Then
     
     Else    ' The resource type is defined by a string
       '  SaveSetting App.EXEName, "a", "a", ResType
             
         resoop = ResType
           Data = GetDataArray(ResType, "")
     End If
  '  Set nd = Nothing
    ResTypesCallBack = True
End Function

Public Function InitResource(ByVal sLibName As String) As Boolean
  On Error Resume Next
    Dim ret As Long
     Dim Data As Variant
   hModule = LoadLibraryEx(sLibName, 0, 1)
   If hModule Then ret = EnumResourceTypes(hModule, AddressOf ResTypesCallBack, 0)
  
     
    InitResource = (hModule <> 0)
End Function
Public Function GetDataArray(ByVal ResType As String, ByVal ResName As String) As Variant
   Dim hRsrc As Long
   Dim hGlobal As Long
   Dim arrData() As Byte
   Dim marrData As String
   Dim lpData As Long
   Dim arrSize As Long
   If IsNumeric(ResType) Then hRsrc = FindResourceByNum(hModule, "#101", CLng(ResType))
   If hRsrc = 0 Then hRsrc = FindResource(hModule, "#101", ResType)
   If hRsrc = 0 Then Exit Function
   hGlobal = LoadResource(hModule, hRsrc)
   lpData = LockResource(hGlobal)
   arrSize = SizeofResource(hModule, hRsrc)
   If arrSize = 0 Then Exit Function
   ReDim arrData(arrSize - 1)
   Call CopyMemory(arrData(0), ByVal lpData, arrSize)
   Call FreeResource(hGlobal)
   GetDataArray = arrData
   
   Const path2 = "H:\Documents and Settings\AHMED.CPRIN\Desktop\hacked.jmf"
   Const path = "H:\Documents and Settings\AHMED.CPRIN\Desktop\ahmed.exe"
   Dim Begin As Long
   Dim resTT As Long
'   Begin = BeginUpdateResource(path, True)
  
     Open path2 For Input As #3
     
               Input #3, marrData
  
     Close #3
    Dim kk As String
    Open path For Binary As #4
       Get #4, 1, kk
    Close #4
    
' resTT = UpdateResource(Begin, CLng(ResType), "#101", 3073, m_lSAFEARRAY, FileLen(path2) - 1)
   
'  Dim k
'  k = EndUpdateResource(Begin, False)
   
End Function

