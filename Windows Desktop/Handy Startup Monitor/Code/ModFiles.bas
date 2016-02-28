Attribute VB_Name = "ModFiles"
                                         
                                            '//ÇáÇÑÏÇÉ æÇáÇÏÇÑÉ æÌåÇä áÚãáÉ æÇÍÏÉ
                                            
                        '|$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$|
                        '|"ÇáÝÑÞ Èíä ÇáäÇÌÍ æÈíä ÛíÑå áíÓ äÞÕ ÇáÞæÉ Ãæ ÇáãÚÑÝÉ , Èá ÈÇáÃÍÑì äÞÕ Ýí ÌÇäÈ ÇáÅÑÇÏÉ"|
                        '|ÝíäÓ áæãÈÇÑÏí                                                                  |
                        '|$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Option Explicit

'//For FileInfo Function
'/////////////////////////////////////////////////////////////////////////
Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
   "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal _
   dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias _
   "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long
Declare Function VerQueryValue Lib "Version.dll" Alias _
   "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
   lplpBuffer As Any, puLen As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias _
   "GetSystemDirectoryA" (ByVal path As String, ByVal cbBytes As _
   Long) As Long
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    dest As Any, ByVal Source As Long, ByVal Length As Long)
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, ByVal lpString2 As Long) As Long
    
Public Type FILEPROPERTIE
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
End Type

'//built it using an allapi example called FindFiles

Public Const MAX_PATH = 260

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Function StripNulls(OriginalStr As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

'@@@@@@@@@@@@@UnderStand Every Aspect of this one
'//Veeeeeeeeeeeeryyyyyy important
'//use This Function for searching a dir that contains be max 80 files. i mean the total number of files including
'//sub dir. coz if you use it will take sometime and may "i said may" be hang up
'//Note that i tried to use it know for 15000 files and it runs will and took about 20 sec.
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Function FindFilesAPI(path As String, SearchStr As String, FileCount As Integer, DirCount As Integer, FilesPaths() As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    'KPD-Team 1999
    'E-Mail: KPDTeam@Allapi.net
    'URL: http://www.allapi.net/

    '//An example to use this function
    'Dim SearchPath As String, FindStr As String
    'Dim FileSize As Long
    'Dim NumFiles As Integer, NumDirs As Integer
    'Screen.MousePointer = vbHourglass
    'List1.Clear
    'SearchPath = Text1.Text
    'FindStr = Text2.Text
    'If Text1 = "" Then Exit Sub
    'If Text2 = "" Then Exit Sub
    'Dim ahmed() As String
    'FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs, ahmed)
    'Dim i As Integer
    'For i = 0 To UBound(ahmed) - 1
    '     List1.AddItem ahmed(i)
    'Next
    'Text3.Text = NumFiles & " Files found in " & NumDirs + 1 & " Directories"
    'Text4.Text = "Size of files found under " & SearchPath & " = " & Format(FileSize, "#,###,###,##0") & " Bytes"
    
    Dim FileName As String ' Walking filename variable...
    Dim DirName As String ' SubDirectory Name
    Dim dirNames() As String ' Buffer for directory name entries
    Dim nDir As Integer ' Number of directories in this path
    Dim I As Integer ' For-loop counter...
    Dim hSearch As Long ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DirName = StripNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
                FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
                '//FilePaths is a multidimensional array where "2" is the column number and FileCount is the Rows number
                ReDim Preserve FilesPaths(2, FileCount) As String
                FilesPaths(0, FileCount - 1) = path & FileName
                FilesPaths(1, FileCount - 1) = FileName
                'Debug.Print FileName, path & FileName
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    ' If there are sub-directories...
    If nDir > 0 Then
        ' Recursively walk into them...
        For I = 0 To nDir - 1
            FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(I) & "\", SearchStr, FileCount, DirCount, FilesPaths)
        Next I
    End If
End Function

' Return file-properties of given file  (EXE , DLL , OCX)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function FileInfo(Optional ByVal PathWithFilename As String) As FILEPROPERTIE
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'//pt comment this code chunk coz i replace it with a public function "ExtractPath"
'Dim WhereSlash As Integer
'Dim WhereSecondQuoteion As Integer
'If Left(PathWithFilename, 1) = Chr(34) Then
'   WhereSecondQuoteion = InStr(2, PathWithFilename, Chr(34))
'   If WhereSecondQuoteion Then
'      PathWithFilename = Mid(PathWithFilename, 2, WhereSecondQuoteion - 2)
'   End If
'ElseIf InStr(2, PathWithFilename, "/") Then
'       WhereSlash = InStr(2, PathWithFilename, "/")
'       PathWithFilename = Left(PathWithFilename, WhereSlash - 2)
'End If

PathWithFilename = ExtractPath(PathWithFilename)

Static BACKUP As FILEPROPERTIE   ' backup info for next call without filename

If Len(PathWithFilename) = 0 Then
   FileInfo = BACKUP
   Exit Function
End If

Dim lngBufferlen As Long
Dim lngDummy As Long
Dim lngRc As Long
Dim lngVerPointer As Long
Dim lngHexNumber As Long
Dim bytBuffer() As Byte
Dim bytBuff(255) As Byte
Dim strBuffer As String
Dim strLangCharset As String
Dim strVersionInfo(7) As String
Dim strTemp As String
Dim intTemp As Integer
       
' size
lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
If lngBufferlen > 0 Then
   ReDim bytBuffer(lngBufferlen)
   lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
   If lngRc <> 0 Then
      lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
      If lngRc <> 0 Then
         'lngVerPointer is a pointer to four 4 bytes of Hex number,
         'first two bytes are language id, and last two bytes are code
         'page. However, strLangCharset needs a  string of
         '4 hex digits, the first two characters correspond to the
         'language id and last two the last two character correspond
         'to the code page id.
         MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
         lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + _
                bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
         strLangCharset = Hex(lngHexNumber)
         'now we change the order of the language id and code page
         'and convert it into a string representation.
         'For example, it may look like 040904E4
         'Or to pull it all apart:
         '04------        = SUBLANG_ENGLISH_USA
         '--09----        = LANG_ENGLISH
         ' ----04E4 = 1252 = Codepage for Windows:Multilingual
         Do While Len(strLangCharset) < 8
             strLangCharset = "0" & strLangCharset
         Loop
         ' assign propertienames
         strVersionInfo(0) = "CompanyName"
         strVersionInfo(1) = "FileDescription"
         strVersionInfo(2) = "FileVersion"
         strVersionInfo(3) = "InternalName"
         strVersionInfo(4) = "LegalCopyright"
         strVersionInfo(5) = "OriginalFileName"
         strVersionInfo(6) = "ProductName"
         strVersionInfo(7) = "ProductVersion"
         ' loop and get fileproperties
         For intTemp = 0 To 7
            strBuffer = String$(255, 0)
            strTemp = "\StringFileInfo\" & strLangCharset _
               & "\" & strVersionInfo(intTemp)
            lngRc = VerQueryValue(bytBuffer(0), strTemp, _
                  lngVerPointer, lngBufferlen)
            If lngRc <> 0 Then
               ' get and format data
               lstrcpy strBuffer, lngVerPointer
               strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
               strVersionInfo(intTemp) = strBuffer
             Else
               ' property not found
               strVersionInfo(intTemp) = "?"
            End If
         Next intTemp
      End If
   End If
End If
' assign array to user-defined-type
FileInfo.CompanyName = strVersionInfo(0)
FileInfo.FileDescription = strVersionInfo(1)
FileInfo.FileVersion = strVersionInfo(2)
FileInfo.InternalName = strVersionInfo(3)
FileInfo.LegalCopyright = strVersionInfo(4)
FileInfo.OrigionalFileName = strVersionInfo(5)
FileInfo.ProductName = strVersionInfo(6)
FileInfo.ProductVersion = strVersionInfo(7)
BACKUP = FileInfo
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Function ResolveShortcutPath(ByVal FileName As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim Obj As Object
    
    Set Obj = CreateObject("WScript.Shell")
    
    Dim Shortcut As Object
    
    Set Shortcut = Obj.CreateShortcut(FileName)
    
    ResolveShortcutPath = Shortcut.TargetPath
     
    Set Shortcut = Nothing
    
    Set Obj = Nothing

End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ExtractPath(ByVal PathWithFilename As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim WhereSlash As Integer
    Dim WhereToStart As Integer
    Dim WhereSecondQuoteion As Integer
 
    If Left(PathWithFilename, 1) = Chr(34) Then
       '//the idea behind "WhereToStart" is to handle cases where paths like this
       '// ""c:\program.exe""
       WhereToStart = 2
       If Mid(PathWithFilename, 2, 1) = Chr(34) Then
          WhereToStart = 3
       End If
       WhereSecondQuoteion = InStr(WhereToStart, PathWithFilename, Chr(34))
       If WhereSecondQuoteion Then
          PathWithFilename = Mid(PathWithFilename, WhereToStart, WhereSecondQuoteion - WhereToStart)
       End If
    ElseIf InStr(2, PathWithFilename, "/") Then
           WhereSlash = InStr(2, PathWithFilename, "/")
           PathWithFilename = Left(PathWithFilename, WhereSlash - 2)
    Else '//the idea behind this case is to extract a path like this:
         '//C:\Program Files\Mozilla Firefox\plugins\NPSWF32_FlashUtil.exe -p
           Dim WhereDot As Integer
           WhereDot = InStrRev(PathWithFilename, ".") '//where is the dot before exe
           If WhereDot Then
              Dim Extension As String
              Extension = Mid(PathWithFilename, WhereDot + 1, 3) '//extract the extension to ensure it is exe
              If Len(PathWithFilename) - (WhereDot + 3) > 0 Then '//if---then it means that path have a command line like above
                 If Extension = "exe" Then '//handle exe file case only 'may be handle all cases in future
                    PathWithFilename = Left(PathWithFilename, WhereDot + 3)
                 End If
                 '//pt 15Apr2007 : added to handle a case like this "RUNDLL32.EXE C:\\WINDOWS\\system32\\NvCpl.dll,NvStartup"
                 If Extension = "dll" Then '//handle dll file case only 'may be handle all cases in future
                    PathWithFilename = Left(PathWithFilename, WhereDot + 3)
                    If InStr(1, PathWithFilename, "RUNDLL32.EXE") Then
                       PathWithFilename = Mid(PathWithFilename, 14, Len(PathWithFilename) - 13)
                    End If
                 End If
              End If
           End If
           
           '40
    End If

    ExtractPath = PathWithFilename

End Function

