Attribute VB_Name = "Utils"

                    '//"·«  —ﬂ“ ⁄·Ï «·„«÷Ï , ›ﬁÿ «” ⁄„·Â ··Œ—ÊÃ »«·Õﬂ„… Ê«·›«∆œ… , À„ « —ﬂÂ Œ·›ﬂ

Option Explicit
   
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ExtractFileNameFromPath(path As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
'Print ExtractFileNameFromPath("C:\WINDOWS\system32\ctfmon.exe")
'Returns ctfmon.exe

'Note : Must call "ExtractPath" function before calling this if the path may contains quotaions

   Dim WhereBackSlash As Integer
   WhereBackSlash = InStrRev(path, "\")
   ExtractFileNameFromPath = Mid(path, WhereBackSlash + 1, Len(path) - WhereBackSlash)
   
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ExtractFileExtensionFromPath(path As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'Print ExtractFileLocationFromPath("C:\WINDOWS\system32\ctfmon.exe")
'Returns exe
Dim ExtensionCharsNum As Integer

'Note : Must call "ExtractPath" function before calling this if the path may contains quotaions
   Dim WhereBackSlash As Integer
   WhereBackSlash = InStrRev(path, ".")
   'ExtractFileNameFromPath = Mid(path, WhereBackSlash + 1, Len(path) - WhereBackSlash)
   ExtensionCharsNum = Len(path) - WhereBackSlash
   ExtractFileExtensionFromPath = Mid(path, WhereBackSlash + 1, ExtensionCharsNum) '(path, WhereBackSlash + 1, Len(path) - WhereBackSlash)
      
End Function
