Attribute VB_Name = "Registre"

                        '|$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$|
                        '|"íÓÊáÒã ÇáÃãÑ 20 ÓäÉ áÈäÇÁ ÓãÚÉ ØíÈÉ , æÎãÓ ÏÞÇÆÞ áÊÏãíÑ åÐå ÇáÓãÚÉÒ ÅÐÇ ÝßÑÊ ãä åÐÇ ÇáãäÙæÑ , ÝÓÊÄÏì ÃÔíÇÁ ßËíÑÉ ÈÔßá ãÎÊáÝ"|
                        '|æÇÑä ÈÇÝíÊ                                                                                                            |
                        '|$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$|

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&


Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1 ' Unicode nul terminated String
    Public Const REG_DWORD = 4 ' 32-bit number
    
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub savekey(hKey As Long, strPath As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
  
    Dim keyhand&
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function getstring(hKey As Long, strPath As String, strValue As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                getstring = Left$(strBuf, intZeroPos - 1)
            Else
                getstring = strBuf
            End If
        End If
    End If
End Function


'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub Savestring(hKey As Long, strPath As String, strValue As String, strdata As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub


'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Function getdword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    'EXAMPLE:
    '
    'text1.text = getdword(HKEY_CURRENT_USER
    '     , "Software\VBW\Registry", "Dword")
    '
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    ' Get length/data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)


    If lResult = ERROR_SUCCESS Then


        If lValueType = REG_DWORD Then
            getdword = lBuf
        End If
        'Else
        'Call errlog("GetDWORD-" & strPath, Fals
        '     e)
    End If
    r = RegCloseKey(keyhand)
End Function


'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    'EXAMPLE"
    '
    'Call SaveDword(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry", "Dword", text1.text)
    '
    '
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then
    '     Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function


'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW")
    '
    Dim r As Long
    r = RegDeleteKey(hKey, strKey)
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "MSMSGS")
    '
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function EnumValues(Root As Long, KeyPath As String) As Long
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Const BUFFER_SIZE As Long = 255
    
    Dim KPath As String
    Dim CompanyName As String
    Dim FileDescription As String
    Dim hKey As Long, cnt As Long, sName As String, sData As String, Ret As Long, RetData As Long
    
    cnt = 0
    
    With frmmain.lstStartupPrograms.ListItems
    If RegOpenKey(Root, KeyPath, hKey) = 0 Then
        'initialize
        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
        Ret = BUFFER_SIZE
        RetData = BUFFER_SIZE
        
        'enumerate the values
        While RegEnumValue(hKey, cnt, sName, Ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            'show data
            If RetData > 0 Then
               
               .Add , , Left$(sName, Ret)
               .Item(.Count).SubItems(1) = Left$(sData, RetData - 1)
               CompanyName = FileInfo(.Item(.Count).SubItems(1)).CompanyName
               FileDescription = FileInfo(.Item(.Count).SubItems(1)).FileDescription
               If FileDescription = "" Then
                  FileDescription = .Item(.Count).Text
                  .Item(.Count).SubItems(2) = FileDescription
               Else
                  .Item(.Count).SubItems(2) = FileDescription & " by " & CompanyName
               End If
              
              .Item(.Count).ToolTipText = ExtractPath(Left$(sData, RetData - 1))
              
              AddIconFileToList .Item(.Count).SubItems(1)
               
               Select Case Root
                      
                      Case HKEY_USERS
                      
                           KPath = "HKCR\" & KeyPath
                      
                      Case HKEY_CURRENT_USER
                      
                           KPath = "HKCU\" & KeyPath
                    
                      Case HKEY_CLASSES_ROOT
                           
                           KPath = "KHCR\" & KeyPath
                           
                      Case HKEY_LOCAL_MACHINE
                      
                           KPath = "HKLM\" & KeyPath
                      
               End Select
               
               .Item(.Count).SubItems(3) = KPath
               
        End If
            
            'prepare for next value
            cnt = cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            Ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        'Close the registry key
        RegCloseKey hKey
        
        EnumValues = 1
        
    Else
    
        EnumValues = 0
    
    End If
    
    End With
    
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function IsRegValueStillThere(Root As Long, KeyPath As String, ValueTitle As String, Optional Value As String) As Boolean
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Const BUFFER_SIZE As Long = 255
    
    Dim KPath As String
    Dim hKey As Long, cnt As Long, sName As String, sData As String, Ret As Long, RetData As Long
    
    cnt = 0
    
    If RegOpenKey(Root, KeyPath, hKey) = 0 Then
            'initialize
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            Ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
            
            'enumerate the values
            While RegEnumValue(hKey, cnt, sName, Ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
                'show data
                If RetData > 0 Then
                   
                   Value = Left$(sData, RetData - 1)
                   
                   Dim ValueName As String
                   ValueName = Left$(sName, Ret)
                   If ValueName = ValueTitle Then
                      RegCloseKey hKey
                      IsRegValueStillThere = True
                      Exit Function
                   End If
                   
                End If
                
                'prepare for next ValueTitle
                cnt = cnt + 1
                sName = Space(BUFFER_SIZE)
                sData = Space(BUFFER_SIZE)
                Ret = BUFFER_SIZE
                RetData = BUFFER_SIZE
            
            Wend
            
            'Close the registry key
            RegCloseKey hKey
            
    Else
    
            IsRegValueStillThere = False
    
    End If
    
    IsRegValueStillThere = False

End Function
