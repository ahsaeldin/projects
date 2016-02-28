Attribute VB_Name = "ModCore"
                        
                        '|$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$|
                        '|"„« Ì Œ·Â ÊÌ’œﬁÂ «·„—¡ „‰« , ›≈‰Â ﬁ«œ— ⁄·Ï  ÕﬁÌﬁÂ"|
                        '|«»·ÌÊ‰ ÂÌ·                                      |
                        '|$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$|

Option Explicit

Private Type SHFILEINFO
    hicon As Long                      '  out: icon
    iIcon As Long          '  out: icon index'
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80         '  out: type name
End Type

Public Enum StartUpType

    AllUser = 0
    CurrentUser = 1

End Enum

Private Declare Function DuplicateIcon Lib "shell32.dll" (ByVal hInst As Long, ByVal hicon As Long) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CheckList()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim I As Integer
    Dim Root As Long
    Dim Key As String
    Dim FileSize As Long
    Dim RegItemPath As String
    Dim FullRegKeyPath As String
    Dim StartupFolderPath As String
    Dim IsStartUpFolders As Boolean
    Dim ValueIsStillThere As Boolean
    Dim StartUpItemsArray() As String
    Dim NumFiles As Integer, NumDirs As Integer
                    
    'If Not rMonitorAllowed Then Exit Sub
    
    'If Not rchkRemoved Then GoTo SkipAndCheckAdded:
    
    Debug.Print Rnd * 23452345
    
    With frmmain.lstStartupPrograms.ListItems
    
    For I = 1 To .Count
           
        FullRegKeyPath = .Item(I).SubItems(3)

        Key = Mid(FullRegKeyPath, 6, Len(FullRegKeyPath) - 5)
        
        Select Case Left(FullRegKeyPath, 4)
               
               Case "HKCU"
                    
                    Root = HKEY_CURRENT_USER
                    
                    ValueIsStillThere = IsRegValueStillThere(Root, Key, .Item(I), RegItemPath)
                    
                    If Not ValueIsStillThere Then '//if ValueIsStillThere is not there
                       ShowRemovedAlert .Item(I)
                       .Remove (I) '//then remove this item
                       Exit For
                    Else '//if ValueIsStillThere then check if the path changed
                       If .Item(I).SubItems(1) <> RegItemPath Then
                          .Item(I).SubItems(1) = RegItemPath
                       End If
                    End If
                    
               Case "HKLM"
                    
                    Root = HKEY_LOCAL_MACHINE
                    
                    ValueIsStillThere = IsRegValueStillThere(Root, Key, .Item(I), RegItemPath)
                    
                    If Not ValueIsStillThere Then '//if ValueIsStillThere is not there
                       ShowRemovedAlert .Item(I)
                       .Remove (I) '//then remove this item
                       Exit For
                    Else '//if ValueIsStillThere then check if the path changed
                       If .Item(I).SubItems(1) <> RegItemPath Then
                          .Item(I).SubItems(1) = RegItemPath
                       End If
                    End If
               
               Case Else
                    
                    NumDirs = 0
                    NumFiles = 0
        
                    If Left(FullRegKeyPath, 3) = "All" Then
                       
                       FileSize = FindFilesAPI(AllUStartupFolderPath, .Item(I).Tag, NumFiles, NumDirs, StartUpItemsArray)
                       
                       If NumFiles = 0 Then '//File Not Found
                          '//„·ÕÊŸ… «‰« „‘ Â«⁄„· „—«ﬁ»… ·ﬁÌ„… «·‘Ê—  ﬂ  «·„ „À·… ›Ï „”«— «·»—‰«„Ã Ê»«· «·Ï
                          '//·Ê « €Ì— „”«— »—‰«„Ã «·‘Ê—  ﬂ  «·»—‰«„Ã „‘ ÂÌ⁄—› .. ⁄·Ï ⁄ﬂ” „« »ÌÕ’· ›Êﬁ
                          ShowRemovedAlert ExtractFileNameFromPath(.Item(I).SubItems(1))
                          .Remove (I)
                          Exit For
                       End If
                          
                    ElseIf Left(FullRegKeyPath, 7) = "Current" Then
                    
                       FileSize = FindFilesAPI(CUStartupFolderPath, .Item(I).Tag, NumFiles, NumDirs, StartUpItemsArray)
                       
                       If NumFiles = 0 Then  '//File Not Found
                          '//„·ÕÊŸ… «‰« „‘ Â«⁄„· „—«ﬁ»… ·ﬁÌ„… «·‘Ê—  ﬂ  «·„ „À·… ›Ï „”«— «·»—‰«„Ã Ê»«· «·Ï
                          '//·Ê « €Ì— „”«— »—‰«„Ã «·‘Ê—  ﬂ  «·»—‰«„Ã „‘ ÂÌ⁄—› .. ⁄·Ï ⁄ﬂ” „« »ÌÕ’· ›Êﬁ
                          ShowRemovedAlert ExtractFileNameFromPath(.Item(I).SubItems(1))
                          .Remove (I)
                          Exit For
                       End If
                     
                    End If
               
        End Select
        
    Next I
    
    End With
     
'SkipAndCheckAdded:

    'If Not rchkNewlyAdded Then Exit Sub
    'If Not rMonitorAllowed Then Exit Sub
    
    Dim Ret As Long
    
    'Current User, Run
    Ret = CheckRegForNewItems(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
 
    'Current User, RunOnce
    Ret = CheckRegForNewItems(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
 
    'All Users, Run
    Ret = CheckRegForNewItems(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
 
    'All Users, RunOnce
    Ret = CheckRegForNewItems(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
 
    'All Users, RunOnceEx
    Ret = CheckRegForNewItems(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx")
 
    CheckStartUpFoldersItems AllUStartupFolderPath, AllUser
      
    CheckStartUpFoldersItems CUStartupFolderPath, CurrentUser

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub FillList()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
       
      Dim Ret  As Long
       
      frmmain.lstStartupPrograms.ListItems.Clear
        
      'Current User, Run
      Ret = EnumValues(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
      
      'Current User, RunOnce
      Ret = EnumValues(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
      
      'All Users, Run
      Ret = EnumValues(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
      
      'All Users, RunOnce
      Ret = EnumValues(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
      
      'All Users, RunOnceEx
      Ret = EnumValues(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx")
      
      Err.Description = "Can't Get StartUp Items Data."
      
      If Ret = 0 Then SysError Err, "0x0f4hsm111"
      
      Dim CUStartUpFolderItems() As String
      
      Dim AllUStartUpFolderItems() As String
      
      AddStartUpItems CUStartupFolderPath, CUStartUpFolderItems, CurrentUser
      
      AddStartUpItems AllUStartupFolderPath, AllUStartUpFolderItems, AllUser
       
      Erase CUStartUpFolderItems
      
      Erase AllUStartUpFolderItems
      
'//pt 18Feb2008 6:18 pm : to avoid mess and duplicating codes i commented the next code chunk and replace it by a new function below called "AddStartUpItems"
''''''''//pt 17Feb2008 06:32 pm : Add StartUp Folder Items''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Dim i As Integer
'      Dim FileSize As Long
'      Dim CUStartUpFolderItems() As String
'      Dim AllUStartUpFolderItems() As String
'      Dim NumFiles As Integer, NumDirs As Integer
'
'      FileSize = FindFilesAPI(CUStartupFolderPath, "*.*", NumFiles, NumDirs, CUStartUpFolderItems)
'
'      With frmmain.lstStartupPrograms.ListItems
'
'          For i = 0 To UBound(CUStartUpFolderItems, 2) - 1
'
'              '//No Need To Display "desktop.ini" file
'              If CUStartUpFolderItems(1, i) <> "desktop.ini" Then
'                 If CUStartUpFolderItems(1, i) <> "" Then
'                    'Debug.Print CUStartUpFolderItems(1, i)
'                    .Add , , CUStartUpFolderItems(1, i)
'                    .Item(.Count).SubItems(3) = "Current User Startup" 'CUStartUpFolderItems(0, i)
'                    .Item(.Count).SubItems(1) = ResolveShortcutPath(CUStartUpFolderItems(0, i))
'                    Dim FileDescription As String
'                    FileDescription = FileInfo(.Item(.Count).SubItems(1)).FileDescription
'                    If FileDescription <> "" Then
'                       .Item(.Count).SubItems(2) = FileDescription & " by " & FileInfo(.Item(.Count).SubItems(1)).CompanyName
'                    End If
'                 End If
'              End If
'
'          Next
'
'          NumDirs = 0
'          NumFiles = 0
'
'          FileSize = FindFilesAPI(AllUStartupFolderPath, "*.*", NumFiles, NumDirs, AllUStartUpFolderItems)
'
'          For i = 0 To UBound(AllUStartUpFolderItems, 2) - 1
'
'              '//No Need To Display "desktop.ini" file
'              If AllUStartUpFolderItems(1, i) <> "desktop.ini" Then
'              'If AllUStartUpFolderItems(1, i) <> "" Then
'                    'Debug.Print AllUStartUpFolderItems(0, i)
'                    .Add , , AllUStartUpFolderItems(1, i)
'                    .Item(.Count).SubItems(3) = "All Users Startup"   ' AllUStartUpFolderItems(0, i)
'                    .Item(.Count).SubItems(1) = ResolveShortcutPath(AllUStartUpFolderItems(0, i))
'                    FileDescription = ""
'                    FileDescription = FileInfo(.Item(.Count).SubItems(1)).FileDescription
'                    If FileDescription <> "" Then
'                       .Item(.Count).SubItems(2) = FileDescription & " by " & FileInfo(.Item(.Count).SubItems(1)).CompanyName
'                    End If
'              'End If
'              End If
'
'          Next
'
'      End With
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddStartUpItems(StartupFolderPath As String, StartUpItemsArray() As String, StartUpFolderType As StartUpType)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

      '//pt 17Feb2008 06:32 pm : Add StartUp Folder Items//////////////////////////////////////////////////////////
      Dim I As Integer
      Dim FileSize As Long
    '  Dim CUStartUpFolderItems() As String
    '  Dim StartUpItemsArray() As String
      Dim NumFiles As Integer, NumDirs As Integer
      
      Dim FileExtension As String
      
      NumDirs = 0
      NumFiles = 0
          
      FileSize = FindFilesAPI(StartupFolderPath, "*.*", NumFiles, NumDirs, StartUpItemsArray)
      
      With frmmain.lstStartupPrograms.ListItems
      
          For I = 0 To UBound(StartUpItemsArray, 2) - 1
              
              '//No Need To Display "desktop.ini" file
              If StartUpItemsArray(1, I) <> "desktop.ini" Then
                 If StartUpItemsArray(1, I) <> "" Then  'And InStr(1, StartUpItemsArray(1, I), ".")
                    'Debug.Print StartUpItemsArray(1, i)
                    If InStr(1, StartUpItemsArray(1, I), ".") Then '//if the file has extension then
                       Dim FileExtensionCharsNum As Integer
                       FileExtensionCharsNum = Len(ExtractFileExtensionFromPath(StartUpItemsArray(1, I)))
                       .Add , , Left$(StartUpItemsArray(1, I), Len(StartUpItemsArray(1, I)) - (FileExtensionCharsNum + 1))
                    Else
                       .Add , , StartUpItemsArray(1, I)
                    End If
                    '.Add , , Left$(StartUpItemsArray(1, I), Len(StartUpItemsArray(1, I)) - (FileExtensionCharsNum + 1))
                    .Item(.Count).SubItems(3) = IIf(StartUpFolderType, "Current User Startup", "All Users Startup")
                    '//Save the shortcut file name in the tag property in order to use in "CheckList" function
                    .Item(.Count).Tag = StartUpItemsArray(1, I)
                    '.Item(.Count).SubItems(3) = StartUpFolderType 'StartUpItemsArray(0, i)
                    FileExtension = ExtractFileExtensionFromPath(StartUpItemsArray(0, I))
                    '//Handle the case where the startup folder contains direct files not shortcuts to
                    '//only resolve "url" or "lnk" to avoid "ResolveShortcutPath" Error
                    If FileExtension = "url" Or FileExtension = "lnk" Then
                       .Item(.Count).SubItems(1) = ResolveShortcutPath(StartUpItemsArray(0, I))
                    Else
                       .Item(.Count).SubItems(1) = StartUpItemsArray(0, I)
                    End If
                    AddIconFileToList .Item(.Count).SubItems(1)
                    Dim FileDescription As String
                    FileDescription = FileInfo(.Item(.Count).SubItems(1)).FileDescription
                    If FileDescription <> "" Then
                       .Item(.Count).SubItems(2) = FileDescription & " by " & FileInfo(.Item(.Count).SubItems(1)).CompanyName
                    Else
                       .Item(.Count).SubItems(2) = .Item(.Count).Text
                    End If
                 End If
              End If
          
          Next
      End With
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddIconFileToList(FilePath As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

     Dim FI As SHFILEINFO
    
     Dim hDuplIcon As Long
    
     Const SHGFI_ICON = &H100

     SHGetFileInfo FilePath, 0, FI, Len(FI), SHGFI_ICON

     If FI.hicon = 0 Then
 
        'FI.hicon = LoadIcon(ByVal 0&, IDI_APPLICATION)
        'koko
        FilePath = ExtractPath(FilePath)
        FI.hicon = 0
        '//pt 15Apr2008: changed next line coz it coz's a GPF in hesham machine
        'FI.hicon = ExtractAssociatedIcon(App.hInstance, FilePath, 5)
        ExtractIconEx FilePath, 0, FI.hicon, ByVal 0&, 1
 
     End If
     
     hDuplIcon = DuplicateIcon(ByVal 0&, FI.hicon)  ' Duplicate the returned handle
  
     frmmain.PIco.Cls '//if we didn't cls PIco then PIco will display old icons with the new one
        
     DrawIcon frmmain.PIco.hDC, 0, 0, hDuplIcon  '//Copy the icon to PIco
    
     DestroyIcon hDuplIcon
            
    '//To avoid this error in line 10 (next line) "ImageList cannot be modified while another control is bound to it"
    '@@@@@@@@@@@@@@@@@@@frmmain.lstStartupPrograms.SmallIcons = Nothing
    '//make a fixed size for the icon, if we didn't add the following 2 lines
    '//then the icon size will be larger after any update for the listview "i don't know why"
    '//note that the size of PIco must be 480,480 "bardo i don't know why"
    frmmain.ILAppIcons.ImageHeight = 16
    frmmain.ILAppIcons.ImageWidth = 16
          
    'Debug.Print "Out of memory goes here if timer1_timer alot"
    '@@@@@@@@@@@@@@@@ "Out of Memory error here"'''''''''''''''''''''''''''''''''''''''''''''
20  frmmain.ILAppIcons.ListImages.Add , , frmmain.PIco.Image  '//Add the icon to the listview
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    frmmain.lstStartupPrograms.SmallIcons = frmmain.ILAppIcons '//bind the imagelist to the listview
      
    frmmain.lstStartupPrograms.ListItems(frmmain.lstStartupPrograms.ListItems.Count).SmallIcon = frmmain.ILAppIcons.ListImages.Count
  
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CheckRegForNewItems(Root As Long, KeyPath As String) As Long
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Const BUFFER_SIZE As Long = 255
    
    Dim KPath As String
    Dim NewItem As ListItem
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
               
               Set NewItem = Nothing
               
               Set NewItem = frmmain.lstStartupPrograms.FindItem(Left$(sName, Ret))
               
               On Error Resume Next
               If NewItem.Text = "" Then
               
                      'Dim FileName As String
                      'FileName = ExtractFileNameFromPath(Left$(sData, RetData - 1))
                      'If FileName <> "" Then
                      '   ShowAddedAlert ExtractFileNameFromPath(Left$(sData, RetData - 1))
                      'Else
                      '   ShowAddedAlert Left$(sName, Ret)
                      'End If
                      
                      '//pt 31Mar2008 04:23 pm : For unknown reason : if the sort is enabled, adding a new item wrongly doesn't added after the last
                      '//item and added in the middle of items causing alot of errors
                      frmmain.lstStartupPrograms.Sorted = False
                      
                      ShowAddedAlert Left$(sName, Ret)
                      
                      Dim NewItemIndex As Integer
                      
                      NewItemIndex = .Count + 1
                      .Add , , Left$(sName, Ret)
                      .Item(NewItemIndex).SubItems(1) = Left$(sData, RetData - 1)
                      CompanyName = FileInfo(.Item(NewItemIndex).SubItems(1)).CompanyName
                      FileDescription = FileInfo(.Item(NewItemIndex).SubItems(1)).FileDescription
                      If FileDescription = "" Then
                         FileDescription = .Item(.Count).Text
                         .Item(NewItemIndex).SubItems(2) = FileDescription
                      Else
                         .Item(NewItemIndex).SubItems(2) = FileDescription & " by " & CompanyName
                      End If
                      .Item(NewItemIndex).ToolTipText = ExtractPath(Left$(sData, RetData - 1))
                      AddIconFileToList .Item(NewItemIndex).SubItems(1)
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
                    
                      .Item(NewItemIndex).SubItems(3) = KPath
               End If
               On Error GoTo 0
               
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
        
        CheckRegForNewItems = 1
        
    Else
    
        CheckRegForNewItems = 0
    
    End If
    
    End With
    
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CheckStartUpFoldersItems(StartupFolderPath As String, StartUpFolderType As StartUpType)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

      '//pt 17Feb2008 06:32 pm : Add StartUp Folder Items//////////////////////////////////////////////////////////
      Dim I As Integer
      Dim FileSize As Long
      Dim NewItem As ListItem
      Dim ItemName As String
    '  Dim CUStartUpFolderItems() As String
    '  Dim StartUpItemsArray() As String
      Dim NumFiles As Integer, NumDirs As Integer
      Dim StartUpItemsArray() As String
      Dim FileExtension As String
      
      NumDirs = 0
      NumFiles = 0
          
      FileSize = FindFilesAPI(StartupFolderPath, "*.*", NumFiles, NumDirs, StartUpItemsArray)
      
      With frmmain.lstStartupPrograms.ListItems
      
          For I = 0 To UBound(StartUpItemsArray, 2) - 1
              
              '//No Need To Display "desktop.ini" file
              If StartUpItemsArray(1, I) <> "desktop.ini" Then
                 If StartUpItemsArray(1, I) <> "" Then  'And InStr(1, StartUpItemsArray(1, I), ".")
                    
                    Set NewItem = Nothing

                    'Debug.Print StartUpItemsArray(1, i)
                    If InStr(1, StartUpItemsArray(1, I), ".") Then '//if the file has extension then
                       Dim FileExtensionCharsNum As Integer
                       FileExtensionCharsNum = Len(ExtractFileExtensionFromPath(StartUpItemsArray(1, I)))
                       ItemName = Left$(StartUpItemsArray(1, I), Len(StartUpItemsArray(1, I)) - (FileExtensionCharsNum + 1))
                    Else
                       ItemName = StartUpItemsArray(1, I)
                    End If
                    
                    Set NewItem = frmmain.lstStartupPrograms.FindItem(ItemName)
                    
                    On Error Resume Next
                                           
                     '   Dim Proceed As Boolean
                     '   If NewItem.Text <> "" Then
                     '       Select Case StartUpFolderType
                     '
                     '              Case 0
                     '
                     '                   If NewItem.SubItems(3) <> "All Users Startup" Then
                     '                      NewItem.Text = ""
                     '                   End If
                     '
                     '              Case 1
                     '
                     '                   If NewItem.SubItems(3) <> "Current User Startup" Then
                     '                      NewItem.Text = ""
                     '                   End If
                     '
                     '       End Select
                     '   End If
                        
                        If NewItem.Text = "" Then
                            'Proceed = False
                            ShowAddedAlert ItemName
                            .Add , , ItemName
                            '.Add , , Left$(StartUpItemsArray(1, I), Len(StartUpItemsArray(1, I)) - (FileExtensionCharsNum + 1))
                            .Item(.Count).SubItems(3) = IIf(StartUpFolderType, "Current User Startup", "All Users Startup")
                            '//Save the shortcut file name in the tag property in order to use in "CheckList" function
                            .Item(.Count).Tag = StartUpItemsArray(1, I)
                            '.Item(.Count).SubItems(3) = StartUpFolderType 'StartUpItemsArray(0, i)
                            FileExtension = ExtractFileExtensionFromPath(StartUpItemsArray(0, I))
                            '//Handle the case where the startup folder contains direct files not shortcuts to
                            '//only resolve "url" or "lnk" to avoid "ResolveShortcutPath" Error
                            If FileExtension = "url" Or FileExtension = "lnk" Then
                               .Item(.Count).SubItems(1) = ResolveShortcutPath(StartUpItemsArray(0, I))
                            Else
                               .Item(.Count).SubItems(1) = StartUpItemsArray(0, I)
                            End If
                            AddIconFileToList .Item(.Count).SubItems(1)
                            Dim FileDescription As String
                            FileDescription = FileInfo(.Item(.Count).SubItems(1)).FileDescription
                            If FileDescription <> "" Then
                               .Item(.Count).SubItems(2) = FileDescription & " by " & FileInfo(.Item(.Count).SubItems(1)).CompanyName
                            Else
                               .Item(.Count).SubItems(2) = .Item(.Count).Text
                            End If
                     End If
                 End If
              End If
          
          Next
      End With
      
      Erase StartUpItemsArray()
      
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ShowAddedAlert(Prompt As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    If rMonitorAllowed = False Then Exit Sub
 
    If rchkNewlyAdded = False Then Exit Sub
    
    If frmmain.TrayIcon.IsDisplayed Then
       frmmain.TrayIcon.ShowBalloon Prompt & " added to the startup", "Application Added", Info, 10000
    Else
       frmmain.TrayIcon.Show frmmain.Icon, SoftwareName
       frmmain.TrayIcon.ShowBalloon Prompt & " added to the startup", "Application Added", Info, 10000
       frmmain.tmrGenChecks.Enabled = False
       frmmain.tmrEnabletmrGenChecks.Enabled = True
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ShowRemovedAlert(Prompt As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If rMonitorAllowed = False Then Exit Sub
    
    If rchkRemoved = False Then Exit Sub
    
    If frmmain.TrayIcon.IsDisplayed Then
       frmmain.TrayIcon.ShowBalloon Prompt & " removed from the startup", "Application Removed", Info, 10000
    Else
       frmmain.TrayIcon.Show frmmain.Icon, SoftwareName
       frmmain.TrayIcon.ShowBalloon Prompt & " removed from the startup", "Application Removed", Info, 10000
       frmmain.tmrGenChecks.Enabled = False
       frmmain.tmrEnabletmrGenChecks.Enabled = True
    End If
    
End Sub

'//////////////////////////////////////////////////////////////'
'//HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\

'Current User, Run       "Done but Test"

'Current User, RunOnce   "Done but Test"

'==================

'StartUp Folder

'StartUp, Current User

'StartUp, All Users

'==================

'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\

'All Users, Run        "Done but Test"

'All Users, RunOnce    "Done but Test"

'All Users, RunOnceEx  "Done but Test"

'==================

'//won't implement this
'//HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad
'All Users, DelayLoad

