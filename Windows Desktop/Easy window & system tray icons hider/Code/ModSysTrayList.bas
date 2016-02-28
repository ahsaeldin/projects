Attribute VB_Name = "ModSysTrayList"
                                            '¶¶¶¶¶¶¶¶
                                            '¶¶¶¶¶¶¶
                                            '¶¶¶
                                            '¶¶¶¶¶¶
                                            '¶¶¶¶¶¶ aithful
                                            '¶¶¶
                                            '¶¶¶
                                            '
                                            '¶¶¶¶¶¶¶
                                            '¶¶¶__¶¶¶
                                            '¶¶¶__¶¶¶
                                            '¶¶¶¶¶¶¶
                                            '¶¶¶¶¶¶ espectful
                                            '¶¶¶_¶¶¶
                                            '¶¶¶__¶¶¶
                                            '
                                            '¶¶¶
                                            '¶¶¶
                                            '¶¶¶
                                            '¶¶¶ ntelligent
                                            '¶¶¶
                                            '¶¶¶
                                            '¶¶¶
                                            '
                                            '¶¶¶¶¶¶¶¶
                                            '¶¶¶¶¶¶¶
                                            '¶¶¶
                                            '¶¶¶¶¶¶ verlasting Or loyal
                                            '¶¶¶
                                            '¶¶¶¶¶¶¶
                                            '¶¶¶¶¶¶¶¶
                                            '
                                            '¶¶¶__¶¶¶
                                            '¶¶¶¶_¶¶¶
                                            '¶¶¶¶¶¶¶¶
                                            '¶¶¶¶¶¶¶¶ aughtyand
                                            '¶¶¶_¶¶¶¶
                                            '¶¶¶__¶¶¶
                                            '¶¶¶__¶¶¶
                                            '
                                            '¶¶¶¶¶¶
                                            '¶¶¶¶¶¶¶¶
                                            '¶¶¶__¶¶¶¶
                                            '¶¶¶___¶¶¶earest of all
                                            '¶¶¶__¶¶¶¶
                                            '¶¶¶¶¶¶¶¶

Option Explicit
            
Public Type TrayIconInfoType
    hwnd As Long
    uId As Long
    ucallbackMessage As Long
    Param1(1) As Long
    hIcon As Long
    Param2(2) As Long
    APath As String
    ToolTip As String
    Bitmap As Long
    Command As Long
    State As Byte
    Style As Byte
    Data As Long
    str As Long
End Type

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '"Success is a lousy teacher. It seduces smart people into thinking they can't lose.”'
            '                                                                    Bill Gates Qutoe'
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public hProc As Long
Public ptrVM As Long
Private Declare Function DuplicateIcon Lib "shell32.dll" (ByVal hInst As Long, ByVal hIcon As Long) As Long
'//Why i added this declare and didn't use the one stored in the type lib??
'//because the one stored in the type lib is "ByVal lpIconName As String" and this cozs a failed to load
'//IDI_APPLICATION  "IconHandle = LoadIcon(ByVal 0&, IDI_APPLICATION)" in LoadSysTrayList function below
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub LoadSysTrayList()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: LoadSysTrayList
' Description: LoadSysTrayList Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo LoadSysTrayListErr
        
    Dim i As Integer
    Dim iIndex As Integer
    Dim IconHandle As Long
    Dim HiddenIndex As Integer
    Dim SysListItemIndex As Integer
    Dim IsThereAreHiddenItems As Boolean
    Dim TrayListArrType() As TrayIconInfo
      
    For SysListItemIndex = 1 To frmmain.lstSystrayIcons.ListItems.count
        If frmmain.lstSystrayIcons.ListItems.Item(SysListItemIndex).Checked Then
           If frmmain.lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1) = "Hide" Then
              HiddenIndex = HiddenIndex + 1
              frmmain.lstHiddenItems.ListItems.Add
              
              frmmain.lstHiddenItems.ListItems.Item(frmmain.lstHiddenItems.ListItems.count) = _
              frmmain.lstSystrayIcons.ListItems.Item(SysListItemIndex)
              
              frmmain.lstHiddenItems.ListItems.Item(frmmain.lstHiddenItems.ListItems.count).SubItems(1) = _
              frmmain.lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1)
              
              frmmain.lstHiddenItems.ListItems.Item(frmmain.lstHiddenItems.ListItems.count).SubItems(2) = _
              frmmain.lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2)
              
              frmmain.lstHiddenItems.ListItems.Item(frmmain.lstHiddenItems.ListItems.count).SubItems(3) = _
              frmmain.lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(3)
              
              frmmain.lstHiddenItems.ListItems.Item(frmmain.lstHiddenItems.ListItems.count).SubItems(4) = _
              frmmain.lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(4)
              
              frmmain.lstHiddenItems.ListItems.Item(frmmain.lstHiddenItems.ListItems.count).SubItems(5) = _
              frmmain.lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(5)
              
              frmmain.lstHiddenItems.ListItems.Item(frmmain.lstHiddenItems.ListItems.count).SubItems(6) = _
              frmmain.lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(6)
              
              frmmain.lstHiddenItems.ListItems.Item(frmmain.lstHiddenItems.ListItems.count).Tag = _
              frmmain.lstSystrayIcons.ListItems.Item(SysListItemIndex).Tag
 
              IsThereAreHiddenItems = True
              
           End If
        End If
    Next SysListItemIndex
    
    TrayListArrType = frmmain.TrayList.GetSysTrayIcons
    
    With frmmain
        .lstSystrayIcons.ListItems.Clear
        .lstSystrayIcons.SmallIcons = Nothing
        .ILSysIcons.ListImages.Clear
        '//make a fixed size for the icon, if we didn't add the following 2 lines
        '//then the icon size will be larger after any update for the listview "i don't know why"
        '//note that the size of PIco must be 480,480 "bardo i don't know why"
        .ILSysIcons.ImageHeight = 16
        .ILSysIcons.ImageWidth = 16
        
        For i = 0 To UBound(TrayListArrType)
            '//if icon tooltip isn't empty and there is already an icon
            If TrayListArrType(i).ToolTip <> "" Then
                If TrayListArrType(i).hIcon <> 0 Then
                   
                   .PicSysIco.Cls
                   
                   IconHandle = CopyIcon(TrayListArrType(i).hIcon)
                   'IconHandle = DuplicateIcon(ByVal 0&, TrayListArrType(i).hIcon)
                   '//added "If IconHandle <> 0 Then" to avoid displaying MSN and McAfee icons until i found
                   '//a solution to their bug
                
                   If Not IsWindowsNT Then GoTo win98:
                   If IconHandle <> 0 Then
                      'IconHandle = DuplicateIcon(ByVal 0&, TrayListArrType(i).hIcon)
                      '//In some icons like "McAfee Security Center" icon, the iconhandle = 0 so no icon displayed
                      '//hence i will try to get the icon in the next block and if i failed i will load IDI_APPLICATION
                      '//Application Icon
                      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
win98:
                      If IconHandle = 0 Then
                      
                         IconHandle = SendMessage(TrayListArrType(i).hwnd, WM_GETICON, ICON_BIG, 0&)
                     
                         If IconHandle = 0 Then '//try to get icon if the above failed
                            IconHandle = CopyIcon(GetClassLong(TrayListArrType(i).hwnd, GCL_HICONSM))
                         End If
                                     
                         If IconHandle = 0 Then '//try to get icon if the above failed
                            IconHandle = SendMessage(TrayListArrType(i).hwnd, WM_GETICON, ICON_SMALL, 0&)
                         End If
                                     
                         If IconHandle = 0 Then '//try to get icon if the above failed
                            IconHandle = CopyIcon(GetClassLong(TrayListArrType(i).hwnd, GCL_HICON))
                         End If
                                    
                         If IconHandle = 0 Then '//If this too returns false, give up and load a default icon by calling LoadIcon
                            IconHandle = LoadIcon(ByVal 0&, IDI_QUESTION)
                         End If
                   
                      End If
                      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                   
                      DrawIcon .PicSysIco.hDC, 0, 0, IconHandle
                      '//Why didn't we use i reather than using iIndex??
                      '//coz sometimes hicon = 0 hence the next icon that hicon <> 0 will have
                      '//an index in the ILSysIcons that preceed an empty index
                      '//i.e. ILSysicon(3) but ILsysicon(2) isn't existed yet and this will cause index out of bound error
                      iIndex = iIndex + 1
                      .ILSysIcons.ListImages.Add iIndex, , .PicSysIco.Image
                      .lstSystrayIcons.SmallIcons = .ILSysIcons
                      .lstSystrayIcons.ListItems.Add iIndex, , TrayListArrType(i).ToolTip, , iIndex
                      .lstSystrayIcons.ListItems.Item(iIndex).SubItems(1) = "Show"
                      .lstSystrayIcons.ListItems.Item(iIndex).SubItems(2) = TrayListArrType(i).hwnd
                      .lstSystrayIcons.ListItems.Item(iIndex).SubItems(3) = TrayListArrType(i).uId
                      .lstSystrayIcons.ListItems.Item(iIndex).SubItems(4) = TrayListArrType(i).hIcon
                      .lstSystrayIcons.ListItems.Item(iIndex).SubItems(5) = TrayListArrType(i).ToolTip
                      .lstSystrayIcons.ListItems.Item(iIndex).SubItems(6) = TrayListArrType(i).ucallbackMessage
                      .lstSystrayIcons.ListItems.Item(iIndex).ListSubItems(2).ForeColor = vbWhite
                      .lstSystrayIcons.ListItems.Item(iIndex).ListSubItems(3).ForeColor = vbWhite
                      .lstSystrayIcons.ListItems.Item(iIndex).ListSubItems(4).ForeColor = vbWhite
                      .lstSystrayIcons.ListItems.Item(iIndex).ListSubItems(5).ForeColor = vbWhite
                      .lstSystrayIcons.ListItems.Item(iIndex).ListSubItems(6).ForeColor = vbWhite
                      
                      '//Used to check get the selected Item later in TrayList
                      .lstSystrayIcons.ListItems.Item(iIndex).Tag = TrayListArrType(i).hIcon
                   End If
                End If
            End If
        Next i
    End With
    
    i = 1
    
    If IsThereAreHiddenItems Then
       
       For i = 1 To HiddenIndex
           
           frmmain.PicSysIco.Cls
           
           IconHandle = CopyIcon(frmmain.lstHiddenItems.ListItems.Item(i).Tag)
           
           If IconHandle = 0 Then
                      
              IconHandle = SendMessage(TrayListArrType(i).hwnd, WM_GETICON, ICON_BIG, 0&)
                                        
              If IconHandle = 0 Then '//try to get icon if the above failed
                 IconHandle = CopyIcon(GetClassLong(frmmain.lstHiddenItems.ListItems.Item(i).SubItems(2), GCL_HICONSM))
              End If
                                     
              If IconHandle = 0 Then '//try to get icon if the above failed
                 IconHandle = SendMessage(frmmain.lstHiddenItems.ListItems.Item(i).SubItems(2), WM_GETICON, ICON_SMALL, 0&)
              End If
                                     
              If IconHandle = 0 Then '//try to get icon if the above failed
                 IconHandle = CopyIcon(GetClassLong(frmmain.lstHiddenItems.ListItems.Item(i).SubItems(2), GCL_HICON))
              End If
                                    
              If IconHandle = 0 Then '//If this too returns false, give up and load a default icon by calling LoadIcon
                 IconHandle = LoadIcon(ByVal 0&, IDI_QUESTION)
              End If
           End If
                          
           DrawIcon frmmain.PicSysIco.hDC, 0, 0, IconHandle
                          
           frmmain.ILSysIcons.ListImages.Add iIndex + i, , frmmain.PicSysIco.Image

           frmmain.lstSystrayIcons.SmallIcons = frmmain.ILSysIcons
           
           frmmain.lstSystrayIcons.ListItems.Add , , , , iIndex + i
           
           frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count) = _
           frmmain.lstHiddenItems.ListItems.Item(i)
       
           frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(1) = _
           frmmain.lstHiddenItems.ListItems.Item(i).SubItems(1)
       
           frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(2) = _
           frmmain.lstHiddenItems.ListItems.Item(i).SubItems(2)
       
           frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(3) = _
           frmmain.lstHiddenItems.ListItems.Item(i).SubItems(3)
       
           frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(4) = _
           frmmain.lstHiddenItems.ListItems.Item(i).SubItems(4)
       
           frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(5) = _
           frmmain.lstHiddenItems.ListItems.Item(i).SubItems(5)
       
           frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(6) = _
           frmmain.lstHiddenItems.ListItems.Item(i).SubItems(6)
       
           frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).Tag = _
           frmmain.lstHiddenItems.ListItems.Item(i).Tag
           
           Dim res As Integer
           res = GetINIProfileInt("chkGray", 1)
           If res Then
              frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).ForeColor = HiddenWindowForeColor
           End If
                  
       Next i
       
        frmmain.lstHiddenItems.ListItems.Clear
        
    End If

    
LoadSysTrayListExit:
    On Error Resume Next

    Exit Sub
    
LoadSysTrayListErr:

    '//pt 14Aug2007: added End to close the software if "Out of memory"
    End
    
    '//0x0f3ee120 number is a unique number for MyLoadSysTrayList because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee120"
    
GoTo LoadSysTrayListExit

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CheckSysTrayList()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: CheckSysTrayList
' Description: CheckSysTrayList Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo CheckSysTrayListErr
        
    Dim ret As Long
    Dim hTray As Long
    Dim ptrTip As Long
    Dim tbut As TBBUTTON
    Static Old_i As Long
    Dim TChar As Integer
    Dim iIndex As Integer
    Dim CheckDup As Boolean
    Dim Icons() As TrayIconInfoType
    Dim i  As Long   '//number of visible windows
    Dim k As Long, r As Long   '//r and k ues for the loops below
   
    
    '//pt 12July2007: commented next line and replaced with the preceding code block because:
    '//frequenlty calling to TrayList.IconCount causes a "OverFlow Error"
    '//well, now you ask what is the reason of OverFlow Error??
    '//Okay, open ASTC source and open GetTrayIcons fucntion then
    '//you will found this line:
    '//"If ptrVM = 0 Then Err.Raise &H80070000 & Err.LastDllError" and this is the reason
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '//For unknown reason, this line in CheckSysTrayList function
    '//i = frmmain.TrayList.IconCount
    '//causes a delay in the program startup, you can notice this if you click
    '//frmopions just after you run the exe
    '//hence, i will make tmrUpdateListView.Interval = 4000 in the property pages
    '//then and i will update the listviews in form_load
    '//also inside tmrUpdateListView_Timer, i will return the timer interval again to 1, just after the startup period get passed
    'i = frmmain.TrayList.IconCount
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    hTray = GetHwndTray
     
    If IsWindowsNT Then
       i = SendMessage(hTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)
    Else
       i = frmmain.TrayList.IconCount
       Icons = frmmain.TrayList.GetSysTrayIcons
    End If
    
    If i = Old_i Then
       If IsWindowsNT Then
                   
                   '//OurTraget: Update the listview if the tooltip or icon changed
                   '//to achieve this target we need to compare listview items with the current systray info
                   '//so we need to get systray info 3ly tool
                   '//if we get systray info 3ly tool, we will get an "over flow" error after a period because of using "VirtualAllocEx" in ASTC Control
                   '//hence i added CreateVirutalMem function to use VirtualAllocEx only once and to get the icons info manually
            
                   ReDim Icons(i - 1)
              
                   For k = 0 To frmmain.lstSystrayIcons.ListItems.count - 1
              
                       On Error Resume Next
                       If frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(1) = "Hide" Then GoTo NextItem:
                       On Error GoTo CheckSysTrayListErr
              
                       For r = 0 To i - 1
              
                           SendMessage hTray, TB_GETBUTTON, r, ByVal ptrVM
              
                           ReadProcessMemory hProc, ByVal ptrVM, tbut, Len(tbut), ret
              
                           ptrTip = tbut.iString
              
                           Do
              
                             ReadProcessMemory hProc, ByVal ptrTip, TChar, 2&, ret
              
                             Icons(r).ToolTip = Icons(r).ToolTip & ChrW$(TChar)
              
                             ptrTip = ptrTip + 2
              
                           Loop While TChar
              
                           ReadProcessMemory hProc, ByVal tbut.dwData, Icons(r).hwnd, 36&, ret
              
                           If Icons(r).hIcon <> frmmain.lstSystrayIcons.ListItems.Item(k + 1).Tag Then
              
                              CheckDup = True '//Not found this systray icons in the GloTrayListArrType
              
                           Else
              
                              '//here 's another case:
                              '//If the user open firefox, and suppose that firefox has a systray icon which is
                              '//"Mozilla FireFox" however if he opens any site after this the tooltip will change to "Site Name - Mozilla FireFox"
                              '//hence the word "Mozilla FireFox" is still in the title and this will cause instr function to return a velue, if
                              '//InStr returns a value then the tooltip won't change in the next below line altought it is already changed , hence
                              '//i will check here if the len of the window title is larger than its item in the listview by only 1 char hence
                             ' If Len(Icons(r).Tooltip) - Len(frmmain.lstSystrayIcons.ListItems.Item(k + 1)) = 1 Then
              
                                  '//this case is something like this "Mozilla FireFox " and "Mozilla FireFox"
              
                             '    CheckDup = False  '//Found this window in the Tasklist
              
                             '    Exit For
              
                             ' Else
                                  '//this case is something like this "Site Name - Mozilla FireFox " and "Mozilla FireFox"
              
                                 CheckDup = False '//Not Found this window in the TaskList
                                 Exit For
                             ' End If
              
                           End If
              
                       Next r
              
                       If CheckDup Then
              
                          For iIndex = 0 To i - 1
              
                               If frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(2) = Icons(iIndex).hwnd Then
                                  frmmain.PicSysIco.Cls
                                  DrawIcon frmmain.PicSysIco.hDC, 0, 0, Icons(iIndex).hIcon
                                  frmmain.ILSysIcons.ListImages.Add frmmain.ILSysIcons.ListImages.count + 1, , frmmain.PicSysIco.Image
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).Text = Icons(iIndex).ToolTip
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SmallIcon = frmmain.ILSysIcons.ListImages.count
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(2) = Icons(iIndex).hwnd
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(3) = Icons(iIndex).uId
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(4) = Icons(iIndex).hIcon
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(5) = Icons(iIndex).ToolTip
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(6) = Icons(iIndex).ucallbackMessage
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).Tag = Icons(iIndex).hIcon
              
                                  Exit Sub
              
                              End If
              
                          Next iIndex
              
                          LoadSysTrayList
              
                       End If
            
NextItem:
                   Next k
                   
       Else 'Any window rather than WinNT
       
                  For k = 0 To frmmain.lstSystrayIcons.ListItems.count - 1
              
                       On Error Resume Next
                       If frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(1) = "Hide" Then GoTo NextItem2:
                       On Error GoTo CheckSysTrayListErr
              
                       For r = 0 To i - 1
                 
                           If Icons(r).hIcon <> frmmain.lstSystrayIcons.ListItems.Item(k + 1).Tag Then
              
                              CheckDup = True '//Not found this systray icons in the GloTrayListArrType
              
                           Else
              
                              '//here 's another case:
                              '//If the user open firefox, and suppose that firefox has a systray icon which is
                              '//"Mozilla FireFox" however if he opens any site after this the tooltip will change to "Site Name - Mozilla FireFox"
                              '//hence the word "Mozilla FireFox" is still in the title and this will cause instr function to return a velue, if
                              '//InStr returns a value then the tooltip won't change in the next below line altought it is already changed , hence
                              '//i will check here if the len of the window title is larger than its item in the listview by only 1 char hence
                             ' If Len(Icons(r).Tooltip) - Len(frmmain.lstSystrayIcons.ListItems.Item(k + 1)) = 1 Then
              
                                  '//this case is something like this "Mozilla FireFox " and "Mozilla FireFox"
              
                             '    CheckDup = False  '//Found this window in the Tasklist
              
                             '    Exit For
              
                             ' Else
                                  '//this case is something like this "Site Name - Mozilla FireFox " and "Mozilla FireFox"
              
                                 CheckDup = False '//Not Found this window in the TaskList
                                 Exit For
                             ' End If
              
                           End If
              
                       Next r
              
                       If CheckDup Then
              
                          For iIndex = 0 To i - 1
              
                               If frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(2) = Icons(iIndex).hwnd Then
                                  frmmain.PicSysIco.Cls
                                  DrawIcon frmmain.PicSysIco.hDC, 0, 0, Icons(iIndex).hIcon
                                  frmmain.ILSysIcons.ListImages.Add frmmain.ILSysIcons.ListImages.count + 1, , frmmain.PicSysIco.Image
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).Text = Icons(iIndex).ToolTip
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SmallIcon = frmmain.ILSysIcons.ListImages.count
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(2) = Icons(iIndex).hwnd
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(3) = Icons(iIndex).uId
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(4) = Icons(iIndex).hIcon
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(5) = Icons(iIndex).ToolTip
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).SubItems(6) = Icons(iIndex).ucallbackMessage
                                  frmmain.lstSystrayIcons.ListItems.Item(k + 1).Tag = Icons(iIndex).hIcon
              
                                  Exit Sub
              
                              End If
              
                          Next iIndex
              
                          LoadSysTrayList
              
                       End If
            
NextItem2:
                   Next k
       End If
       
    Else
       
       LoadSysTrayList
       
       Old_i = i
    
    End If
  
    
CheckSysTrayListExit:
    On Error Resume Next

    Exit Sub
    
CheckSysTrayListErr:
    
    '//pt 2Aug2007: added End to close the software if "Subscript out of rang"
    End
    
    '//CheckSysTrayListunique number is a unique number for MyCheckSysTrayList because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee121"
    
GoTo CheckSysTrayListExit

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetHwndTray()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: GetHwndTray
' Description: GetHwndTray Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo GetHwndTrayErr
        
    Dim hwnd As Long, hwnd2 As Long
      
    hwnd = FindWindow("Shell_TrayWnd", vbNullString)
    hwnd = FindWindowEx(hwnd, ByVal 0&, "TrayNotifyWnd", vbNullString)
    hwnd2 = FindWindowEx(hwnd, ByVal 0&, "SysPager", vbNullString)
      
    If (hwnd2 = 0) Then hwnd2 = hwnd ' ME,2000
      
    hwnd2 = FindWindowEx(hwnd2, ByVal 0&, "ToolbarWindow32", vbNullString) ' ME, 2000, XP...
      
    If (hwnd2 = 0) Then
       GetHwndTray = hwnd ' 95,98
    Else
       GetHwndTray = hwnd2  ' ME, 2000, XP...
    End If
    
GetHwndTrayExit:
    On Error Resume Next

    Exit Function
    
GetHwndTrayErr:

    '//0x0f3ee122 number is a unique number for MyGetHwndTray because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee122"
    
GoTo GetHwndTrayExit

End Function

'//OurTraget: Update the listview if the tooltip or icon changed
'//to achieve this target we need to compare listview items with the current systray info
'//so we need to get systray info 3ly tool
'//if we get systray info 3ly tool, we will get an "over flow" error after a period because of using "VirtualAllocEx" in ASTC Control
'//hence i added CreateVirutalMem function to use VirtualAllocEx only once and to get the icons info manually
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CreateVirutalMem()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
    Dim pid As Long
    Dim tbut As TBBUTTON
    
    GetWindowThreadProcessId GetHwndTray, pid
    
    hProc = OpenProcess(PROCESS_VM, 0, pid)
    
    ptrVM = VirtualAllocEx(hProc, ByVal 0&, Len(tbut), MEM_COMMIT, PAGE_READWRITE)

End Sub
