VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'//pt 31May2007: added this class module to handle all methods of related to the windows
'//like IsTask , Get the visible windows etc...
    
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'
    '“You may never know what results come of your action, but if you do nothing there will be no result” '
    '                                                                                       Mahatma Gandhi'
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'

'//created in 28May2007 ,removed from ModGlobels and here in 31May2007
'//PT 28May2007: IsTask check by window's Hwnd, if it is a visible window
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function IsTask(hwndTask As Long) As Long
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: IsTask
' Description: Returns 1 if the window is visible, or 0 if not or, 2 the function fails
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software

' Modification History:

'============================================================

    On Error GoTo IsTaskErr
        
    Dim intLen As Long
    
    Dim strTitle As String
     
    Dim DragWindowClassNameLength As Long
    
    Dim DragWindowClassName As String * 32
           
    Const IsTaskStyle = WS_VISIBLE Or WS_BORDER '//used to check bitwise below with WndStyle
    
    Dim WndStyle As Long '//receives the return value of GetWindowLong API function
    
    Dim IsHiddenWindow As Long '//Used to check if the window is hidden by this software
    
    '//////////////////////////////////////////////////////////////////////////////////////////////////////
    'i 've come to believe that all my past failure and frustration were actually
    'laying the foundation for the understandings that have created the new level of living I now enjoy.
    '//Anthony Robbins
    '//////////////////////////////////////////////////////////////////////////////////////////////////////
    
    '//'''''''''''''''''''''''''''''''''''Check if this window has a hidden Systray Icon''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim IsInTheList As Boolean
    Dim IsHiddenSysIconHwnd As Long
   
    '//if the window has SysHwnd prob then it was hidden before closing Me
    IsHiddenSysIconHwnd = GetProp(hwndTask, "SysHwnd")
     
    Dim i As Integer
    
    '//Check First if it isn't already in the listview
    For i = 1 To frmmain.lstSystrayIcons.ListItems.count
        
        If frmmain.lstSystrayIcons.ListItems.Item(i).SubItems(1) = "Hide" Then
           If frmmain.lstSystrayIcons.ListItems.Item(i).SubItems(2) = IsHiddenSysIconHwnd Then
              IsInTheList = True
           End If
        End If
        
    Next i
      
    '//if it is not in the listview
    If IsHiddenSysIconHwnd <> 0 Then
       If Not IsInTheList Then
          
          Dim SysuId As Long
          Dim SyshIcon As Long
          Dim FFile As Integer
          Dim iIndex As Integer
          Dim IconHandle As Long
          Dim SysToolTip As String * 64
          Dim TotalRecordsNumber As Long
          Dim SysucallbackMessage As Long
          Dim SysToolTipTypeRecordLen As Long
          Dim SysToolTipType As SysIcoToolTip
            
          FFile = FreeFile()
          SysToolTipTypeRecordLen = Len(SysToolTipType)
          '//Extract the icon tooltip from the file
          Open AppPath + "ico.dat" For Random As FFile Len = SysToolTipTypeRecordLen
               TotalRecordsNumber = FileLen(AppPath + "ico.dat") / SysToolTipTypeRecordLen
               For iIndex = 1 To TotalRecordsNumber
                   Get #FFile, iIndex, SysToolTipType
                   '//compare the hwnd stored in the prob with hwnd stored in the file
                   If SysToolTipType.hwnd = IsHiddenSysIconHwnd Then '//if found and okay then
                      '//Get the other probs
                      SysuId = GetProp(IsHiddenSysIconHwnd, "SysuId")
                      
                      SyshIcon = GetProp(IsHiddenSysIconHwnd, "SyshIcon")
                      
                      SysToolTip = SysToolTipType.ToolTip
                      
                      SysucallbackMessage = GetProp(IsHiddenSysIconHwnd, "SysucallbackMessage")
                      '//Prepare the ico to be displayed in the small icons of the listview and fill the item
                      frmmain.PicSysIco.Cls
                      
                      IconHandle = CopyIcon(SyshIcon)
                      
                      DrawIcon frmmain.PicSysIco.hDC, 0, 0, IconHandle
                      
                      frmmain.ILSysIcons.ListImages.Add , , frmmain.PicSysIco.Image
                      
                      frmmain.lstSystrayIcons.SmallIcons = frmmain.ILSysIcons
                      
                      frmmain.lstSystrayIcons.ListItems.Add , , , , frmmain.ILSysIcons.ListImages.count
                       
                      '//Why we use "Left(SysToolTip, InStr(1, SysToolTip, "#&") - 1)" rather then using "SysToolTip"??
                      '//because "SysToolTip" contains a spaces "chr (32)" from the right to the end of the tooltip
                      '//the problem of the spaces appears in ShowSysTrayIcon function in this line:
                      '//"TrayI.szTip = lstSystrayIcons.ListItems(Index).SubItems(5) & Chr$(0)"
                      '//in other words Chr$(0) appears after the spaces hence the spaces appears in the tooltip when you
                      '//re-display the icon and this cozs the "..." after the tooltip
                      '//and for of of this i add a delimiter "#&" to the tooltip in cmdSysHide_Click to check here for
                      If InStr(1, SysToolTip, "#&") <> 0 Then
                         frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count) = _
                         Left(SysToolTip, InStr(1, SysToolTip, "#&") - 1)
                      Else
                         frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count) = _
                         SysToolTip
                      End If
                       
                      frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(1) = _
                      "Hide"
                    
                      Dim res As Integer
                      res = GetINIProfileInt("chkGray", 1)
                      If res Then
                         frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).ForeColor = HiddenWindowForeColor
                      End If
                    
                      frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(2) = _
                      IsHiddenSysIconHwnd
                    
                      frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(3) = _
                      SysuId
                    
                      frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(4) = _
                      SyshIcon
                    
                      '//Why we use "Left(SysToolTip, InStr(1, SysToolTip, "#&") - 1)" rather then using "SysToolTip"??
                      '//because "SysToolTip" contains a spaces "chr (32)" from the right to the end of the tooltip
                      '//the problem of the spaces appears in ShowSysTrayIcon function in this line:
                      '//"TrayI.szTip = lstSystrayIcons.ListItems(Index).SubItems(5) & Chr$(0)"
                      '//in other words Chr$(0) appears after the spaces hence the spaces appears in the tooltip when you
                      '//re-display the icon and this cozs the "..." after the tooltip
                      '//and for of of this i add a delimiter "#&" to the tooltip in cmdSysHide_Click to check here for
                      If InStr(1, SysToolTip, "#&") <> 0 Then
                         frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(5) = _
                         Left(SysToolTip, InStr(1, SysToolTip, "#&") - 1)
                      Else
                         frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(5) = _
                         SysToolTip
                      End If
                       
                      frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).SubItems(6) = _
                      SysucallbackMessage
                    
                      frmmain.lstSystrayIcons.ListItems.Item(frmmain.lstSystrayIcons.ListItems.count).Tag = _
                      SyshIcon

                   End If
               Next iIndex
            Close FFile
        
        End If
           
     End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If hwndTask = 0 Or hwndTask = Null Then '//handle the case if hwndTask = 0 or Null
       Err.Description = "Window Handle is Null"
       If InnerSysError Then SysError Err, "0x0f3ee15" '//display syserror if inner syserror is enabled by me
       IsTask = 2 '//returns 2 because the functions fails and can't proceed
       Exit Function
    End If
    
    WndStyle = GetWindowLong(hwndTask, GWL_STYLE) '//Get the window style in order check if it is a visible window
    
    If WndStyle = 0 Then 'Handle the case if GetWindowLong fails and returns 0
       Err.Description = "Can't get window Style"
      ' If InnerSysError Then SysError Err, "0x0f3ee15" '//display syserror if inner syserror is enabled by me
       IsTask = 2 '//returns 2 because the functions fails and can't proceed
       LogError "Can't get window Style "
       End '//End to avoid getting in a loop of displaying SysError messagebox
    End If
    IsHiddenWindow = GetProp(hwndTask, "VisibiltyState")  '//check if hidden by this software
    If IsHiddenWindow = 7878 Then '//7878 is a value specified by me and means it is hidden
       IsTask = True: Exit Function '//ok this window is a Task
    End If
   
    '//This comment clause to explain why we write these lines
    '//Const IsTaskStyle = WS_VISIBLE Or WS_BORDER
    '//If (WndStyle And IsTaskStyle) = IsTaskStyle Then
    
    '//First read this quoted clause from MSDN 98
    '//"GetWindowLong(hWnd, GWL_STYLE) returns a LONG value, which
    '//contains the currently active styles combined by the Boolean OR operator."
    '//so the return value of GetWindowLong returns something like "WS_VISIBLE Or WS_BORDER"
    '//and that's why we compare it with IsTaskStyle constant.
    If (WndStyle And IsTaskStyle) = IsTaskStyle Then
        IsTask = True
    End If
    '//What is the idea of this line?????
    '//If (WndStyle And IsTaskStyle) = IsTaskStyle Then
    '//this line is a bitwise AND
    
    '//well, but i can't understand this line
    '//If (WndStyle And IsTaskStyle) = IsTaskStyle Then
    '//here we make a "bitwise and", i.e compare IsTaskStyle '//bits with WndStyle bits using a bitwise and
    '//see this example:
    
    '//a = 5 or 6
    '//b = 5 or 6 or 8
    '//so with bitwise a and b = a
        
   
    strTitle = Space(32)
    
    GetClassName hwndTask, strTitle, 32  ' Get caption
   
    If InStr(1, strTitle, "Jet-Audio") <> 0 Then IsTask = True
        
    If InStr(1, strTitle, "WMPTransition") <> 0 Then
       IsTask = True
    End If
    
    If IsTask Then
    
        '//Why we Added + 1
        '//If the GetWindowText succeeds, the return value is the length, in characters,
        '//of the copied string, not including the terminating null character
        intLen = GetWindowTextLength(hwndTask) + 1 ' Get length
        'suppose that window title is ahmed, it means 5 char
        'GetWindowText returns title without null terminate
        'hence it will return ahme but if you add + 1 it will return ahmed
                 
        strTitle = Space(intLen)
                
        intLen = GetWindowText(hwndTask, strTitle, intLen) ' Get caption
                
                 
        '//added next line to ignore and avoiding adding tooltip window appearing when moving the mouse
        '//over a the X control box
        If Asc(strTitle) = 0 Then
           IsTask = False: Exit Function
        End If
        
        '//added next line to ignore and avoiding adding frmtip to the listview
        '//because if frmtip added to the listview it will be hided if the user press Hide All button
        Dim IsTipWindow As Long
        IsTipWindow = GetProp(hwndTask, "IsTipWindow")
        If IsTipWindow = 111 Then
            IsTask = False: Exit Function
        End If
        
        '//added next nested if to ignore "Drag" window appears if you drag any icon in the desktop
        If InStr(1, strTitle, "Drag") <> 0 Then
        
           DragWindowClassNameLength = GetClassName(hwndTask, DragWindowClassName, 32)
           
           If InStr(1, DragWindowClassName, "SysDragImage") <> 0 Then
              IsTask = False: Exit Function
           End If
        
        End If
        
    End If
    
IsTaskExit:
    On Error Resume Next

    Exit Function
    
IsTaskErr:

    '//0x0f3ee15 number is a unique number for MyTemplate because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee15"
    
GoTo IsTaskExit

End Function

'//created in 28May2007 ,removed from ModGlobels and here in 31May2007
'//PT 29May2007: ' Fills the Task structure with captions and hWnd(s) of all active programs
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function FillTaskList() As Long
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: FillTaskList
' Description:  Fills the Task structure with captions and hWnd of all active programs and returns the number of windows
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo FillTaskListErr
        
    Dim intLen As Long
    Dim cnt As Integer '//contains the number of the visible windows
    Dim hwndTask As Long '//will be used through out this function
    Dim bTask As Boolean '//holds the ret of IsTask Function
    Dim strTitle As String
    
    '//“Bill Gates is a very rich man today .. and do you want to know why? The answer is one word: versions.”
    '// Dave Barry
 
    For cnt = 0 To 200
        TaskList(cnt).TaskID = 0
        TaskList(cnt).TaskName = ""
    Next cnt
    
    cnt = 0
    
    hwndTask = GetWindow(GetDesktopWindow, GW_CHILD)  'Get First Window Handle
     
    If hwndTask = 0 Or hwndTask = Null Then
       Err.Description = "Can't Get Desktop Window"
       If InnerSysError Then SysError Err, "0x0f3ee17"
       FillTaskList = 0 '//returns 0 because the functions fails and can't proceed
       Exit Function
    End If
   
    Do While hwndTask 'repeat for all windows
    
        bTask = IsTask(hwndTask) '//check if hwndtask is a visible window
        
        If bTask Then
                    
            '//Why we Added + 1
            '//If the GetWindowText succeeds, the return value is the length, in characters,
            '//of the copied string, not including the terminating null character
            intLen = GetWindowTextLength(hwndTask) + 1 ' Get length
            'suppose that window title is ahmed, it means 5 char
            'GetWindowText returns title without null terminate
            'hence it will return ahme but if you add + 1 it will return ahmed
             
            strTitle = Space(intLen)
            
            intLen = GetWindowText(hwndTask, strTitle, intLen) ' Get caption
            
            If intLen = 0 Then LogError "GetWindowText returns zero"
            
            '//added next line to ignore and avoiding adding tooltip window appearing when moving the mouse
            '//over a the X control box
            If Asc(strTitle) = 0 Then GoTo ContinueLoop:
            
            TaskList(cnt).TaskName = strTitle
            
            TaskList(cnt).TaskID = hwndTask
            
            cnt = cnt + 1
        
        End If
        
        hwndTask = GetWindow(hwndTask, GW_HWNDNEXT) 'Get the next window in the z order
ContinueLoop:
    Loop
    
    FillTaskList = cnt
     
FillTaskListExit:
    On Error Resume Next

    Exit Function
    
FillTaskListErr:

    '//Templateunique number is a unique number for MyTemplate because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee17"
    
GoTo FillTaskListExit

End Function

'//created in 28May2007 ,removed from ModGlobels and here in 31May2007
'//PT 29May2007: added next function to activeate a window and bring it to top of z order
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function SwitchTo(hwnd As Long) As Long
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: SwitchTo
' Description: added next function to activeate a window and bring it to top of z order
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo SwitchToErr
        
    Dim ret As Long 'ret value for ShowWindow and SetWindowPos
    
    Dim WStyle As Long ' Window Style bits
    
    If hwnd = 0 Or hwnd = Null Then '//handle the case if hwnd = 0 or Null
       Err.Description = "Window's Handle to be switched is Null"
       If InnerSysError Then SysError Err, "0x0f3ee16"
       SwitchTo = 0 '//returns 0 because the functions fails and can't proceed
       Exit Function
    End If
    
    WStyle = GetWindowLong(hwnd, GWL_STYLE) 'Get style bits for window

    If WStyle = 0 Then
       Err.Description = "Can't get window style"
       If InnerSysError Then SysError Err, "0x0f3ee16"
       SwitchTo = 0 '//returns 0 because the functions fails and can't proceed
       Exit Function
    End If
    
    If WStyle And WS_MINIMIZE Then 'If minimized do a restore
        ret = ShowWindow(hwnd, SW_RESTORE)
        '//If the window was previously visible, the return value is nonzero.
        '//If the window was previously hidden, the return value is zero.
        '//==no need to check the ret value of ShowWindow Function
    End If
    
    'Move window to top of z-order/activate; no move/resize
    ret = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)

    SwitchTo = ret
    
SwitchToExit:
    On Error Resume Next

    Exit Function
    
SwitchToErr:

    '//0x0f3ee16 number is a unique number for MyTemplate because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee16"
    
GoTo SwitchToExit

End Function
