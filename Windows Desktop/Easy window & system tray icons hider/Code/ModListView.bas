Attribute VB_Name = "ModListView"
Option Explicit

'//pt 31May2007: added this  module to handle all methods of related to the listview
 
'¸.•*´¨`*•.¸¸.•*´¨`*•.¸¸.•*´¨`*•.¸¸.•*´¨`*•.¸¸.•
'::: (\_(\ ...*...*...*...*...*...*...*...*...*...*...*...*
'*: (=' :') : Kiss your life. Accept it, just as it is.
'•.. (,('')('')¤...*...*...*...*...*...*...*...*...*...*..
'¸.•*´¨`*•.¸¸.•*´¨`*•.¸¸.•*´¨`*•.¸¸.•*´¨`*•.¸¸.

    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'
    '“If you made a list of reasons why any couple got married,and another list of the reasons'
    'for their divorce, you'd have a lot of overlapping.”                                     '
    '                     Mignon McLaughlin quotes (American Journalist and Author, 1913-1983)'
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'//PT 7Aug2006: load the listview with the data TaskList (private member of CWindows)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub LoadListView(count As Integer, Optional TaskHwnd As Long)  '//count contains the number of visible windows
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: LoadListView
' Description: load the listview with the data TaskList
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo LoadListViewErr
    
    '\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\'
    '                                                                                                 '
    '                                                                                                 '
    '  ''/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*\*'  '
    '  '“The software is where the magic is. If you're going to have all this power be simple      '  '
    '  'enough, appealing enough and cool enough, it's going to be because the software is right.” '  '
    '  '                                                                               Bill Gates  '  '
    '  '/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*\*/'  '
    '                                                                                                 '
    '                                                                                                 '
    '/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/'
    
    Dim i As Integer
    Dim IsHidden As Long
    Dim WindowIcon As Long '//to get window's icon handle
    
    frmmain.Lstwindows.ListItems.Clear '//Clear the listView
    
    '//To avoid this error in line 10 (next line) "ImageList cannot be modified while another control is bound to it"
    frmmain.Lstwindows.SmallIcons = Nothing
 
    '//To avoid this error in line 20 below "ImageList cannot be modified while another control is bound to it"
10   frmmain.ILWindowsIcons.ListImages.Clear
           
    '//make a fixed size for the icon, if we didn't add the following 2 lines
    '//then the icon size will be larger after any update for the listview "i don't know why"
    '//note that the size of PIco must be 480,480 "bardo i don't know why"
    
    frmmain.ILWindowsIcons.ImageHeight = 16
    
    frmmain.ILWindowsIcons.ImageWidth = 16
        
    '//note that ListView index start from 1
    '//TaskList index start from 0
    For i = 0 To count - 1
                 
        '//why use the following alogrism is here==> http://www.thewonderfulicon.com/faqs.html
        '//try to get icon
        WindowIcon = SendMessage(TaskList(i).TaskID, WM_GETICON, ICON_BIG, 0&)
                    
        If WindowIcon = 0 Then '//try to get icon if the above failed
           WindowIcon = CopyIcon(GetClassLong(TaskList(i).TaskID, GCL_HICONSM))
        End If
                 
        If WindowIcon = 0 Then '//try to get icon if the above failed
           WindowIcon = SendMessage(TaskList(i).TaskID, WM_GETICON, ICON_SMALL, 0&)
        End If
                 
        If WindowIcon = 0 Then '//try to get icon if the above failed
           WindowIcon = CopyIcon(GetClassLong(TaskList(i).TaskID, GCL_HICON))
        End If
                
        If WindowIcon = 0 Then '//If this too returns false, give up and load a default icon by calling LoadIcon
           WindowIcon = LoadIcon(ByVal 0&, IDI_APPLICATION)
        End If
        
        frmmain.PIco.Cls '//if we didn't cls PIco then PIco will display old icons with the new one
        
        DrawIcon frmmain.PIco.hDC, 0, 0, WindowIcon '//Copy the icon to PIco
                 
20      frmmain.ILWindowsIcons.ListImages.Add i + 1, , frmmain.PIco.Image  '//Add the icon to the listview
                   
        frmmain.Lstwindows.SmallIcons = frmmain.ILWindowsIcons '//bind the imagelist to the listview
      
        '//starts with index 1 in the listview and increment
        frmmain.Lstwindows.ListItems.Add , , TaskList(i).TaskName, , i + 1
        '//hence the first item in the listview will have the index 1 and so on
        
        IsHidden = GetProp(TaskList(i).TaskID, "VisibiltyState") 'check if it is hidden window by this software
        If IsHidden = 7878 Then '//if hidden by Me
            Dim res As Integer
            res = GetINIProfileInt("chkGray", 1)
            If res Then
               frmmain.Lstwindows.ListItems.Item(i + 1).ForeColor = HiddenWindowForeColor 'then make the item grayed
            End If
            frmmain.Lstwindows.ListItems.Item(i + 1).SubItems(1) = "Hide"
        Else
            frmmain.Lstwindows.ListItems.Item(i + 1).SubItems(1) = "Show"
        End If
        
        frmmain.Lstwindows.ListItems.Item(i + 1).Tag = TaskList(i).TaskID
        'Store the window's Hwnd in the Tag in order to use in Lstwindows_MouseUp
        
    Next i
    
    
LoadListViewExit:
    On Error Resume Next

    Exit Sub
    
LoadListViewErr:

    '//Templateunique number is a unique number for MyTemplate because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee18"
     
    '//suppose you download a file using Skype software, the window title will change
    '//every second and this causes to call LoadListView much more and may cause this error
    '//at line 20 .. i know that i fixed this issue in GetList function by changeing only
    '//the window that had been changed
   End
    If Err.Number = 7 Then '//Out of memory
        End
    End If
    
GoTo LoadListViewExit

End Sub

'//PT 1Jun2007: added next function to check if new window open or closed and to fill the listview
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub GetList()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: GetList
' Description: GetList Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo GetListErr
    
    '''''''''''''''/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''//“The Linux philosophy is 'Laugh in the face of danger'. Oops. Wrong One.'
    '''''''''''''''//'Do it yourself'. Yes, that's it.”                                      '
    '''''''''''''''                                                           Linus Torvalds '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/'
    
    Dim k, r '//r and k used for the loops below
    Dim iIndex As Integer
    Dim CheckDup As Boolean
    Static Old_i As Integer
    Dim i As Integer '//number of visible windows
      
    i = CVisWindows.FillTaskList '//Fill TaskList struct and get the number of visible windows
        
    '//i represents the nubmer of the current visible windows
    '//Old_i represents the number of the old number of the visible windows before a change occur
    '//if i <> old_i then a change occur hence we 'll call LoadListView
    '//if i == old_i then we 'll check if a change occur while the number of the windows is still the same
    '//for example a title of a window changed
    '//so we 'll loop through Lstwindows items and check for every item in the listitems if it isn't
    '//in Tasklist==>if an item in the listview is in th tasklist structure then we 'll go to check the next
    '//if the items isn't in the tasklist then we'll assign CheckDup to true in order to call LoadListView
    '//now the question is why we made all of this???? well,if we update the listview everysecond then
    '//it will refreshed every second "tr3sh ya3ny" and that's why we make this check.
    If i = Old_i Then
    
       For k = 0 To frmmain.Lstwindows.ListItems.count - 1
       
           For r = 0 To i - 1
                
                If InStr(1, TaskList(r).TaskName, frmmain.Lstwindows.ListItems.Item(k + 1)) = 0 Then
                    
                    CheckDup = True '//Not Found this window in the TaskList
                
                Else
                    
                    '//here 's another case:
                    '//If the user open firefox, then firefox window title will be
                    '//"Mozilla FireFox" however if he opens any site after this the title will change to "Site Name - Mozilla FireFox"
                    '//hence the word "Mozilla FireFox" is still in the title and this will cause instr function to return a velue, if
                    '//InStr returns a value then the title won't change in the next below line altought it is already changed , hence
                    '//i will check here if the len of the window title is larger than its item in the listview by only 1 char hence
                    If Len(TaskList(r).TaskName) - Len(frmmain.Lstwindows.ListItems.Item(k + 1)) = 1 Then
                       '//this case is something like this "Mozilla FireFox " and "Mozilla FireFox"
                       CheckDup = False '//Found this window in the Tasklist
                       Exit For
                    Else
                       '//this case is something like this "Site Name - Mozilla FireFox " and "Mozilla FireFox"
                       CheckDup = True '//Not Found this window in the TaskList
                    End If
                    
                End If
           
           Next r
       
           If CheckDup Then '//if the number of the windows doesn't change and we didn't found the window
              '//then that's means, the window title has been changed
              For iIndex = 0 To 100 '//hence we 'll get that window using its HWND stored in the Tag
                        
                  'Debug.Print frmmain.Lstwindows.ListItems.Item(k + 1).Tag, TaskList(iIndex).TaskID
                    
                    '//the hwnd stored in the tag == to hwnd stored in tasklist
                    If frmmain.Lstwindows.ListItems.Item(k + 1).Tag = TaskList(iIndex).TaskID Then
                          '//then we found the window, let's change its title only
                          frmmain.Lstwindows.ListItems.Item(k + 1).Text = TaskList(iIndex).TaskName
                          
                          Exit Sub
                      
                    End If
                  
              Next iIndex
              
              '//In case the above algoratiom failed, we 'll load the list view
              LoadListView i
              LogError "Load List view because can't found the window"
              
            End If
       
       Next k
    
    Else
        
        LoadListView i
        
        Old_i = i
        
    End If
    
    
GetListExit:
    On Error Resume Next

    Exit Sub
    
GetListErr:

    '//0x0f3ee19 number is a unique number for MyTemplate because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee19"
    
GoTo GetListExit

End Sub
