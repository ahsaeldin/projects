Attribute VB_Name = "ModOptions"
Option Explicit
                
                                '(¯`v´¯)
                                '`*.¸.*´
                                '¸.´¸.*¨) ¸.*¨)
                                '(¸.´ (¸.´ .´ ¸¸.¨¯`
                                '_____****__________**** ______
                                '___***____***____***__ *** ___
                                '__***________****_______***___
                                '_***__________**_________***__
                                '_***_______ JUST _________***_
                                '_***________ TO___________***_
                                '__***_______SAY__________***__
                                '___***______ I ________***___
                                '____***_____ LOVE___***_____
                                '______***___YOU___***_______
                                '________***_______***_________
                                '__________***___***___________
                                '____________*****_____________
                                '_____________***______________
                                '______________*_______________

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'If you can't make it good, at least make it look good.'
                '                                        Bill Gates Quote'
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'//To now how to add a new shortcut, Search the whole project for '//ShortCut How To
'//And don't forget to Hit Reset Defaults Button after any change you made to the shortcut scheme

'//add the new shortcut variable here
'//pt 16June2007:add as a global variable for Defaults ShortCut keys names in case i want to change it later
'//pt 1July2007: commented next line to Delete Revert Shortcut key
'//Public Const RevertWindowsVisibilty As String = "Revert"
Public Const My_Macros As String = "My Macros"
Public Const OptionsWin As String = "Options"
Public Const Pause_Play_back As String = "Pause Play back"
Public Const Pause_Recording As String = "Pause Recording"
Public Const Resume_Play_back As String = "Resume Play back"
Public Const Resume_Record As String = "Resume Recording"
Public Const Show_Perfect_Macro_Recorder As String = "Show Perfect Macro Recorder"
Public Const Start_Recording As String = "Start Recording"
Public Const Stop_Play_back As String = "Stop Play back"
Public Const Stop_Record As String = "Stop Recording"

'//pt 1July2007: commented next line to Delete Show Foreground Window Shortcut key
'//Public Const Show_Foreground_Window As String = "Show Foreground Window"

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CheckOptionsShortcutKeys()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

   Dim i As Integer
   Dim iIndex As Integer
   Dim OptionHwnd As Long
   Dim OptionParent As Long
   Dim DecodedKey As String
   Dim PressedKey As Integer
   Dim OptionsWindowOwner As Long
   Dim key1 As String, key2 As String
   Dim MyItemInTheListView As ListItem
   Dim dkey As String * 15  '//why 15, coz it larger than "1,1,1#256*
   Dim DefChar1 As String
   '//pt 1July2007: commented next line to Delete Revert Shortcut key
   '//Dim DefChar2 As String
   Dim DefChar3 As String
   Dim DefChar4 As String
   Dim DefChar5 As String
   '//pt 1July2007: commented next line to Delete Show Foreground Window Shortcut key
   '//Dim DefChar6 As String
   Dim DefChar7 As String
   Dim DefChar8 As String
   Dim DefChar9 As String
   Dim DefChar10 As String
   Dim DefChar11 As String
   Dim DefChar12 As String
 
   'Dim DefCharX As String '//ShortCut How To
   
   '//added Next For loop to Disable ShortCut keys if frmOptions2 is Displayed
   
   Dim OptionsHwnd As Long
   
   OptionsHwnd = FindWindow(vbNullString, "Perfect Macro Recorder Options")
   
   If OptionsHwnd Then
      If IsWindowVisible(OptionsHwnd) Then
      '  Debug.Print "Options Is Displayed"
        Exit Sub
      End If
   End If
   
   PressedKey = GetPressedKey

   If GetAsyncKeyState(vbKeyControl) Then
      DecodedKey = 1
   Else
      DecodedKey = 0
   End If
   
   If GetAsyncKeyState(18) Then
      DecodedKey = DecodedKey & "," & 1
   Else
      DecodedKey = DecodedKey & "," & 0
   End If
   
   '//Tell Me why the space after the hash "# " ??
   '//for unknown reason, It added automatic in LoadShortcutKeys
   '//[i mean in GetINIProfileString lines in LoadShortcutKeys  ] and i don't know why
   '//and that's why i added it here ino order to match the keys
   
   If GetAsyncKeyState(vbKeyShift) Then
      DecodedKey = DecodedKey & "," & 1 & "# " & PressedKey & "*"
   Else
      DecodedKey = DecodedKey & "," & 0 & "# " & PressedKey & "*"
   End If
   
   'Debug.Print DecodedKey, Chr(PressedKey), PressedKey
   
   DefChar1 = Str(vbKeyF9) '//Get the ASC value of F9
   '//pt 1July2007: commented next line to Delete Revert Shortcut key
   '//DefChar2 = Str(Asc("R")) '//Get the ASC value of R
   DefChar3 = Str(Asc("O")) '//Get the ASC value of O
   DefChar4 = Str(vbKeyPause)  '//Get the ASC value of H
   DefChar5 = Str(Asc("M")) '//Get the ASC value of M
   '//pt 1July2007: commented next line to Delete Show Foreground Window Shortcut key
   '//DefChar6 = Str(Asc("F")) '//Get the ASC value of F
   DefChar7 = Str(vbKeyF11)  '//Get the ASC value of N
   DefChar8 = Str(vbKeyInsert)  '//Get the ASC value of X
   DefChar9 = Str(vbKeyPause)  '//Insert
   DefChar10 = Str(Asc("S")) '//Get the ASC value of S
   DefChar11 = Str(vbKeyInsert) '//Insert
   DefChar12 = Str(vbKeyF10)
   
   '//DefCharX = Str(Asc(" ")) '//Get the ASC value of Space '//ShortCut How To
   
   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString Stop_Record, "0,1,0#" & DefChar12 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
         If frmJMacro.cmdRecord.Caption = "&Stop Recording" Then
            '//close ballon if it is displayed
            frmJMacro.TrayIcon1.CloseBalloon
            '//pt 28Aug2007: commented next line coz i will use frmoption2 and the new Shortcut Model
            frmJMacro.cmdRecord_Click
          End If
          Debug.Print Stop_Record
      End If
   End If
      
   Dim res As Integer
   
   res = GetINIProfileInt("DASK", 0)
   
   If res Then
      
      '//Shortcut Keys are disableds
      Exit Sub
     
   End If
     
   dkey = ""
   
   '//Get the value of OptionsWin stored in the ini file
   GetINIProfileString OptionsWin, "1,1,0#" & DefChar3 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then '//if it is not a blank field
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
         frmOptions2.Show , frmJMacro
         Debug.Print OptionsWin
      End If
   End If
   '//Wait a second, what is the blank field, we reterived a blank vlaue from ini if the user
   '//reassign a shortcut key and use a shortcut key used for another function before then the old function
   '//shortcut key is got blank

   '//pt 1July2007: commented next code clause to Delete Revert Shortcut key
   '//dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   '//GetINIProfileString RevertWindowsVisibilty, "1,1,0#" & DefChar2 & "*", dkey, Len(dkey)
   '//If InStr(1, dkey, "#") <> 0 Then
   '//   key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
   '//   key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
   '//   If key1 = key2 Then
   '//      frmmain.CmdRevert_Click
   '//      Debug.Print RevertWindowsVisibilty
   '//   End If
   '//End If
      
   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString My_Macros, "1,1,0#" & DefChar5 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
         frmManage.Show
         Debug.Print My_Macros
      End If
   End If
   
   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString Pause_Play_back, "0,0,0#" & DefChar4 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
         'Debug.Print Pause_Play_back
      End If
   End If
   
   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString Start_Recording, "0,1,0#" & DefChar1 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
         If frmJMacro.cmdRecord.Caption = "&Record" Then '//only if not recording
            frmJMacro.cmdRecord_Click
         End If
         Debug.Print Start_Recording
      End If
   End If
   
   '//dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   '//GetINIProfileString Show_Foreground_Window, "1,1,0#" & DefChar6 & "*", dkey, Len(dkey)
   '//If InStr(1, dkey, "#") <> 0 Then
   '//   key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
   '//   key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
   '//   If key1 = key2 Then
   '//      For i = 1 To frmmain.Lstwindows.ListItems.count
   '//          If frmmain.Lstwindows.ListItems(i).Tag = GetForegroundWindow Then
   '//             frmmain.Lstwindows.ListItems(i).Checked = True
   '//             frmmain.cmdshowwindow_Click
   '//             frmmain.Lstwindows.ListItems(i).Checked = False
   '//          End If
   '//      Next i
   '//      Debug.Print Show_Foreground_Window
   '//   End If
   '//End If
  
   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString Stop_Play_back, "0,1,0#" & DefChar7 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
          If IsPlayingMode Then
            'Debug.Print Stop_Play_back
          End If
      End If
   End If
   
   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString Pause_Recording, "0,0,0#" & DefChar8 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
         'Debug.Print Pause_Recording
      End If
   End If
   
   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString Resume_Play_back, "0,1,0#" & DefChar9 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
         'frmmain.cmdSysShowAll_Click
         'Debug.Print Resume_Play_back
      End If
   End If
   
   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString Show_Perfect_Macro_Recorder, "1,1,0#" & DefChar10 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
          Dim JHwnd As Long
          JHwnd = IsWindowVisible(frmJMacro.hWnd)
          If JHwnd = 0 Then
             frmJMacro.Show
             frmJMacro.WindowState = 0
             frmJMacro.SetFocus
          End If
          Debug.Print Show_Perfect_Macro_Recorder
      End If
   End If
   
   dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   GetINIProfileString Resume_Record, "0,1,0#" & DefChar11 & "*", dkey, Len(dkey)
   If InStr(1, dkey, "#") <> 0 Then
      key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
      key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
      If key1 = key2 Then
          'Debug.Print Resume_Record
      End If
   End If
      
   '//ShortCut How To/////////////////////////////////////////////////////////////////////////
   '$$$dkey = "" '//we must clear dkey otherwise it may hold the value of the previous use
   '//and this will cause a bug in displaying the key
   '//for example: if a function has no key, then it will be stored as a blank value in ini
   '//and if dkey has this value from the prevoius use "1,1,1#256*" , then
   '//it will be "1,1#256*"
   '$$$GetINIProfileString Key Name here, "1,1,0#" & DefCharX & "*", dkey, Len(dkey)
   '$$$If InStr(1, dkey, "#") <> 0 Then
   '$$$  key1 = Mid(dkey, 1, InStr(1, dkey, "*"))
   '$$$  key2 = Mid(DecodedKey, 1, InStr(1, DecodedKey, "*"))
   '$$$  If key1 = key2 Then
   '$$$  Debug.Print Key Name here
   '$$$   End If
   '$$$End If
   '/////////////////////////////////////////////////////////////////////////////////////////////
   
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetPressedKey() As Integer
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Dim cnt As Integer

    Dim IgnoreKey As Boolean
    
    '//Well, to avoid process keys like Ctrl alt shift, we start the loop from 2o to 256
    '//however the TAB keycode is 9 hence i will process the TAB key indviudaly here
    If GetKeyState(9) And &H8000 Then
          GetPressedKey = 9
          Exit Function
    End If
                 
    If GetKeyState(vbKeyPause) And &H8000 Then
          GetPressedKey = vbKeyPause
          Exit Function
    End If
    
    '//for checkstopkey function becuase while playbck GetKeyState doesn't work
    If GetAsyncKeyState(vbKeyPause) Then
           GetPressedKey = vbKeyPause
          Exit Function
    End If
    
    
   ' frmJMacro.Caption = 4
    '//note that we must ignore these keys
    '//Left Control = 162
    '//Right Control = 163

    '//Left Shift 160
    '//Right Shift 161

    '//Left Alt = 164
    '//Righ Alt = 165
    For cnt = 20 To 256
       
       'Get the keystate of a specified key
       '//pt 5July2007: commented next line because:
       '//if you assign a single key like "q" to any of hiding shortcuts in frmOptions2 then ...
       '//after you press cmdOk, the hiding commend will be executed although the user doesn't press
       '//that key, but this is because of using "GetAsyncKeyState" however if we use "GetKeyState(cnt) And &H8000"
       '//then this problem won't happen
       If GetAsyncKeyState(cnt) <> 0 Then
       '//If GetKeyState(cnt) And &H8000 Then
          If cnt = 162 Or cnt = 163 Or cnt = 160 Or cnt = 161 Or cnt = 164 Or cnt = 165 Then
             IgnoreKey = True
          End If
          
          GetPressedKey = cnt
          
          '// if any of the above key then don't Exit For
          '//in other words if key = left control "162", then don't exit for because may be the user
          '//already press a key is larger than 162 like "`" or "-" keys
          '//so the questions is , why i added the next line?
          '//to process keys like "`" or "-" which thier keycods is larger than left control "162"
          If Not IgnoreKey Then Exit For
          
       End If
       
    Next cnt
      
End Function

'//Used above in LoadShortcutKeys to translate the key stored in ini file to a value like this "Ctrl + D"
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'//Why we added KeyCode Parameter and why make it optional and why byref??
'//A way to extract only the asc value only form Key Parameter,
'//we make it an optional, coz i added this para after writing this function
'//and i didn't want to change the function calls in the whole module
'//also coz i want to get the keycode only when i need to do this.
'//i make byref, in order to pass an empty string and got the key in it [Out parameter ya3ny]
Public Function DecodeKey(Key As String, Optional ByRef KeyCode As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: Template
' Description: Template Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

            
    '//because the value stored in ini file is like this  1,1,1#256*
    '//hence we have four values
    '//"1,1,1#256*
    '//Ctrl State,Alt State,Shift State#key code*
    
    Dim ResolvedKey As String
    
    Dim part1, part2, part3, part4, Where#, WhereStar
    
    If InStr(1, Key, "#") = 0 Then Exit Function
    
    part1 = IIf(Left(Key, 1) = 1, "Ctrl", "") '//check if the key contains Ctrl
    
    part2 = IIf(Mid(Key, 3, 1) = 1, "Alt", "")  '//check if the key contains Alt
    
    part3 = IIf(Mid(Key, 5, 1) = 1, "Shift", "") '//check if the key contains Shift
    
    '//Get the key code:
    
    '//get the first delimiter
    Where# = InStr(1, Key, "#")
    '//get the second delimiter
    WhereStar = InStr(1, Key, "*")
    
    part4 = Chr(Mid(Key, Where# + 1, WhereStar - Where# - 1)) '//get the asc value by start copy
    '// the first char of it with length equal to (delimiter 2 - delimiter 1 -1)
    
    If part4 = " " Then part4 = "Space" '//if the asc value is equal to "Space"
    
    '//Why the next line ?? and why i didn't assign KeyCode == part4??
    '//To exract the asc value from Key, and i didn't assign keycode to part4 coz
    '//i need the asc value not chr value
    KeyCode = Mid(Key, Where# + 1, WhereStar - Where# - 1)
    
     '//check if the key is like F1 and F2 an so on
    part4 = TranslateShortCutKey(Int(KeyCode))
    
    '//if TranslateShortCutKey returns none
    If part4 = "none" Then
       part4 = Chr(Mid(Key, Where# + 1, WhereStar - Where# - 1)) '//get the asc value by start copy
    End If
    
    '//assemple the DecodeKey
    DecodeKey = part1 & IIf(part1 <> "", " + ", "") & part2 & IIf(part2 <> "", " + ", "") _
    & part3 & IIf(part3 <> "", " + ", "") & part4
    '//Note this means IIf(part1 <> "", " + ", "" : if there is a value of part1 then i 'll put the "+" otherwise i won't put it

End Function

'//pt 24June2007: Added next function to convert a shortcut key from a form like this
'//Ctrl + Shift + S to a encoded form like this "1,1,0#83*
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function EncodeKey(Key As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
    '//This function alias is 0x0f3ee117
    Dim Res1 As Integer
    Dim Res2 As Integer
    Dim KeyChr As String
    Dim EnKeyBuf As String
    Dim TranslatedKey As String
     
    '//Note that txtshortcut.Tag contains only the KeyCode of the Shortcut being Displayed in it
    '//Ctrl + S , then it contains the asc("S") only, if Ctrl + then it contains asc("Ctrl") only
    If frmOptions2.txtshortcut.Tag = "" Then
       'LogError "Can't Get Key Code"
       Err.Description = "Can't Get Key Code"
       SysError Err, "EncodeKey"
       Exit Function
    End If
    
    '//dont forget to write a comment about why  you chooses the
    '//word TraslatedKey
    
    '//Note that txtshortcut.Tag contains only the KeyCode of the Shortcut being Displayed in it
    '//Ctrl + S , then it contains the asc("S") only, if Ctrl + then it contains asc("Ctrl") only
    KeyChr = Chr(frmOptions2.txtshortcut.Tag)
    TranslatedKey = TranslateShortCutKey(frmOptions2.txtshortcut.Tag)
    Res1 = InStr(1, Key, KeyChr)
    Res2 = InStr(1, Key, TranslatedKey)
    
    If Res1 = 0 Then
       If Res2 = 0 Then
          'LogError "KeyCode Error, " & "Key = " & Key & ", KeyChar = " & KeyChr & " ,TranslatedKey = " & TranslatedKey
          Err.Description = "KeyCode Error"
          SysError Err, "EncodeKey"
          Exit Function
       End If
    End If
    
    '//Create the Encoded Key and transform it from something like this Ctrl+S to "1,0,0# 83*"
    If InStr(1, Key, "Ctrl") Then
       EnKeyBuf = 1
    Else
       EnKeyBuf = 0
    End If
    
    If InStr(1, Key, "Alt") Then
       EnKeyBuf = EnKeyBuf & "," & 1
    Else
       EnKeyBuf = EnKeyBuf & "," & 0
    End If
 
    '//Note that txtshortcut.Tag contains only the KeyCode of the Shortcut being Displayed in it
    '//Ctrl + S , then it contains the asc("S") only, if Ctrl + then it contains asc("Ctrl") only
    '//.....
    '//Tell Me why the space after the hash "# " ??
    '//for unknown reason, It added automatic in LoadShortcutKeys
    '//[i mean in GetINIProfileString lines in LoadShortcutKeys  ] and i don't know why
    '//and that's why i added it here ino order to match the keys
    If InStr(1, Key, "Shift") Then
       EnKeyBuf = EnKeyBuf & "," & 1 & "# " & frmOptions2.txtshortcut.Tag & "*"
    Else
       EnKeyBuf = EnKeyBuf & "," & 0 & "# " & frmOptions2.txtshortcut.Tag & "*"
    End If
 
    EncodeKey = EnKeyBuf
    
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TranslateShortCutKey(KeyCode As Integer) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    '//take some keycodes like [F1 and F2 and F2] and convert it to its meaning like F1
    
    If KeyCode = vbKeyF1 Then
        TranslateShortCutKey = "F1"
    ElseIf KeyCode = vbKeyF2 Then
        TranslateShortCutKey = "F2"
    ElseIf KeyCode = vbKeyF3 Then
        TranslateShortCutKey = "F3"
    ElseIf KeyCode = vbKeyF4 Then
        TranslateShortCutKey = "F4"
    ElseIf KeyCode = vbKeyF5 Then
        TranslateShortCutKey = "F5"
    ElseIf KeyCode = vbKeyF6 Then
        TranslateShortCutKey = "F6"
    ElseIf KeyCode = vbKeyF7 Then
        TranslateShortCutKey = "F7"
    ElseIf KeyCode = vbKeyF8 Then
        TranslateShortCutKey = "F8"
    ElseIf KeyCode = vbKeyF9 Then
        TranslateShortCutKey = "F9"
    ElseIf KeyCode = vbKeyF10 Then
        TranslateShortCutKey = "F10"
    ElseIf KeyCode = vbKeyF11 Then
        TranslateShortCutKey = "F11"
    ElseIf KeyCode = vbKeyF12 Then
        TranslateShortCutKey = "F12"
    ElseIf KeyCode = vbKeyInsert Then
        TranslateShortCutKey = "Insert"
    ElseIf KeyCode = vbKeyDelete Then
        TranslateShortCutKey = "Delete"
    ElseIf KeyCode = 192 Then ' ` key
        TranslateShortCutKey = "`"
    ElseIf KeyCode = 189 Then
        TranslateShortCutKey = "-"
    ElseIf KeyCode = 187 Then
        TranslateShortCutKey = "="
    ElseIf KeyCode = vbKeyHome Then
        TranslateShortCutKey = "Home"
    ElseIf KeyCode = vbKeyPageDown Then
        TranslateShortCutKey = "Page Down"
    ElseIf KeyCode = vbKeyPageUp Then
        TranslateShortCutKey = "Page Up"
    ElseIf KeyCode = vbKeyPause Then
        TranslateShortCutKey = "Pause"
    ElseIf KeyCode = 219 Then
        TranslateShortCutKey = "["
    ElseIf KeyCode = 221 Then
        TranslateShortCutKey = "]"
    ElseIf KeyCode = 220 Then
        TranslateShortCutKey = "\"
    ElseIf KeyCode = 186 Then
        TranslateShortCutKey = ";"
    ElseIf KeyCode = 222 Then
        TranslateShortCutKey = "'"
    ElseIf KeyCode = vbKeyEnd Then
        TranslateShortCutKey = "End"
    ElseIf KeyCode = 226 Then
        TranslateShortCutKey = "\"
    ElseIf KeyCode = 188 Then
        TranslateShortCutKey = ","
    ElseIf KeyCode = 190 Then
        TranslateShortCutKey = "."
    ElseIf KeyCode = 191 Then
        TranslateShortCutKey = "/"
    ElseIf KeyCode = vbKeyUp Then
        TranslateShortCutKey = "Up"
    ElseIf KeyCode = vbKeyDown Then
        TranslateShortCutKey = "Down"
    ElseIf KeyCode = vbKeyRight Then
        TranslateShortCutKey = "Right"
    ElseIf KeyCode = vbKeyLeft Then
        TranslateShortCutKey = "Left"
    ElseIf KeyCode = vbKeyTab Then
        TranslateShortCutKey = "Tab"
    ElseIf KeyCode = vbKeySpace Then
        TranslateShortCutKey = "Space"
    ElseIf KeyCode = vbKeyNumlock Then
        TranslateShortCutKey = "Num lock"
    ElseIf KeyCode = vbKeyNumpad0 Then
        TranslateShortCutKey = "Num 0"
    ElseIf KeyCode = vbKeyNumpad1 Then
        TranslateShortCutKey = "Num 1"
    ElseIf KeyCode = vbKeyNumpad2 Then
        TranslateShortCutKey = "Num 2"
    ElseIf KeyCode = vbKeyNumpad3 Then
        TranslateShortCutKey = "Num 3"
    ElseIf KeyCode = vbKeyNumpad4 Then
        TranslateShortCutKey = "Num 4"
    ElseIf KeyCode = vbKeyNumpad5 Then
        TranslateShortCutKey = "Num 5"
    ElseIf KeyCode = vbKeyNumpad6 Then
        TranslateShortCutKey = "Num 6"
    ElseIf KeyCode = vbKeyNumpad7 Then
        TranslateShortCutKey = "Num 7"
    ElseIf KeyCode = vbKeyNumpad8 Then
        TranslateShortCutKey = "Num 8"
    ElseIf KeyCode = vbKeyNumpad9 Then
        TranslateShortCutKey = "Num 9"
    ElseIf KeyCode = vbKeyNumpad9 Then
        TranslateShortCutKey = "Num 9"
    ElseIf KeyCode = 107 Then '//"Num +" Key
        TranslateShortCutKey = "Num +"
    ElseIf KeyCode = 109 Then '//"Num -" Key
        TranslateShortCutKey = "Num -"
    ElseIf KeyCode = 106 Then '//"Num *" Key
        TranslateShortCutKey = "Num *"
    ElseIf KeyCode = 111 Then '//"Num /" Key
        TranslateShortCutKey = "Num /"
    ElseIf KeyCode = 110 Then '//"Num ." Key
        TranslateShortCutKey = "."
    Else '//if a normal key then return none
        TranslateShortCutKey = "none"
    End If
  
End Function
