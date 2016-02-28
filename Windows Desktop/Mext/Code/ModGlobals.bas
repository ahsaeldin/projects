Attribute VB_Name = "ModGlobals"
Option Explicit

    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'
    '“Optimism is the faith that leads to achievement. Nothing can be done without hope and confidence.”'
    '                                                                                       Helen Keller'
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *  '
 
 
'//Function Alias to display to the user whenever an error occur in the  function
'"0x0f3ee11" == FadeForm in ModGlobals.bas
'"0x0f3ee12" == GetINIProfileInt in ModGlobals.bas
'"0x0f3ee13" == GetINIProfileString in ModGlobals.bas
'"0x0f3ee14" == WriteToINIProfile in ModGlobals.bas
'"0x0f3ee116" == txtshortcut_KeyDown in frmoptions.frm
'"0x0f3ee117" == EncodeKey in frmoptions.frm
'"0x0f3ee118" == DecodeKey in frmoptions.frm
'"0x0f3ee119" == CheckShortcutKeys in ModOptions.bas

                                    '''''''''''''''''''''''''''''''''''''''
                                    '___________@@@@@@@@__________@@@@    '
                                    '________@@@________@@_____@@@@@@@    '
                                    '________@@___________@@__@@@______@@ '
                                    '________@@____________@@@__________@@'
                                    '__________@@____A BIG FLOWER_______@@'
                                    '____@@@@@@______@@@@@___________@@   '
                                    '__@@@@@@@@@__@@@@@@@_________@@      '
                                    '__@@____________@@@@@@@@_______@@    '
                                    '_@@____________@@@@@@@@@_____@@      '
                                    '_@@____________@@@@@@@@___@@@        '
                                    '_@@@___________@@@@@@@______@@       '
                                    '__@@@@__________@@@@@_________@@     '
                                    '____@@@@@@______FOR ____________@@   '
                                    '_________@@_______YOU______________@@'
                                    '________@@___________@@___________@@ '
                                    '________@@@________@@@@@@@@@@@       '
                                    '_________@@@_____@@@_@@@@@@@@        '
                                    '__________@@@@@@@@ @@@@@@            '
                                    '___________@@@@@_@                   '
                                    '____________________@                '
                                    '_____________________@               '
                                    '______________________@              '
                                    '_______________________@             '
                                    '_________________@_________@@@       '
                                    '_______________@@@@__@__@____@@      '
                                    '_____________@_______@@@____@@       '
                                    '______________@@@@____@__@@          '
                                    '______________________@              '
                                    '''''''''''''''''''''''''''''''''''''''
                                    
'PT 22Nov2006: added next 3 lines in support of make a global constants for
'Email address, Software Name and SiteName in case that any of them changes not
'to change in the whole Project.
Public Const Email = "support@perfectiontools.com"
Public Const SoftwareName = "Easy Text Inserter"
Public Const SiteName = "http://www.perfectiontools.com/easy-text-inserter.html"
Public Const OrderPage = "http://www.perfectiontools.com/purchase_4.html"

Public Const IniFileName = "eti.ini"
'//pt 8Nov2007: added next line to make a global varibale for the foreground color of the
'//hidden window in the listview in order to change it as we want later
Public HiddenWindowForeColor As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'//pt 12June2007: we used this const for GetClassLong, it is not the type lib
Public Const GCL_HICONSM = -34

'//pt 29May2007: added next 2 lines
Public InnerSysError As Boolean '//to enable or disable syserror message inside the function body
Public OnErrorSysError As Boolean '//to enable or disable syserror message in the functions end
'//pt 30May2007: added next line
Public IsLogError As Boolean '//to enable or disable logging to error log

'//PT 22Nov2006: added next line in support of FadeForm Function that will fade any Form
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
  
'//Added because there is a problem addind using SetFocus from the type lib
Public Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
 
'//This is an abstract function that won't be used, it is just a prototype
'To use This Template replace the word Template in the function bode with the apporpate words
'//PT 7Aug2006: added next function as a template for a function with error handling
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub MyTemplate()
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

    On Error GoTo TemplateErr
        
    'add your Code Here
    
TemplateExit:
    On Error Resume Next

    Exit Sub
    
TemplateErr:

    '//Templateunique number is a unique number for MyTemplate because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "Templateunique number"
    
GoTo TemplateExit

End Sub

'//PT 22Dec2006: added next function as our Public SysError function that will report an error both log and msgbox
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SysError(syerr As ErrObject, FunctionAlias As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
  
    '//added this commented clause as an exmple of how to use FormatMessage in displaying LastDllErrorCode
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    'If Err.LastDllError <> 0 Then
    'Dim lastError As Long
    'Dim msgbuf As String
    'Dim i As Long
    '  lastError = Err.LastDllError
    '  msgbuf = Space(256)
    '  i = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lastError, 0, _
    '      msgbuf, 256, 0)
    '  MsgBox "Error in API: " & Left(msgbuf, i)
    'End If
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    
    ''How to use this function
    '//0x0ffee13 is a unique number for GetListedIndexMacroInTheCombo because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    'SysError Err, "0x0ffee13"
    ''
        Dim sMsg As String
        Dim EmailMessage As String
    
        With syerr
            sMsg = .Description & vbCrLf
            EmailMessage = Date$ & ":" & Time$ & vbCrLf
            EmailMessage = EmailMessage & .Description & vbCrLf
            sMsg = sMsg & "You can email " & Email & " and report this error condition:" & vbCrLf
            sMsg = sMsg & "Error: " & .Number & vbCrLf
            EmailMessage = EmailMessage & "Error: " & .Number & vbCrLf
            sMsg = sMsg & "Des: " & .Description & vbCrLf
            EmailMessage = EmailMessage & "Des: " & .Description & vbCrLf
            '//CL means code location, FunctionAlias == Function Name
            sMsg = sMsg & "version " & "(" & App.Major & "." & App.Minor & "." & App.Revision & ")" & " CL" & ":: " & FunctionAlias & vbCrLf
            EmailMessage = EmailMessage & "version " & "(" & App.Major & "." & App.Minor & "." & App.Revision & ")" & " CL" & ":: " & FunctionAlias & vbCrLf
            sMsg = sMsg & "S: " & .Source & vbCrLf
            EmailMessage = EmailMessage & "S: " & .Source & vbCrLf
            '//LDE means Last DLL Error
            sMsg = sMsg & "LDE: " & .LastDllError & vbCrLf
            EmailMessage = EmailMessage & "LDE: " & .LastDllError & vbCrLf
            Clipboard.Clear
            Clipboard.SetText EmailMessage
            sMsg = sMsg & "(You can paste this entire message by pressing Ctrl+V now in your e-mail software)" & vbCrLf
            sMsg = sMsg & vbCrLf & vbCrLf
            sMsg = sMsg & "Thank you for you patience." & vbCrLf
            MsgBox sMsg, vbOKOnly + vbExclamation, "Error"
        End With
        
        'add error to log file
        Dim iFreeFile As Integer
        iFreeFile = FreeFile
        If Right$(App.path, 1) = "\" Then
            Open App.path & App.Title & ".log" For Append As #iFreeFile
        Else
            Open App.path & "\" & App.Title & ".log" For Append As #iFreeFile
        End If
        Print #iFreeFile, Now, App.Title, syerr.Number, syerr.Description, syerr.Source
        Close #iFreeFile

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub LogError(error As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If Not IsLogError Then Exit Sub 'only log if IsLogError is true
    'add error to log file
    Dim iFreeFile As Integer
    iFreeFile = FreeFile
    If Right$(App.path, 1) = "\" Then
        Open App.path & App.Title & ".log" For Append As #iFreeFile
    Else
        Open App.path & "\" & App.Title & ".log" For Append As #iFreeFile
    End If
    Print #iFreeFile, Now, error
    Close #iFreeFile
    
End Sub

'//PT 22Dec2006: added next function to Fade any Form
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub FadeForm(FadeWhat As Form)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: FadeForm
' Description: FadeForm Description
' Author: Ahmed Saad
' Date: 22/12/06 5:51:29 PM
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo FadeFormErr
        
    Const LWA_ALPHA = &H2
    Const WS_EX_LAYERED = &H80000
    Dim Ret As Long
    Dim I As Integer
    'Set the window style to 'Layered'
    Ret = GetWindowLong(FadeWhat.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    'Set the opacity of the layered window to 128
    For I = 0 To 51
          SetWindowLong FadeWhat.hwnd, GWL_EXSTYLE, Ret
          On Error Resume Next
          SetLayeredWindowAttributes FadeWhat.hwnd, vbBlue, I * 5, LWA_ALPHA
          If I = 1 Then
              FadeWhat.Visible = True
              'you must refresh frmmain to avoid blackness
              FadeWhat.Refresh
          End If
    Next I
    
    WriteToINIProfile "Fade", 1
FadeFormExit:
    On Error Resume Next

    Exit Sub
    
FadeFormErr:

    '//FadeFormunique number is a unique number for FadeForm because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee11"
    
GoTo FadeFormExit

End Sub

'//PT 22Dec2006: Use the function to get the Dir path of the executable file
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function AppPath()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    If Right(App.path, 1) <> "\" Then AppPath = App.path + "\" Else AppPath = App.path
   
End Function

'//PT: 22Dec2006: added entire function for getting integers from IniFileName const above
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
Public Function GetINIProfileInt(Key As String, DefaultInt As Integer) As Integer
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
'============================================================

' Routine Name: GetINIProfileInt
' Description: Use this function to get an integer value from eotm.ini
' Author: Ahmed Saad
' Date: 22/12/06 6:30:29 PM
' Copyright © 2006 - 2007, - 2007  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo GetINIProfileIntErr
        
    Dim Data As Integer
    Dim eotmPath As String
  
    eotmPath = AppPath + IniFileName '//Get the Path of the exe in order to use below in GetPrivateProfileInt
    GetINIProfileInt = GetPrivateProfileInt(SoftwareName, Key, DefaultInt, eotmPath)
    '//The Parameters of GetPrivateProfileInt
    '//Section Name, Key Name, Value to return if key not found, Filename with full path to FoldMon.ini, if the return value = the default value, then the key did not exist
  
    
GetINIProfileIntExit:
    On Error Resume Next

    Exit Function
    
GetINIProfileIntErr:

    '//GetINIProfileIntunique number is a unique number for GetINIProfileInt because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee12"
    
GoTo GetINIProfileIntExit

End Function

'//PT: 22Dec2006: added entire function for getting stings from IniFileName const above
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
Public Function GetINIProfileString(Key As String, DefaultString As String, Data As String, DataSize As Integer) As Integer
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
'============================================================

' Routine Name: GetINIProfileString
' Description: Use this function to get a strings from FoldMon.ini
' Author: Ahmed Saad
' Date: 22/12/06 6:45:29 PM
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo GetINIProfileStringErr
        
    '//example of using this function
    '//Dim dkey As String * 11
    '//GetINIProfileString "Show All Windows", "1,1,0,32,", dkey, Len(dkey)

    Dim eotmPath As String
  
    eotmPath = AppPath + IniFileName '//Get the Path of My Folders Monitor in order to use below in GetPrivateProfileInt
    GetINIProfileString = GetPrivateProfileString(SoftwareName, Key, DefaultString, Data, DataSize, eotmPath)
   'The Parameters of GetPrivateProfileString
   'Section name, Key Name, Default String, Destination Buffer, Size of Destination Buffer , Filename with full path to FoldMon.ini

    
GetINIProfileStringExit:
    On Error Resume Next

    Exit Function
    
GetINIProfileStringErr:

    '//GetINIProfileStringunique number is a unique number for GetINIProfileString because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee13"
    
GoTo GetINIProfileStringExit

End Function

'//PT: 22Dec2006: added entire function for writing data to IniFileName const above
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
Public Sub WriteToINIProfile(Key As String, Data As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
'============================================================

' Routine Name: WriteToINIProfile
' Description: WriteToINIProfile Description
' Author: Ahmed Saad
' Date: 22/12/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo WriteToINIProfileErr
        
    Dim eotmPath As String
 
    eotmPath = AppPath + IniFileName '//Get the Path of the exe in order to use below in GetPrivateProfileInt
    WritePrivateProfileString SoftwareName, Key, Data, eotmPath
   'The Parameters of WritePrivateProfileString
   'Section name, Key Name, Destination Buffer, Filename with full path to FoldMon.ini
    
WriteToINIProfileExit:
    On Error Resume Next

    Exit Sub
    
WriteToINIProfileErr:

    '//WriteToINIProfileunique number is a unique number for WriteToINIProfile because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee14"
    
GoTo WriteToINIProfileExit

End Sub
