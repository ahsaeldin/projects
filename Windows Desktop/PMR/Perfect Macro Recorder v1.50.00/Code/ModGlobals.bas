Attribute VB_Name = "ModGlobals"
Option Explicit

    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'
    '“Optimism is the faith that leads to achievement. Nothing can be done without hope and confidence.”'
    '                                                                                       Helen Keller'
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *  '
 


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
Public Const IniFileName = "pmr.ini"

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

    'On Error GoTo GetINIProfileIntErr
        
    Dim Data As Integer
    Dim eotmPath As String
  
    eotmPath = AppPath + IniFileName '//Get the Path of the exe in order to use below in GetPrivateProfileInt
    GetINIProfileInt = GetPrivateProfileInt(SoftwareName, Key, DefaultInt, eotmPath)
    '//The Parameters of GetPrivateProfileInt
    '//Section Name, Key Name, Value to return if key not found, Filename with full path to FoldMon.ini, if the return value = the default value, then the key did not exist
   
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
 '   If OnErrorSysError Then SysError Err, "0x0f3ee13"
    
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
 '   If OnErrorSysError Then SysError Err, "0x0f3ee14"
    
GoTo WriteToINIProfileExit

End Sub
