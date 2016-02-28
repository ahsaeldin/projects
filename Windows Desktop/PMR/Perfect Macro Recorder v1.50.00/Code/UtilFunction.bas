Attribute VB_Name = "UtilFunction"
Option Explicit
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'//cprin 7Aug2006: Added this Module to hold all utilites function to be a public througout the project
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'//cprin 28Aug2006: added next Consts to use as a small icon for the listview in frmEdit
Public Const DelayIcon = 3
Public Const MouseIcon = 1
Public Const KeyboardIcon = 2

'//cprin 28Aug2006: added next Consts to use as a colors for the events in the listview in frmEdit
Public Const MouseEventColor = vbMagenta
Public Const KeyboradEventColor = 6523464
Public Const DelayEventColor = vbBlack

'//cprin 7Aug2006: added next function as our Public SysError function that will report an error both log and msgbox
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SysError(syerr As ErrObject, FunctionAlias As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim sMsg As String
    Dim EmailMessage As String

    With syerr
        sMsg = .Description & vbCrLf
        EmailMessage = Date$ & ":" & time$ & vbCrLf
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
    If right$(App.Path, 1) = "\" Then
        Open App.Path & App.Title & ".log" For Append As #iFreeFile
    Else
        Open App.Path & "\" & App.Title & ".log" For Append As #iFreeFile
    End If
    Print #iFreeFile, App.Title, syerr.Number, syerr.Description, syerr.Source, Now
    Close #iFreeFile

End Sub
'//cprin 27Aug2006: added next function to get the current selected macro in the combo box and
'// we can return the macro name in combo in the MacName Parameter
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetListedIndexMacroInTheCombo(Optional MacName As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: GetListedIndexMacroInTheCombo
' Description: Get the current selected macro in the combo box
' Author: Ahmed Saad
' Date: 27/08/06 1:51:29 ã
' Copyright © 2006  Perfection Tools Software
' Notes: Use this Function any where in the code where you want to get the current selected macro in the combo box

' Modification History:

'============================================================

    On Error GoTo GetListedIndexMacroInTheComboErr
        
    Dim SelectedMacroName As String
    'add your Code Here
    '//cprin 27Aug2006: added next line to get the selected macro Path in the combobox
    SelectedMacroName = AppPath & "Data\" & frmJMacro.CobMacros.List(frmJMacro.CobMacros.ListIndex) & ".pmf"
    MacName = frmJMacro.CobMacros.List(frmJMacro.CobMacros.ListIndex)  '//Copy Macro Name to Mac Name function
    GetListedIndexMacroInTheCombo = SelectedMacroName '//Return the Path

GetListedIndexMacroInTheComboExit:
    On Error Resume Next

    Exit Function
    
GetListedIndexMacroInTheComboErr:

    '//0x0ffee13 is a unique number for GetListedIndexMacroInTheCombo because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    SysError Err, "0x0ffee13"
    
GoTo GetListedIndexMacroInTheComboExit

End Function


'//This is an abstract function that won't be used, it is just a prototype
'To use This Template replace the word Template in the function bode with the apporpate words
'//cprin 7Aug2006: added next function as a template for a function with error handling
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub MyTemplate()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: Template
' Description: Template Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006  Perfection Tools Software
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
    SysError Err, "Templateunique number"

GoTo TemplateExit

End Sub
'//pt 9Jun2007: added next function to use rather than PathFileExists API function because it fails
Public Function FileExists(FullFileName As String) As Boolean

    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function

